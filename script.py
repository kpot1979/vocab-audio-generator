"""Создает учебные MP3- и MP4-материалы из строк выбранного листа Excel.

Для каждой строки скрипт синтезирует слова и пример предложения на двух языках,
собирает их в последовательность по TEMPLATE_STR и создает карточку с текстом.
Для работы нужны ffmpeg, pandas, pydub, Pillow, mutagen и выбранные движки TTS (piper, gtts, kokoro, elevenlabs, azure).
"""

import os
import re
import shutil
import subprocess
import pandas as pd
from pydub import AudioSegment
from PIL import Image, ImageDraw, ImageFont

# Импортируем секретные ключи из отдельного файла config.py
try:
    from config import AZURE_SPEECH_KEY, ELEVENLABS_API_KEY
except ImportError:
    print("[ОШИБКА] Не найден файл 'config.py'! Создайте его в корне проекта и добавьте туда ключи.")
    exit()

# ==================== НАСТРОЙКИ КОНСТРУКТОРА И ПАРАМЕТРОВ ====================
EXCEL_FILENAME = "vocabulary.xlsx"

# --- 1. НАСТРОЙКА ЯЗЫКОВ И НЕЗАВИСИМЫХ ДВИЖКОВ ОЗВУЧКИ ---
# Выбор движка для ИЗУЧАЕМОГО языка ("piper", "gtts", "kokoro", "elevenlabs" или "azure")
TTS_ENGINE_LEARN = "gtts" 

# Выбор движка для РОДНОГО языка ("piper", "gtts", "kokoro", "elevenlabs" или "azure")
TTS_ENGINE_NATIVE = "gtts"

# Резервный движок, если основной движок вернет ошибку или тишину
TTS_FALLBACK_ENGINE = "piper"

LANG_LEARN = "de"  # Изучаемый язык (для слов и фраз)
LANG_NATIVE = "ru" # Родной язык (для переводов)

# --- НАСТРОЙКИ ДЛЯ AZURE TTS ---
AZURE_SPEECH_REGION = os.getenv("AZURE_SPEECH_REGION", "germanywestcentral")
AZURE_VOICE_LEARN = "de-DE-ConradNeural"
AZURE_VOICE_NATIVE = "ru-RU-SvetlanaNeural"
AZURE_SPEED = "0.85"  # Скорость (коэффициент)

# --- НАСТРОЙКИ ДЛЯ ELEVENLABS ---
ELEVENLABS_VOICE_LEARN = "MLFHn2hZ3zKifXrugl26"
ELEVENLABS_VOICE_NATIVE = "21m00Tcm4TlvDq8ikWAM"
ELEVENLABS_MODEL_ID = "eleven_multilingual_v2"
ELEVENLABS_SPEED = 0.7

# --- НАСТРОЙКИ ДЛЯ KOKORO ---
KOKORO_MODEL_PATH = "kokoro-v1.0.onnx"
KOKORO_VOICES_PATH = "voices-v1.0.bin"
KOKORO_VOICE_LEARN = "af_kore"
KOKORO_VOICE_NATIVE = "af_bella"

# --- НАСТРОЙКИ ДЛЯ gTTS ---
GTTS_TLD_LEARN = "de"
GTTS_TLD_NATIVE = "ru"
GTTS_SLOW_SPEED = False

# --- НАСТРОЙКИ ДЛЯ PIPER ---
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
PIPER_BASE_DIR = os.path.join(CURRENT_DIR, "piper")
MODEL_LEARN_FILENAME = "de_DE-thorsten-high.onnx"
MODEL_NATIVE_FILENAME = "ru_RU-irina-medium.onnx"

# --- 2. СООТВЕТСТВИЕ КОЛОНОК В EXCEL (1 = колонка A, 2 = B, 3 = C и т.д.) ---
COL_LW = 1  # Слово на изучаемом языке
COL_RW = 2  # Перевод слова на родном языке
COL_DF = 3  # Фраза (пример) на изучаемом языке (если нет колонки, поставьте 0)
COL_RF = 4  # Перевод фразы на родном языке (если нет колонки, поставьте 0)

# --- 3. НАСТРОЙКИ ТАЙМИНГОВ И ПАУЗ (в миллисекундах) ---
START_PAUSE_MS = 1000       
PAUSE_MULTIPLIER = 1.2      
MANUAL_EXTRA_MS = 1200      
SHORT_PAUSE_MS = 800        
TRANSITION_PAUSE_MS = 2000  

# --- 4. НАСТРОЙКИ АУДИО И КАЧЕСТВА ---
AUDIO_SAMPLE_RATE = 22050   
AUDIO_CHANNELS = 1          
MP3_BITRATE = "192k"        

# --- 5. ВИЗУАЛЬНЫЕ НАСТРОЙКИ ВИДЕО И КАРТОЧЕК ---
VIDEO_WIDTH = 1920
VIDEO_HEIGHT = 1080
VIDEO_FPS = 24

FONT_NAME_MAIN = "arial.ttf"
FONT_NAME_SUB = "arial.ttf"
FONT_SIZE_MAIN = 75
FONT_SIZE_SUB = 44

BG_COLOR = (18, 18, 24)               
COLOR_LEARN_WORD = (255, 255, 255)    
COLOR_NATIVE_WORD = (100, 220, 140)   
COLOR_SEPARATOR = (160, 160, 180)     
COLOR_LEARN_PHRASE = (220, 220, 230)  
COLOR_NATIVE_PHRASE = (140, 180, 220) 

TEXT_MAX_WIDTH = 1680                 
Y_MAIN_START = 280                    
Y_PHRASE_START = 620                  
LINE_SPACING_MAIN = 90                
LINE_SPACING_SUB = 55                 

# --- 6. ТЕГИ И МЕТАДАННЫЕ АУДИОФАЙЛА ---
META_ARTIST = f"Vocabulary Learning ({LANG_LEARN.upper()}-{LANG_NATIVE.upper()})"
META_ALBUM = "Language Lessons 2026"
META_GENRE = "Education / Video & Audio"
META_COMMENT = f"Created via Python (Engines: Learn={TTS_ENGINE_LEARN.upper()}, Native={TTS_ENGINE_NATIVE.upper()})"

# --- 7. ШАБЛОН (КОНСТРУКТОР) АУДИОБЛОКА ---
TEMPLATE_STR = (
    "learnWord + shortPause + nativeWord + shortPause + learnWord + shortPause + learnWord + dynamicPause + "
    "learnPhrase + shortPause + nativePhrase + shortPause + learnPhrase + shortPause + learnPhrase + dynamicPause"
)
# ============================================================================

EXCEL_FILE = os.path.join(CURRENT_DIR, EXCEL_FILENAME)

if not os.path.exists(EXCEL_FILE):
    print(f"Ошибка: Файл '{EXCEL_FILENAME}' не найден в папке скрипта ({CURRENT_DIR})!")
    exit()

excel_obj = pd.ExcelFile(EXCEL_FILE)
available_sheets = excel_obj.sheet_names

print("\n" + "="*50)
print(" ДОСТУПНЫЕ ЛИСТЫ В EXCEL-ФАЙЛЕ:")
print("="*50)
for idx, name in enumerate(available_sheets, 1):
    print(f"  [{idx}] {name}")
print("="*50)

user_input = input(
    f"Введите номер листа (1-{len(available_sheets)}) или его точное название"
    f" (по умолчанию 1): "
).strip()

selected_sheet = available_sheets[0]
if user_input == "":
    selected_sheet = available_sheets[0]
elif user_input.isdigit():
    idx = int(user_input) - 1
    if 0 <= idx < len(available_sheets):
        selected_sheet = available_sheets[idx]
else:
    if user_input in available_sheets:
        selected_sheet = user_input

safe_sheet_name = re.sub(r'[\\/*?:"<>|]', "", selected_sheet).strip()
AUDIO_OUTPUT_FILE = f"{safe_sheet_name}.mp3"
VIDEO_OUTPUT_FILE = f"{safe_sheet_name}.mp4"
LRC_OUTPUT_FILE = f"{safe_sheet_name}.lrc"

df = pd.read_excel(excel_obj, sheet_name=selected_sheet, header=None)

if df.empty or len(df.dropna(how="all")) == 0:
    print(f"\n[ВНИМАНИЕ] Лист '{selected_sheet}' пустой или не содержит данных!")
    exit()

# --- ИЕРАРХИЯ ПАПОК С СУФФИКСОМ МОДЕЛИ (ОЧИСТКА ПРИ КАЖДОМ ЗАПУСКЕ) ---
engine_suffix = f"_{TTS_ENGINE_LEARN.upper()}"
root_audio_dir = os.path.join(CURRENT_DIR, "audio")
words_output_dir = os.path.join(root_audio_dir, f"{safe_sheet_name}{engine_suffix}")
phrases_output_dir = os.path.join(words_output_dir, f"phrase{engine_suffix}")

if os.path.exists(words_output_dir):
    for filename in os.listdir(words_output_dir):
        file_path = os.path.join(words_output_dir, filename)
        if os.path.isfile(file_path):
            try: os.remove(file_path)
            except: pass
    print(f"[ИНФО] Папка для слов '{os.path.relpath(words_output_dir, CURRENT_DIR)}' очищена.")
else:
    os.makedirs(words_output_dir, exist_ok=True)
    print(f"[ИНФО] Создана папка для слов: '{os.path.relpath(words_output_dir, CURRENT_DIR)}'")

if os.path.exists(phrases_output_dir):
    for filename in os.listdir(phrases_output_dir):
        file_path = os.path.join(phrases_output_dir, filename)
        if os.path.isfile(file_path):
            try: os.remove(file_path)
            except: pass
    print(f"[ИНФО] Папка для фраз '{os.path.relpath(phrases_output_dir, CURRENT_DIR)}' очищена.")
else:
    os.makedirs(phrases_output_dir, exist_ok=True)
    print(f"[ИНФО] Создана папка для фраз: '{os.path.relpath(phrases_output_dir, CURRENT_DIR)}'")


class AudioNode:
    def __init__(self, filepath, duration_ms):
        self.filepath = filepath
        self.duration_ms = duration_ms

    def __add__(self, other):
        if isinstance(other, AudioNode):
            return Playlist([self, other])
        elif isinstance(other, DynamicPauseProxy):
            pause_node = other.create_pause_node(self.duration_ms)
            return Playlist([self, pause_node])
        return self

    def __radd__(self, other):
        if isinstance(other, AudioNode):
            return Playlist([other, self])
        return self


class Playlist:
    def __init__(self, items):
        self.items = items

    def __add__(self, other):
        if isinstance(other, AudioNode):
            return Playlist(self.items + [other])
        elif isinstance(other, Playlist):
            return Playlist(self.items + other.items)
        elif isinstance(other, DynamicPauseProxy):
            last_node = None
            for item in reversed(self.items):
                if isinstance(item, AudioNode):
                    last_node = item
                    break
            if last_node:
                pause_node = other.create_pause_node(last_node.duration_ms)
                return Playlist(self.items + [pause_node])
            return self
        return self

    def __radd__(self, other):
        if isinstance(other, AudioNode):
            return Playlist([other] + self.items)
        return self

    def get_files(self):
        files = []
        for item in self.items:
            if isinstance(item, Playlist):
                files.extend(item.get_files())
            elif isinstance(item, AudioNode):
                files.append(item.filepath)
        return files


class DynamicPauseProxy:
    def __init__(self, multiplier, extra_ms, temp_files_list):
        self.multiplier = multiplier
        self.extra_ms = extra_ms
        self.temp_files = temp_files_list

    def create_pause_node(self, base_duration_ms):
        pause_ms = int(base_duration_ms * self.multiplier + self.extra_ms)
        silent_seg = AudioSegment.silent(duration=pause_ms, frame_rate=AUDIO_SAMPLE_RATE)
        temp_pause = f"temp_pause_{os.urandom(4).hex()}.mp3"
        silent_seg.export(
            temp_pause,
            format="mp3",
            bitrate=MP3_BITRATE,
            parameters=["-ar", str(AUDIO_SAMPLE_RATE), "-ac", str(AUDIO_CHANNELS)],
        )
        self.temp_files.append(temp_pause)
        return AudioNode(temp_pause, pause_ms)

    def __radd__(self, other):
        if isinstance(other, AudioNode):
            return Playlist([other, self.create_pause_node(other.duration_ms)])
        return NotImplemented

    def __add__(self, other):
        return self


# Инициализация Kokoro
kokoro_instance = None
if "kokoro" in TTS_ENGINE_LEARN.lower() or "kokoro" in TTS_ENGINE_NATIVE.lower() or "kokoro" in TTS_FALLBACK_ENGINE.lower():
    try:
        from kokoro_onnx import Kokoro
        if os.path.exists(KOKORO_MODEL_PATH) and os.path.exists(KOKORO_VOICES_PATH):
            kokoro_instance = Kokoro(KOKORO_MODEL_PATH, KOKORO_VOICES_PATH)
    except ImportError:
        pass


def generate_tts_for_engine(text, lang, is_learn_lang, engine_name, temp_files_list):
    temp_wav = f"temp_word_{os.urandom(4).hex()}.wav"
    temp_mp3 = f"temp_word_{os.urandom(4).hex()}.mp3"
    engine = engine_name.lower()

    if engine == "azure":
        import azure.cognitiveservices.speech as speechsdk
        speech_config = speechsdk.SpeechConfig(subscription=AZURE_SPEECH_KEY, region=AZURE_SPEECH_REGION)
        voice_name = AZURE_VOICE_LEARN if is_learn_lang else AZURE_VOICE_NATIVE
        speech_config.speech_synthesis_voice_name = voice_name
        speech_config.set_speech_synthesis_output_format(speechsdk.SpeechSynthesisOutputFormat.Audio16Khz32KBitRateMonoMp3)

        audio_config = speechsdk.audio.AudioOutputConfig(filename=temp_mp3)
        synthesizer = speechsdk.SpeechSynthesizer(speech_config=speech_config, audio_config=audio_config)

        try:
            speed_float = float(AZURE_SPEED)
            rate_percent = int((speed_float - 1.0) * 100)
            rate_str = f"{rate_percent:+d}%"
        except:
            rate_str = "0%"

        ssml_text = f"""
        <speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xml:lang='{lang}'>
            <voice name='{voice_name}'>
                <prosody rate='{rate_str}'>{text}</prosody>
            </voice>
        </speak>
        """
        result = synthesizer.speak_ssml_async(ssml_text).get()
        if result.reason != speechsdk.ResultReason.SynthesizingAudioCompleted:
            raise RuntimeError("Azure TTS Error")
        audio = AudioSegment.from_mp3(temp_mp3)

    elif engine == "elevenlabs":
        import requests
        voice_id = ELEVENLABS_VOICE_LEARN if is_learn_lang else ELEVENLABS_VOICE_NATIVE
        url = f"https://api.elevenlabs.io/v1/text-to-speech/{voice_id}"
        headers = {"Accept": "audio/mpeg", "Content-Type": "application/json", "xi-api-key": ELEVENLABS_API_KEY}
        payload = {"text": text, "model_id": ELEVENLABS_MODEL_ID, "voice_settings": {"stability": 0.5, "similarity_boost": 0.75, "speed": ELEVENLABS_SPEED}}
        
        response = requests.post(url, json=payload, headers=headers)
        if response.status_code != 200:
            raise RuntimeError(f"ElevenLabs API Error [{response.status_code}]")
        with open(temp_mp3, "wb") as f: f.write(response.content)
        audio = AudioSegment.from_mp3(temp_mp3)

    elif engine == "kokoro":
        if not kokoro_instance: raise RuntimeError("Kokoro не инициализирован")
        voice = KOKORO_VOICE_LEARN if is_learn_lang else KOKORO_VOICE_NATIVE
        data_path = os.path.join(CURRENT_DIR, "espeak-ng-data")
        import ctypes
        buf = ctypes.create_unicode_buffer(260)
        ctypes.windll.kernel32.GetShortPathNameW(data_path, buf, 260)
        os.environ["ESPEAK_DATA_PATH"] = buf.value if buf.value else data_path

        samples, sample_rate = kokoro_instance.create(text, voice=voice, speed=1.0, lang=lang)
        import soundfile as sf
        sf.write(temp_wav, samples, sample_rate)
        audio = AudioSegment.from_wav(temp_wav)
        if os.path.exists(temp_wav): os.remove(temp_wav)

    elif engine == "piper":
        model_filename = MODEL_LEARN_FILENAME if is_learn_lang else MODEL_NATIVE_FILENAME
        piper_exe = os.path.join(PIPER_BASE_DIR, "piper.exe")
        model_path = os.path.join(PIPER_BASE_DIR, model_filename)
        wav_abs_path = os.path.join(CURRENT_DIR, temp_wav)
        espeak_data_dir = os.path.join(PIPER_BASE_DIR, "espeak-ng-data")

        cmd = [piper_exe, "--model", model_filename, "--output_file", wav_abs_path]
        env = os.environ.copy()
        env["ESPEAK_DATA_PATH"] = espeak_data_dir
        subprocess.run(cmd, input=text.encode("utf-8"), cwd=PIPER_BASE_DIR, env=env, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        audio = AudioSegment.from_wav(wav_abs_path)
        if os.path.exists(wav_abs_path): os.remove(wav_abs_path)

    elif engine == "gtts":
        from gtts import gTTS
        tld = GTTS_TLD_LEARN if is_learn_lang else GTTS_TLD_NATIVE
        tts = gTTS(text=text, lang=lang, tld=tld, slow=GTTS_SLOW_SPEED)
        tts.save(temp_mp3)
        audio = AudioSegment.from_mp3(temp_mp3)
    else:
        raise ValueError(f"Неизвестный движок: {engine}")

    audio = audio.set_frame_rate(AUDIO_SAMPLE_RATE).set_channels(AUDIO_CHANNELS)
    if audio.dBFS == float('-inf') or audio.dBFS < -50:
        raise ValueError("Тишина")
    audio.export(temp_mp3, format="mp3", bitrate=MP3_BITRATE, parameters=["-ar", str(AUDIO_SAMPLE_RATE), "-ac", str(AUDIO_CHANNELS)])
    temp_files_list.append(temp_mp3)
    return AudioNode(temp_mp3, len(audio))


def generate_tts_file(text, lang, is_learn_lang, temp_files_list):
    if not text or not text.strip():
        return create_silent_node(500, temp_files_list)

    primary_engine = TTS_ENGINE_LEARN if is_learn_lang else TTS_ENGINE_NATIVE

    try:
        return generate_tts_for_engine(text, lang, is_learn_lang, primary_engine, temp_files_list)
    except Exception as e:
        print(f"\n[ПРЕДУПРЕЖДЕНИЕ] Ошибка основного движка ({primary_engine}) для текста '{text}': {e}")
        if primary_engine.lower() != TTS_FALLBACK_ENGINE.lower():
            try:
                return generate_tts_for_engine(text, lang, is_learn_lang, TTS_FALLBACK_ENGINE, temp_files_list)
            except Exception:
                pass
        return create_silent_node(1000, temp_files_list)


def create_silent_node(duration_ms, temp_files_list):
    silent_seg = AudioSegment.silent(duration=duration_ms, frame_rate=AUDIO_SAMPLE_RATE)
    temp_path = f"temp_silent_{os.urandom(4).hex()}.mp3"
    silent_seg.export(temp_path, format="mp3", bitrate=MP3_BITRATE, parameters=["-ar", str(AUDIO_SAMPLE_RATE), "-ac", str(AUDIO_CHANNELS)])
    temp_files_list.append(temp_path)
    return AudioNode(temp_path, duration_ms)


def sanitize_filename(name):
    return re.sub(r'[\\/*?:"<>|]', "", name).strip()


def wrap_text_to_lines(text, font, max_width, draw):
    words = text.split()
    if not words: return [""]
    lines, current_line = [], []
    for word in words:
        test_line = " ".join(current_line + [word])
        try:
            bbox = draw.textbbox((0, 0), test_line, font=font)
            w = bbox[2] - bbox[0]
        except:
            w = len(test_line) * 20
        if w <= max_width: current_line.append(word)
        else:
            if current_line: lines.append(" ".join(current_line)); current_line = [word]
            else: lines.append(word)
    if current_line: lines.append(" ".join(current_line))
    return lines


def create_card_image(word_learn, word_native, phrase_learn, phrase_native, output_path):
    img = Image.new("RGB", (VIDEO_WIDTH, VIDEO_HEIGHT), color=BG_COLOR)
    draw = ImageDraw.Draw(img)
    try:
        font_main = ImageFont.truetype(FONT_NAME_MAIN, FONT_SIZE_MAIN)
        font_sub = ImageFont.truetype(FONT_NAME_SUB, FONT_SIZE_SUB)
    except:
        font_main = ImageFont.load_default()
        font_sub = ImageFont.load_default()

    max_w = TEXT_MAX_WIDTH
    if word_learn and word_native:
        main_text = f"{word_learn}   —   {word_native}"
        bbox_main = draw.textbbox((0, 0), main_text, font=font_main)
        w_main = bbox_main[2] - bbox_main[0]
        y_main = Y_MAIN_START
        if w_main <= max_w:
            bbox_learn = draw.textbbox((0, 0), word_learn, font=font_main)
            w_learn = bbox_learn[2] - bbox_learn[0]
            bbox_sep = draw.textbbox((0, 0), "   —   ", font=font_main)
            w_sep = bbox_sep[2] - bbox_sep[0]
            bbox_native = draw.textbbox((0, 0), word_native, font=font_main)
            w_native = bbox_native[2] - bbox_native[0]
            total_w = w_learn + w_sep + w_native
            x_start = (VIDEO_WIDTH - total_w) // 2
            draw.text((x_start, y_main), word_learn, font=font_main, fill=COLOR_LEARN_WORD)
            draw.text((x_start + w_learn, y_main), "   —   ", font=font_main, fill=COLOR_SEPARATOR)
            draw.text((x_start + w_learn + w_sep, y_main), word_native, font=font_main, fill=COLOR_NATIVE_WORD)
        else:
            lines_learn = wrap_text_to_lines(word_learn, font_main, max_w, draw)
            for i, line in enumerate(lines_learn):
                bbox_l = draw.textbbox((0, 0), line, font=font_main)
                lw = bbox_l[2] - bbox_l[0]
                draw.text(((VIDEO_WIDTH - lw) // 2, y_main + i * LINE_SPACING_MAIN), line, font=font_main, fill=COLOR_LEARN_WORD)
            y_main += len(lines_learn) * LINE_SPACING_MAIN + 15
            lines_native = wrap_text_to_lines(word_native, font_main, max_w, draw)
            for i, line in enumerate(lines_native):
                bbox_l = draw.textbbox((0, 0), line, font=font_main)
                lw = bbox_l[2] - bbox_l[0]
                draw.text(((VIDEO_WIDTH - lw) // 2, y_main + i * LINE_SPACING_MAIN), line, font=font_main, fill=COLOR_NATIVE_WORD)

    if phrase_learn or phrase_native:
        y_phrase = Y_PHRASE_START
        if phrase_learn:
            phrase_learn_lines = wrap_text_to_lines(phrase_learn, font_sub, max_w, draw)
            for i, line in enumerate(phrase_learn_lines):
                bbox_l = draw.textbbox((0, 0), line, font=font_sub)
                lw = bbox_l[2] - bbox_l[0]
                draw.text(((VIDEO_WIDTH - lw) // 2, y_phrase + i * LINE_SPACING_SUB), line, font=font_sub, fill=COLOR_LEARN_PHRASE)
            y_phrase += len(phrase_learn_lines) * LINE_SPACING_SUB + 25
        if phrase_native:
            phrase_native_lines = wrap_text_to_lines(phrase_native, font_sub, max_w, draw)
            for i, line in enumerate(phrase_native_lines):
                bbox_l = draw.textbbox((0, 0), line, font=font_sub)
                lw = bbox_l[2] - bbox_l[0]
                draw.text(((VIDEO_WIDTH - lw) // 2, y_phrase + i * LINE_SPACING_SUB), line, font=font_sub, fill=COLOR_NATIVE_PHRASE)

    img.save(output_path)


def ms_to_lrc_timestamp(ms):
    minutes = ms // 60000
    seconds = (ms // 1000) % 60
    centiseconds = (ms % 1000) // 10
    return f"[{minutes:02d}:{seconds:02d}.{centiseconds:02d}]"


def main():
    temp_files = []
    all_audio_files = []
    row_videos = []
    lrc_lines = []

    current_time_ms = 0

    if START_PAUSE_MS > 0:
        start_seg = AudioSegment.silent(duration=START_PAUSE_MS, frame_rate=AUDIO_SAMPLE_RATE)
        start_path = f"temp_start_pause_{os.urandom(4).hex()}.mp3"
        start_seg.export(start_path, format="mp3", bitrate=MP3_BITRATE, parameters=["-ar", str(AUDIO_SAMPLE_RATE), "-ac", str(AUDIO_CHANNELS)])
        temp_files.append(start_path)
        all_audio_files.append(start_path)
        current_time_ms += len(start_seg)

    transition_seg = AudioSegment.silent(duration=TRANSITION_PAUSE_MS, frame_rate=AUDIO_SAMPLE_RATE)
    transition_path = f"temp_transition_{os.urandom(4).hex()}.mp3"
    transition_seg.export(transition_path, format="mp3", bitrate=MP3_BITRATE, parameters=["-ar", str(AUDIO_SAMPLE_RATE), "-ac", str(AUDIO_CHANNELS)])
    temp_files.append(transition_path)

    print(f"\nНайдено строк в листе '{selected_sheet}': {len(df)}")
    print(f"Изучаемый язык: {LANG_LEARN.upper()} ({TTS_ENGINE_LEARN.upper()}) | Родной язык: {LANG_NATIVE.upper()} ({TTS_ENGINE_NATIVE.upper()})")
    print("Генерация аудио и видеокадров для строк...")

    for index, row in df.iterrows():
        w_learn_raw = row.iloc[COL_LW - 1] if len(row) >= COL_LW else None
        w_native_raw = row.iloc[COL_RW - 1] if len(row) >= COL_RW else None
        p_learn_raw = row.iloc[COL_DF - 1] if COL_DF > 0 and len(row) >= COL_DF else None
        p_native_raw = row.iloc[COL_RF - 1] if COL_RF > 0 and len(row) >= COL_RF else None

        word_learn_text = str(w_learn_raw).strip() if pd.notna(w_learn_raw) else ""
        word_native_text = str(w_native_raw).strip() if pd.notna(w_native_raw) else ""
        phrase_learn_text = str(p_learn_raw).strip() if pd.notna(p_learn_raw) else ""
        phrase_native_text = str(p_native_raw).strip() if pd.notna(p_native_raw) else ""

        if not word_learn_text and not word_native_text and not phrase_learn_text:
            print(f"[{index+1}/{len(df)}] [ПРОПУСК] Строка пуста или не содержит данных.")
            continue

        # ВЫВОД ИНФОРМАЦИИ О ТЕКУЩЕМ СЛОВЕ В КОНСОЛЬ
        print(f"\n[{index+1}/{len(df)}] 🔍 Обрабатывается слово: «{word_learn_text}» — «{word_native_text}»")

        # Фиксируем точный таймкод начала текущей строки без дрейфа
        card_start_timestamp = ms_to_lrc_timestamp(current_time_ms)
        lrc_lines.append(f"{card_start_timestamp} {word_learn_text} — {word_native_text}")

        safe_word_name = sanitize_filename(word_learn_text) if word_learn_text else ""

        if word_learn_text:
            print(f"   ├─ [СЛОВО] Генерируется через {TTS_ENGINE_LEARN.upper()}: «{word_learn_text}»")
            learnWord = generate_tts_file(word_learn_text, LANG_LEARN, True, temp_files)
            if safe_word_name:
                dest_word_path = os.path.join(words_output_dir, f"{safe_word_name}.mp3")
                try:
                    shutil.copy(learnWord.filepath, dest_word_path)
                except Exception as e:
                    print(f"   │  [ПРЕДУПРЕЖДЕНИЕ] Не удалось сохранить файл слова: {e}")
        else:
            learnWord = create_silent_node(500, temp_files)

        nativeWord = generate_tts_file(word_native_text, LANG_NATIVE, False, temp_files) if word_native_text else create_silent_node(500, temp_files)
        
        if phrase_learn_text:
            print(f"   ├─ [ФРАЗА] Генерируется через {TTS_ENGINE_LEARN.upper()}: «{phrase_learn_text}»")
            learnPhrase = generate_tts_file(phrase_learn_text, LANG_LEARN, True, temp_files)
            if safe_word_name:
                dest_phrase_path = os.path.join(phrases_output_dir, f"{safe_word_name}.mp3")
                try:
                    shutil.copy(learnPhrase.filepath, dest_phrase_path)
                except Exception as e:
                    print(f"   │  [ПРЕДУПРЕЖДЕНИЕ] Не удалось сохранить файл фразы: {e}")
        else:
            learnPhrase = create_silent_node(500, temp_files)

        nativePhrase = generate_tts_file(phrase_native_text, LANG_NATIVE, False, temp_files) if phrase_native_text else create_silent_node(500, temp_files)

        shortPause = create_silent_node(SHORT_PAUSE_MS, temp_files)
        dynamicPause = DynamicPauseProxy(PAUSE_MULTIPLIER, MANUAL_EXTRA_MS, temp_files)

        local_vars = {
            "learnWord": learnWord,
            "nativeWord": nativeWord,
            "learnPhrase": learnPhrase,
            "nativePhrase": nativePhrase,
            "shortPause": shortPause,
            "dynamicPause": dynamicPause,
        }

        try:
            row_playlist = eval(TEMPLATE_STR, {"__builtins__": None}, local_vars)
        except Exception as e:
            print(f"Ошибка в шаблоне на строке {index+1}: {e}")
            break

        row_audio_files = []
        if isinstance(row_playlist, Playlist):
            row_audio_files.extend(row_playlist.get_files())
        elif isinstance(row_playlist, AudioNode):
            row_audio_files.append(row_playlist.filepath)

        all_audio_files.extend(row_audio_files)
        all_audio_files.append(transition_path)

        row_temp_mp3 = f"temp_row_audio_{index}_{os.urandom(4).hex()}.mp3"
        temp_files.append(row_temp_mp3)

        sub_concat = f"sub_concat_{index}.txt"
        with open(sub_concat, "w", encoding="utf-8") as f:
            for af in row_audio_files:
                f.write(f"file '{os.path.abspath(af).replace(os.sep, '/')}'\n")
        temp_files.append(sub_concat)

        subprocess.run([
            "ffmpeg", "-y", "-f", "concat", "-safe", "0", "-i", sub_concat,
            "-c:a", "libmp3lame", "-b:a", MP3_BITRATE, row_temp_mp3
        ], stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        # Считываем реальную физическую длину собранного аудиофайла строки
        row_audio_segment = AudioSegment.from_mp3(row_temp_mp3)
        row_duration_ms = len(row_audio_segment)
        row_duration_sec = row_duration_ms / 1000.0

        # Накапливаем точное время с учетом длительности строки и паузы перехода
        current_time_ms += row_duration_ms
        current_time_ms += len(transition_seg)

        card_img_path = f"card_{index}_{os.urandom(4).hex()}.png"
        temp_files.append(card_img_path)
        create_card_image(word_learn_text, word_native_text, phrase_learn_text, phrase_native_text, card_img_path)

        row_video_path = f"row_video_{index}_{os.urandom(4).hex()}.mp4"
        temp_files.append(row_video_path)

        cmd_row_video = [
            "ffmpeg", "-y",
            "-loop", "1",
            "-i", card_img_path,
            "-i", row_temp_mp3,
            "-c:v", "libx264",
            "-tune", "stillimage",
            "-c:a", "aac",
            "-b:a", "192k",
            "-pix_fmt", "yuv420p",
            "-t", str(row_duration_sec),
            row_video_path
        ]
        subprocess.run(cmd_row_video, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        row_videos.append(row_video_path)

        trans_video_path = f"trans_video_{index}_{os.urandom(4).hex()}.mp4"
        temp_files.append(trans_video_path)
        trans_duration_sec = TRANSITION_PAUSE_MS / 1000.0

        cmd_trans_video = [
            "ffmpeg", "-y",
            "-loop", "1",
            "-i", card_img_path,
            "-i", transition_path,
            "-c:v", "libx264",
            "-tune", "stillimage",
            "-c:a", "aac",
            "-b:a", "192k",
            "-pix_fmt", "yuv420p",
            "-t", str(trans_duration_sec),
            trans_video_path
        ]
        subprocess.run(cmd_trans_video, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        row_videos.append(trans_video_path)

    if not row_videos:
        print("\n[ОШИБКА] Не удалось создать ни одного видеосегмента.")
        return

    # Сохраняем синхронизированный .lrc файл
    meta_lyrics_str = "\n".join(lrc_lines)
    with open(os.path.join(CURRENT_DIR, LRC_OUTPUT_FILE), "w", encoding="utf-8") as lrc_f:
        lrc_f.write(meta_lyrics_str)
    print(f"\n[ИНФО] Создан точный синхронизированный файл таймкодов: {LRC_OUTPUT_FILE}")

    # ==================== ФИНАЛЬНАЯ СБОРКА АУДИО (С РЕКОДИРОВКОЙ) ====================
    print("\nСборка итогового аудиофайла (.mp3) с перекодировкой...")
    master_concat_audio = "master_concat_audio.txt"
    with open(master_concat_audio, "w", encoding="utf-8") as f:
        for file_path in all_audio_files:
            f.write(f"file '{os.path.abspath(file_path).replace(os.sep, '/')}'\n")
    temp_files.append(master_concat_audio)

    audio_abs = os.path.abspath(AUDIO_OUTPUT_FILE)
    
    cmd_audio = [
        "ffmpeg", "-y", "-f", "concat", "-safe", "0", "-i", master_concat_audio,
        "-c:a", "libmp3lame", "-b:a", MP3_BITRATE, "-ar", str(AUDIO_SAMPLE_RATE),
        audio_abs
    ]
    subprocess.run(cmd_audio, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

    # --- ЗАПИСЬ МЕТАДАННЫХ ЧЕРЕЗ MUTAGEN (ID3-фрейм USLT) ---
    try:
        from mutagen.id3 import ID3, TIT2, TPE1, TALB, TCON, COMM, USLT, ID3NoHeaderError

        try:
            audio_tags = ID3(audio_abs)
        except ID3NoHeaderError:
            audio_tags = ID3()

        audio_tags["TIT2"] = TIT2(encoding=3, text=selected_sheet)
        audio_tags["TPE1"] = TPE1(encoding=3, text=META_ARTIST)
        audio_tags["TALB"] = TALB(encoding=3, text=META_ALBUM)
        audio_tags["TCON"] = TCON(encoding=3, text=META_GENRE)
        audio_tags["COMM"] = COMM(encoding=3, lang='eng', desc='comments', text=META_COMMENT)
        audio_tags["USLT"] = USLT(encoding=3, lang='deu', desc='', text=meta_lyrics_str)

        audio_tags.save(audio_abs, v2_version=3)
        print("[ИНФО] Теги с текстом успешно записаны через Mutagen в формат USLT.")
    except ImportError:
        print("[ПРЕДУПРЕЖДЕНИЕ] Библиотека 'mutagen' не установлена! Выполните: python -m pip install mutagen")
    except Exception as e:
        print(f"[ПРЕДУПРЕЖДЕНИЕ] Не удалось записать теги через Mutagen: {e}")

    # ==================== ФИНАЛЬНАЯ СБОРКА ВИДЕО ====================
    print("Сборка итогового видеофайла для YouTube (.mp4)...")
    master_concat_video = "master_concat_video.txt"
    with open(master_concat_video, "w", encoding="utf-8") as f:
        for vpath in row_videos:
            f.write(f"file '{os.path.abspath(vpath).replace(os.sep, '/')}'\n")
    temp_files.append(master_concat_video)

    video_abs = os.path.abspath(VIDEO_OUTPUT_FILE)
    cmd_video = [
        "ffmpeg", "-y", "-f", "concat", "-safe", "0", "-i", master_concat_video,
        "-c", "copy",
        video_abs
    ]
    subprocess.run(cmd_video, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

    for tf in temp_files:
        if os.path.exists(tf):
            try:
                os.remove(tf)
            except:
                pass

    print(f"\nУспех! Готовые файлы созданы:")
    print(f" 1. Аудио: {audio_abs} (и точный .lrc файл таймкодов)")
    print(f" 2. Видео: {video_abs}")
    print(f" 3. Папка с аудио слов: {words_output_dir}")
    print(f" 4. Папка с аудио фраз: {phrases_output_dir}")


if __name__ == "__main__":
    main()