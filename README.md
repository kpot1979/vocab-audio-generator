# Polyglot Audio & Video Lesson Generator

An automated Python pipeline that transforms Excel vocabulary sheets into synchronized audio lessons (`.mp3` with `.lrc` timestamps) and YouTube-ready video lessons (`.mp4`) using various Text-to-Speech (TTS) engines (ElevenLabs, Azure, Google TTS, Kokoro, and Piper).

## ✨ Features

* **Excel Integration:** Reads vocabulary, translations, and example phrases directly from Excel spreadsheets.

* **Precise Timestamp Sync (`.lrc`):** Measures exact physical audio durations and uses FFmpeg re-encoding (`libmp3lame`) to prevent audio drift and sync text highlighting perfectly.

* **Smart Caching (`script_with_cach.py`):** Caches target language audio files to save API limits and costs on paid TTS providers like ElevenLabs.

* **Fresh Generation (`script.py`):** Clears temporary directories and re-renders all audio elements from scratch upon every run.

* **Multi-Engine Support:** Works with ElevenLabs, Azure Speech, gTTS, Kokoro ONNX, and Piper.

* **Dynamic Pauses:** Automatically calculates intelligent pauses based on the duration of the preceding speech segment.

* **Video Generation:** Creates 1920x1080 video cards with clean text wrapping and smooth transitions for video platforms.

## 🛠️ Step-by-Step Installation & Setup

### Step 1. Install Python (3.11 Recommended)

1. Open PowerShell (`Win + R` -> type `powershell` -> Enter) or your terminal in VS Code (`Ctrl + ~`).

2. Python 3.11 guarantees stable compilation for binary packages. Install it via terminal:

   ```
   winget install Python.Python.3.11
   
   ```

3. *Restart your terminal after installation*, then verify it:

   ```
   python --version
   
   ```

### Step 2. Install FFmpeg

FFmpeg is required for audio/video processing and assembly. Install it via terminal:

```
winget install Gyan.FFmpeg

```

*Restart your terminal and check availability:*

```
ffmpeg -version

```

### Step 3. Install Python Libraries

Run the following command in your project directory to install all required dependencies:

```
pip install pandas gTTS pydub Pillow openpyxl mutagen requests kokoro-onnx soundfile

```

## ⚙️ Configuration & Project Structure

1. **API Keys (`config.py`):**
   If you are using paid or cloud-based TTS engines (like ElevenLabs or Azure), create a `config.py` file in the root directory:

   ```
   AZURE_SPEECH_KEY = "your_azure_key"
   ELEVENLABS_API_KEY = "your_elevenlabs_key"
   
   ```

2. **Excel Vocabulary File (`vocabulary.xlsx`):**
   Place your Excel file in the project folder. By default, the script reads:

   * **Column A (COL_LW):** Learn word

   * **Column B (COL_RW):** Native translation

   * **Column C (COL_DF):** Example phrase in target language

   * **Column D (COL_RF):** Phrase translation

3. **Audio Template (`TEMPLATE_STR`):**
   You can customize the playback structure at the top of the scripts:

   ```
   TEMPLATE_STR = (
       "learnWord + shortPause + nativeWord + shortPause + learnWord + shortPause + learnWord + dynamicPause + "
       "learnPhrase + shortPause + nativePhrase + shortPause + learnPhrase + shortPause + learnPhrase + dynamicPause"
   )
   
   ```

## 🚀 Usage (The Two Scripts)

This repository contains two working scripts with different caching logics:

### 1. `script.py` (Fresh Generation Mode)

Clears previous temporary audio folders and generates all audio files from scratch upon every run. Ideal when you want to fully re-render a lesson after tweaking voices or spreadsheet data.

```
python script.py

```

### 2. `script_with_cach.py` (Smart Caching Mode)

Checks the `audio/` directory. If target language audio files already exist for words/phrases, it reuses them without making redundant API requests, saving your API limits (especially useful for ElevenLabs). It also outputs a detailed summary of cached vs. newly generated items in the console.

```
python script_with_cach.py

```

## 📂 Output Files

Once the script finishes processing your selected Excel sheet, it generates:

1. **`<SheetName>.mp3`** — The complete audio lesson with embedded ID3 tags and lyrics (`USLT`).

2. **`<SheetName>.lrc`** — Synchronized lyric timestamps file.

3. **`<SheetName>.mp4`** — Full 1080p video lesson ready for YouTube.

4. **`audio/` folder** — Cached individual MP3 components for words and phrases.

## 📜 License

This project is open-source and available under the [MIT License](LICENSE).
