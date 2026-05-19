# PDF → Excel Extractor

A Flask web app that uses **Google Gemini AI** to extract structured data from PDFs and text, and export it as a formatted `.xlsx` spreadsheet.

## Features

- 📄 Upload any PDF (up to 32MB)
- 📝 Paste raw text or table data
- 🧠 Gemini AI structures the data intelligently
- 📊 Multi-sheet Excel output with professional formatting
- 💡 Optional custom extraction instructions

## Setup

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Run the app

```bash
python app.py
```

Open `http://localhost:5000` in your browser.

### 3. Get a Gemini API key

1. Go to https://aistudio.google.com/app/apikey
2. Create a free API key
3. Paste it into the web UI on each use (it's never stored)

## Usage

1. Enter your Gemini API key in the configuration field
2. Upload a PDF **and/or** paste text into the text area
3. Optionally add extraction instructions (e.g. "group by category", "only extract prices")
4. Click **Extract to Excel**
5. Your `.xlsx` file downloads automatically

## Project Structure

```
pdf2excel/
├── app.py              # Flask backend
├── requirements.txt    # Python dependencies
├── templates/
│   └── index.html      # Frontend UI
├── uploads/            # Temp PDF storage (auto-cleaned)
└── outputs/            # Generated Excel files
```

## Notes

- The Gemini API key entered in the UI is used per-request and not persisted to disk
- Uploaded PDFs are deleted after text extraction
- Generated Excel files are stored in `outputs/` — you can periodically clean this folder
