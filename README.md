# PDF to Excel Converter

A quick tool to extract data from PDFs and convert to Excel using Google Gemini AI.

## What it does

Takes a PDF file, uses AI to extract the data, and gives you a clean Excel file with everything organized.

## Setup

### You'll need:
- Python
- Google Gemini API key (get it free from https://aistudio.google.com)

### Install

```bash
pip install -r requirements.txt
```

### Configure API Key

Make a `.env` file and add your key:

```
GEMINI_API_KEY=your_key_here
```

### Run it

```bash
streamlit run streamlit_app.py
```

Opens at `http://localhost:8501`

## How to use

1. Upload PDF
2. Click "Start Extraction"  
3. Check the table
4. Download Excel file

## Features

- Uses Gemini 2.5 Flash AI model
- Extracts data into Key-Value-Comments format
- Splits education, work experience into separate fields
- Exports to Excel with serial numbers

## Tech used

- Streamlit (web interface)
- Pandas (data handling)
- Google Gemini API (AI extraction)
- OpenPyXL (Excel export)

## Notes

- Don't commit the .env file (already in .gitignore)
- If extraction fails, check if PDF is text-based (not scanned image)
- Large files might take 30+ seconds

