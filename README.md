# AI-Powered Document to Excel Converter

Convert unstructured PDF data into structured Excel files using Google Gemini AI.

## Features

- Extract data from PDF documents automatically
- Structure data into Key-Value-Comments format
- Export to Excel with serial numbers
- Powered by Google Gemini 2.5 Flash

## Setup Instructions

### 1. Clone the Repository

```bash
git clone <repository-url>
cd Turerz-Assignment
```

### 2. Install Dependencies

```bash
pip install -r requirements.txt
```

### 3. Configure Environment Variables

Create a `.env` file in the project root:

```bash
cp .env.example .env
```

Edit the `.env` file and add your Gemini API key:

```
GEMINI_API_KEY=your_actual_api_key_here
```

**Get your API key from:** https://aistudio.google.com

### 4. Run the Application

```bash
streamlit run streamlit_app.py
```

The application will open in your browser at `http://localhost:8501`

## Usage

1. Upload your PDF document using the file uploader
2. Click "Start Extraction"
3. Review the extracted data in the table
4. Download the structured Excel file

## Requirements

- Python 3.8+
- Google Gemini API Key
- Internet connection for API calls

## Project Structure

```
Turerz-Assignment/
├── streamlit_app.py      # Main application
├── requirements.txt      # Python dependencies
├── .env.example         # Environment variables template
├── .env                 # Your API keys (git-ignored)
└── README.md           # This file
```

## Security

- Never commit your `.env` file to version control
- The `.env` file is already included in `.gitignore`
- Keep your API keys secure and private

## License

MIT