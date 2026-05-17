# PDF Intelligence — Groq AI PDF Categorizer

A web app that reads any PDF, uses Groq AI (LLaMA 3.3 70B) to analyze and categorize the content, and exports it to a structured, formatted Excel workbook.

## Features
- **AI-powered analysis**: Extracts document type, entities, key facts, topics, sentiment, and more
- **4-sheet Excel output**: Summary, Key Entities, Key Facts, and Raw Text by page
- **Fast**: Groq's LPU inference means analysis in seconds
- **Privacy-first**: Your API key is never stored; PDFs are deleted after processing

## Setup

### 1. Install dependencies
```bash
pip install -r requirements.txt
```

### 2. Get a Groq API key (free)
Visit https://console.groq.com and create a free account.

### 3. Run the app
```bash
python app.py
```

### 4. Open in browser
Navigate to: http://localhost:5050

## Usage
1. Paste your Groq API key into the field
2. Upload a PDF (drag & drop or browse)
3. Click "Analyze with Groq AI"
4. Download the Excel report when processing completes

## Excel Output Sheets
| Sheet | Contents |
|-------|----------|
| 📋 Summary | Document type, language, sentiment, topics, action items |
| 👤 Entities | People, organizations, locations, dates, amounts found |
| 📊 Key Facts | Specific facts and data points with categories and page numbers |
| 📄 Raw Text | Full extracted text organized by page number |

## Technical Stack
- **Backend**: Python + Flask
- **PDF Reading**: pdfplumber
- **AI Model**: Groq API → LLaMA 3.3 70B Versatile
- **Excel Export**: openpyxl
- **Frontend**: Vanilla HTML/CSS/JS (no framework needed)
