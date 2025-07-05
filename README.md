# PDF Network Metrics Extractor

## Overview  
This Python script extracts network performance metrics (Packet Loss, Latency, Jitter) from page 3 of multiple PDF reports in a selected folder. It uses **pdfplumber** for text extraction and falls back to **Tesseract OCR** if needed. Extracted data is mapped into specified columns of an Excel sheet named **"Connectivity"**, with timestamps added for each entry. The script supports two data sources: **CMBO** and **PCWK**, each with different Excel column mappings.

## Features  
- Select PDF reports folder via GUI  
- Select Excel file to update via GUI  
- Extracts data from page 3 of PDFs  
- Uses OCR fallback for scanned or image-based PDFs  
- Writes results into specific Excel columns based on source  
- Adds current date-time stamps for each row  
- Saves updated Excel workbook  

## Requirements  
- Python 3.x  
- Libraries: `pdfplumber`, `pillow`, `openpyxl`, `pytesseract`, `numpy`, `opencv-python`, `tkinter` (usually included with Python)  

## Setup  
1. Install dependencies:  
   ```bash
   pip install pdfplumber pillow openpyxl pytesseract numpy opencv-python
