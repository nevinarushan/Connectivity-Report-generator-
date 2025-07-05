PDF Network Metrics Extractor
Overview
This Python script extracts network performance metrics (Packet Loss, Latency, Jitter) from page 3 of multiple PDF reports in a selected folder. It uses pdfplumber for text extraction and falls back to Tesseract OCR if needed. Extracted data is mapped into specified columns of an Excel sheet named "Connectivity", with timestamps added for each entry. The script supports two data sources: CMBO and PCWK, each with different Excel column mappings.

Features
Select PDF reports folder via GUI

Select Excel file to update via GUI

Extracts data from page 3 of PDFs

Uses OCR fallback for scanned or image-based PDFs

Writes results into specific Excel columns based on source

Adds current date-time stamps for each row

Saves updated Excel workbook

Requirements
Python 3.x

Libraries: pdfplumber, pillow, openpyxl, pytesseract, numpy, opencv-python, tkinter (usually included with Python)

Setup
Install dependencies:

nginx
Copy
Edit
pip install pdfplumber pillow openpyxl pytesseract numpy opencv-python
Download and install Tesseract OCR.

Place the tesseract.exe inside a tesseract folder next to the script, or update the script’s Tesseract path accordingly.

Usage
Run the script.

Select the folder containing PDF reports when prompted.

Select the Excel file to update when prompted.

The script processes PDFs, extracts metrics, updates the Excel sheet, and saves the file.

Notes
Ensure the Excel file has a sheet named "Connectivity".

The script processes page 3 of each PDF (index 2). Adjust if needed.

Filenames determine source type (CMBO or PCWK) based on suffix.

Update column mappings in the script if your Excel structure differs.

