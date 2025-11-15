# GujaratiPDFTool

**Gujarati PDF Converter & Search — by SM TECHIE**

Convert non-searchable Gujarati election PDFs into searchable PDFs and search Gujarati names (મેમણ, સલીમ etc.) quickly using OCR.

---

## Features
- Convert folder of PDFs → searchable PDFs (Tesseract OCR, Gujarati + English)
- Fallback OCR per-page for scanned/encoded PDFs
- Fast fuzzy search across many PDFs
- Export results to CSV / Excel
- GUI with progress, dark mode and settings
- Works offline (Tesseract + Poppler required)

---

## Requirements
- Windows 10/11
- Python 3.9+ (if running from source)
- [Tesseract OCR] installed and `guj.traineddata` present in tessdata
- Poppler (pdftoppm available)
- Python packages:
