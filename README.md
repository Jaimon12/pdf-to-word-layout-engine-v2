# PDF to Word Layout Engine V2

Editable coordinate-based OCR reconstruction. It does not embed full-page screenshots and does not append a second OCR transcript section.

## Install Python dependencies
```bash
python -m pip install -r requirements.txt
```

## Install OCR engine
Install Tesseract OCR for Windows separately. Include Malayalam (`mal`) and English (`eng`) trained data. Common executable:
`C:\Program Files\Tesseract-OCR\tesseract.exe`

Install a Unicode Malayalam font such as Noto Sans Malayalam in Windows.

## Run
```bash
streamlit run app.py
```

## Jupyter runner
```python
import subprocess, sys
APP = r"C:\Users\megha.sabu\Documents\Python Files Folder\PDF COVERTED\app.py"
process = subprocess.Popen([sys.executable, "-m", "streamlit", "run", APP, "--server.port", "8501"])
print("http://localhost:8501")
```

## Important
The app detects line bounding boxes and places editable Word paragraphs using their coordinates. It is an approximate layout reconstruction, not a guarantee of exact typography. It does not currently preserve non-text vector artwork or embedded images as editable Word objects. Inspect OCR text and compare the output with the PDF.


## Automatic Tesseract path detection (updated)
The app now checks, in order:
- The executable path entered in the sidebar
- `tesseract` on Windows PATH
- `C:\Program Files\Tesseract-OCR\tesseract.exe`
- `C:\Program Files (x86)\Tesseract-OCR\tesseract.exe`
- Common per-user installation locations

The app displays the exact detected executable in the sidebar.
