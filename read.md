# PDF to PPT Converter

This project provides a user-friendly GUI application to convert PDF documents into PowerPoint (PPT) presentations. The application is built using Python, leveraging libraries such as `tkinter`, `pdf2image`, and `python-pptx`.

## Features

- Select a PDF file to convert.
- Specify the output location for the generated PPT file.
- Converts each page of the PDF into a slide in the PPT.
- Displays a progress bar and animation during conversion.
- Handles errors and provides user-friendly messages.

## Requirements

- Python 3.x
- Required Python libraries:
  - `tkinter` (comes pre-installed with Python)
  - `pdf2image`
  - `python-pptx`
- **Poppler**: Required for PDF to image conversion. Ensure Poppler is installed and added to your system PATH.

### Installing Poppler (Windows)
1. Download Poppler for Windows.
2. Extract the files to `C:\Program Files\poppler`.
3. Add `C:\Program Files\poppler\bin` to your system PATH.

### Installing Required Libraries
Run the following command to install the required Python libraries:
```bash
pip install pdf2image python-pptx
