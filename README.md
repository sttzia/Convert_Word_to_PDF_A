# Convert_Word_to_PDF_A
Python tool to Convert a .docx to a .pdf, with PDF/A properties, including Navigation Bookmarks (if Heading Styles are applied)

# DOCX → PDF/A Converter (Python)

This project converts a DOCX to a **PDF/A‑2b** file and injects **navigation bookmarks** derived from Word heading outline levels.

## What it does

1. **Export DOCX → PDF** using Microsoft Word automation.
2. **Build bookmarks** from Word outline levels (Heading 1/2/3, etc.) and inject them into the PDF.
3. **Convert to PDF/A‑2b** using a proper XMP metadata setup.

## Requirements

- Windows with **Microsoft Word** installed
- Python 3.9+
- Ghostscript version 10.6.0.0 installed in Path "C:\Program Files\gs\gs10.06.0"
- Testing example folder is in "C:\temp\test_pdf_a_export"

## Install

```powershell
python -m venv .venv
\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Prepare your DOCX (important)

Bookmarks in the PDF are created from **Word outline levels**. Make sure your headings are true built‑in Word headings:

- Use **Heading 1**, **Heading 2**, **Heading 3** styles.
- If you created a custom style, it must have an **Outline Level** (1–9) to be detected.

## Usage

### Convert a DOCX to PDF/A with heading bookmarks

```powershell
python src\convert.py "C:\temp\test_pdf_a_export\test.docx" "C:\temp\test_pdf_a_export\test.pdf" --bookmarks headings
```

### Use a custom intermediate PDF path (optional)

```powershell
python src\convert.py "C:\temp\test_pdf_a_export\test.docx" "C:\temp\test_pdf_a_export\test.pdf" --intermediate-pdf "C:\temp\test_pdf_a_export\intermediate.pdf"
```

### Disable bookmarks (optional)

```powershell
python src\convert.py "C:\temp\test_pdf_a_export\test.docx" "C:\temp\test_pdf_a_export\test.pdf" --bookmarks none
```

## Output

- **Final PDF/A:** the output path you provide
- **Intermediate PDF:** automatically created and cleaned up unless `--intermediate-pdf` is specified

## Validation (recommended)

Use a validator such as **veraPDF** or **Acrobat Preflight** to confirm compliance:

```powershell
verapdf "C:\temp\test_pdf_a_export\test.pdf"
```

## Troubleshooting

- If bookmarks are missing, your headings may not be true Heading styles with outline levels.
- In your PDF viewer, open the **Bookmarks/Outline panel**, not the thumbnails panel.
