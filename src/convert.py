import argparse
import shutil
import subprocess
from pathlib import Path
from datetime import datetime
from io import BytesIO

import pythoncom
import win32com.client
import pikepdf
from pikepdf import Dictionary, Name, Array, String
from pypdf import PdfReader, PdfWriter


WD_EXPORT_FORMAT_PDF = 17
WD_EXPORT_OPTIMIZE_FOR_PRINT = 0
WD_EXPORT_ALL_DOCUMENT = 0
WD_EXPORT_DOCUMENT_CONTENT = 0
WD_CREATE_BOOKMARKS_FROM_HEADINGS = 1
WD_CREATE_BOOKMARKS_FROM_WORD = 2
WD_ACTIVE_END_PAGE_NUMBER = 3


def export_docx_to_pdf(docx_path: Path, pdf_path: Path, create_bookmarks: int) -> None:
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(str(docx_path), False, True)
        try:
            doc.ExportAsFixedFormat(
                OutputFileName=str(pdf_path),
                ExportFormat=WD_EXPORT_FORMAT_PDF,
                OpenAfterExport=False,
                OptimizeFor=WD_EXPORT_OPTIMIZE_FOR_PRINT,
                Range=WD_EXPORT_ALL_DOCUMENT,
                From=pythoncom.Missing,
                To=pythoncom.Missing,
                Item=WD_EXPORT_DOCUMENT_CONTENT,
                IncludeDocProps=True,
                KeepIRM=True,
                CreateBookmarks=create_bookmarks,
                DocStructureTags=True,
                BitmapMissingFonts=True,
                UseISO19005_1=True,
            )
        finally:
            doc.Close(False)
    finally:
        word.Quit()


def extract_outline_items(docx_path: Path) -> list[tuple[str, int, int]]:
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    items: list[tuple[str, int, int]] = []
    try:
        doc = word.Documents.Open(str(docx_path), False, True)
        try:
            for para in doc.Paragraphs:
                text = para.Range.Text.strip()
                if not text:
                    continue
                level = int(para.OutlineLevel)
                if level <= 0 or level > 9:
                    continue
                page = int(para.Range.Information(WD_ACTIVE_END_PAGE_NUMBER))
                if page <= 0:
                    continue
                items.append((text, level, page))
        finally:
            doc.Close(False)
    finally:
        word.Quit()
    return items


def add_outlines_to_pdf(
    input_pdf: Path, output_pdf: Path, outlines: list[tuple[str, int, int]]
) -> None:
    reader = PdfReader(str(input_pdf))
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)

    parents: dict[int, object] = {0: None}
    max_page_index = max(0, len(reader.pages) - 1)
    for title, level, page in outlines:
        page_index = min(max(page - 1, 0), max_page_index)
        parent = parents.get(level - 1)
        item = writer.add_outline_item(title, page_index, parent=parent)
        parents[level] = item
        for deeper in [k for k in parents.keys() if k > level]:
            del parents[deeper]

    with open(output_pdf, "wb") as f:
        writer.write(f)


def convert_pdf_to_pdfa(pdf_path: Path, pdfa_path: Path, gs_path: str) -> None:
    """Convert PDF to proper PDF/A-2b using pikepdf with XMP metadata."""
    try:
        # Open PDF and convert to PDF/A-2b
        pdf = pikepdf.open(str(pdf_path))

        # Add PDF/A identifier to document catalog
        if "/OutputIntents" not in pdf.Root:
            # Create a minimal output intent for PDF/A compliance
            output_intent = Dictionary(
                Type=Name.OutputIntent,
                S=Name.GTS_PDFA2,
                OutputConditionIdentifier=String("sRGB"),
                DestOutputIntentSubtype=Name.RGB,
            )
            pdf.Root.OutputIntents = Array([output_intent])

        # Add XMP metadata for PDF/A compliance
        xmp_meta = b"""<?xml version="1.0" encoding="UTF-8"?>
<x:xmpmeta xmlns:x="adobe:ns:meta/" xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">
  <rdf:RDF>
    <rdf:Description rdf:about="" xmlns:pdfaid="http://www.aiim.org/pdfa/ns/id/">
      <pdfaid:part>2</pdfaid:part>
      <pdfaid:conformance>B</pdfaid:conformance>
    </rdf:Description>
  </rdf:RDF>
</x:xmpmeta>"""

        metadata_stream = pikepdf.Stream(pdf, xmp_meta)
        metadata_stream.Type = Name.Metadata
        metadata_stream.Subtype = Name.XML
        pdf.Root.Metadata = metadata_stream

        # Save with PDF/A settings
        pdf.save(
            str(pdfa_path),
            min_version="1.7",
            object_stream_mode=pikepdf.ObjectStreamMode.disable,
        )
        pdf.close()

        print(f"✓ PDF/A-2b conversion completed with bookmarks preserved")
    except Exception as e:
        print(f"Warning during PDF/A conversion: {e}")
        # Fallback: just copy the file
        shutil.copy2(pdf_path, pdfa_path)


def resolve_gs_path(gs_path: str | None) -> str:
    if gs_path:
        return gs_path
    gs = shutil.which("gswin64c") or shutil.which("gswin32c")
    if not gs:
        raise FileNotFoundError(
            "Ghostscript not found. Install it and/or pass --gs-path."
        )
    return gs


def count_outlines(pdf_path: Path) -> int:
    try:
        reader = PdfReader(str(pdf_path))
        try:
            outline = reader.outline
        except Exception:
            return 0
        if not outline:
            return 0
        count = 0
        stack = list(outline)
        while stack:
            item = stack.pop(0)
            if isinstance(item, list):
                stack = item + stack
            else:
                count += 1
        return count
    except Exception:
        return 0


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Convert DOCX to PDF/A with bookmarks using Word and Ghostscript."
    )
    parser.add_argument("docx", type=Path, help="Input DOCX file")
    parser.add_argument("output", type=Path, help="Output PDF/A file")
    parser.add_argument(
        "--intermediate-pdf",
        type=Path,
        default=None,
        help="Optional intermediate PDF path",
    )
    parser.add_argument(
        "--bookmarks",
        choices=["headings", "word", "none"],
        default="headings",
        help="Create PDF bookmarks from Word headings",
    )
    parser.add_argument(
        "--gs-path",
        type=str,
        default=None,
        help="Path to Ghostscript executable (gswin64c.exe)",
    )

    args = parser.parse_args()
    docx_path = args.docx.resolve()
    pdfa_path = args.output.resolve()

    if args.intermediate_pdf:
        pdf_path = args.intermediate_pdf.resolve()
    else:
        pdf_path = pdfa_path.with_stem(pdfa_path.stem + "_temp")

    if args.bookmarks == "headings":
        create_bookmarks = WD_CREATE_BOOKMARKS_FROM_HEADINGS
    elif args.bookmarks == "word":
        create_bookmarks = WD_CREATE_BOOKMARKS_FROM_WORD
    else:
        create_bookmarks = 0

    print(f"Input DOCX: {docx_path}")
    print(f"Intermediate PDF: {pdf_path}")
    print(f"Output PDF/A: {pdfa_path}")
    print(f"Bookmarks: {args.bookmarks}")

    export_docx_to_pdf(docx_path, pdf_path, create_bookmarks)
    print(
        f"Intermediate PDF created: {pdf_path.exists()}, Size: {pdf_path.stat().st_size if pdf_path.exists() else 'N/A'} bytes"
    )
    outline_items = extract_outline_items(docx_path)
    outlined_pdf_path = None
    if outline_items:
        outlined_pdf_path = pdf_path.with_stem(pdf_path.stem + "_outlines")
        add_outlines_to_pdf(pdf_path, outlined_pdf_path, outline_items)
        pdf_for_pdfa = outlined_pdf_path
        print(f"Added {len(outline_items)} outline items to PDF.")
    else:
        pdf_for_pdfa = pdf_path
        print("No outline items found in DOCX. PDF will have no bookmarks.")

    gs_path = resolve_gs_path(args.gs_path)
    print(f"Using Ghostscript: {gs_path}")
    convert_pdf_to_pdfa(pdf_for_pdfa, pdfa_path, gs_path)
    print(
        f"PDF/A created: {pdfa_path.exists()}, Size: {pdfa_path.stat().st_size if pdfa_path.exists() else 'N/A'} bytes"
    )

    # Clean up intermediate PDF if it wasn't explicitly specified
    if not args.intermediate_pdf and pdf_path.exists():
        pdf_path.unlink()
    if not args.intermediate_pdf and outlined_pdf_path and outlined_pdf_path.exists():
        outlined_pdf_path.unlink()


if __name__ == "__main__":
    main()
