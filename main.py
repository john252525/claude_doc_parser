import io
import struct
import zipfile
from urllib.parse import urlparse

import httpx
import olefile
import xlrd
from PyPDF2 import PdfReader
from docx import Document as DocxDocument
from docx.table import Table
from docx.text.paragraph import Paragraph
from fastapi import FastAPI, Query, HTTPException
from openpyxl import load_workbook

app = FastAPI(title="Document Text Extractor")

SUPPORTED_EXTENSIONS = {".pdf", ".doc", ".docx", ".xlsx", ".xls"}
MAX_DOWNLOAD_SIZE = 50 * 1024 * 1024  # 50 MB


# ---------------------------------------------------------------------------
# PDF
# ---------------------------------------------------------------------------
def extract_pdf(data: bytes) -> str:
    reader = PdfReader(io.BytesIO(data))
    pages = []
    for page in reader.pages:
        text = page.extract_text()
        if text:
            pages.append(text)
    return "\n\n".join(pages)


# ---------------------------------------------------------------------------
# DOCX  (paragraphs + tables in document order)
# ---------------------------------------------------------------------------
def extract_docx(data: bytes) -> str:
    doc = DocxDocument(io.BytesIO(data))
    parts: list[str] = []

    for element in doc.element.body:
        tag = element.tag

        # Paragraph
        if tag.endswith("}p"):
            para = Paragraph(element, doc)
            if para.text.strip():
                parts.append(para.text)

        # Table
        elif tag.endswith("}tbl"):
            table = Table(element, doc)
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells]
                parts.append("\t".join(cells))
            parts.append("")  # blank line after table

    return "\n".join(parts)


# ---------------------------------------------------------------------------
# DOC  (legacy Word 97‑2003, OLE2 binary format)
# ---------------------------------------------------------------------------
def extract_doc(data: bytes) -> str:
    ole = olefile.OleFileIO(io.BytesIO(data))
    try:
        word_data = ole.openstream("WordDocument").read()

        # FIB flags → choose the right table stream
        flags = struct.unpack_from("<H", word_data, 0x000A)[0]
        table_name = "1Table" if (flags & 0x0200) else "0Table"
        if not ole.exists(table_name):
            table_name = "0Table" if table_name == "1Table" else "1Table"

        table_data = ole.openstream(table_name).read()

        # CLX offset / size live at fixed FIB positions (Word 97‑2003)
        fc_clx = struct.unpack_from("<I", word_data, 0x01A2)[0]
        lcb_clx = struct.unpack_from("<I", word_data, 0x01A6)[0]
        if lcb_clx == 0:
            raise ValueError("Empty CLX in .doc file")

        clx = table_data[fc_clx : fc_clx + lcb_clx]

        # Skip optional Prc entries (0x01 prefix)
        pos = 0
        while pos < len(clx) and clx[pos] == 0x01:
            cb = struct.unpack_from("<H", clx, pos + 1)[0]
            pos += 3 + cb

        # Pcdt must start with 0x02
        if pos >= len(clx) or clx[pos] != 0x02:
            raise ValueError("Could not locate piece table in .doc")

        pos += 1
        lcb = struct.unpack_from("<I", clx, pos)[0]
        pos += 4
        piece_table = clx[pos : pos + lcb]

        # piece_table = (n+1) CPs (4 bytes each) + n PCDs (8 bytes each)
        n = (len(piece_table) - 4) // 12

        text_parts: list[str] = []
        for i in range(n):
            cp_start = struct.unpack_from("<I", piece_table, i * 4)[0]
            cp_end = struct.unpack_from("<I", piece_table, (i + 1) * 4)[0]

            pcd_offset = (n + 1) * 4 + i * 8
            fc_value = struct.unpack_from("<I", piece_table, pcd_offset + 2)[0]

            is_compressed = bool(fc_value & 0x40000000)
            fc = fc_value & 0x3FFFFFFF
            char_count = cp_end - cp_start

            if is_compressed:
                byte_offset = fc // 2
                raw = word_data[byte_offset : byte_offset + char_count]
                text_parts.append(raw.decode("cp1252", errors="replace"))
            else:
                raw = word_data[fc : fc + char_count * 2]
                text_parts.append(raw.decode("utf-16-le", errors="replace"))

        full_text = "".join(text_parts)

        # Convert Word‑specific control characters
        result: list[str] = []
        for ch in full_text:
            code = ord(ch)
            if code in (0x0D, 0x0B):
                result.append("\n")
            elif code == 0x07:
                result.append("\t")
            elif code == 0x0C:
                result.append("\n\n")
            elif code >= 0x20 or ch in ("\n", "\r", "\t"):
                result.append(ch)
        return "".join(result).strip()
    finally:
        ole.close()


# ---------------------------------------------------------------------------
# XLSX  (modern Excel, via openpyxl)
# ---------------------------------------------------------------------------
def extract_xlsx(data: bytes) -> str:
    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    sheets = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = []
        for row in ws.iter_rows(values_only=True):
            cells = [str(c) if c is not None else "" for c in row]
            if any(cells):
                rows.append("\t".join(cells))
        if rows:
            sheets.append(f"--- {sheet_name} ---\n" + "\n".join(rows))
    wb.close()
    return "\n\n".join(sheets)


# ---------------------------------------------------------------------------
# XLS  (legacy Excel 97‑2003, via xlrd)
# ---------------------------------------------------------------------------
def extract_xls(data: bytes) -> str:
    wb = xlrd.open_workbook(file_contents=data)
    sheets = []
    for sheet_idx in range(wb.nsheets):
        ws = wb.sheet_by_index(sheet_idx)
        rows = []
        for row_idx in range(ws.nrows):
            cells = []
            for col_idx in range(ws.ncols):
                cell = ws.cell(row_idx, col_idx)
                if cell.ctype == xlrd.XL_CELL_EMPTY:
                    cells.append("")
                elif cell.ctype == xlrd.XL_CELL_NUMBER:
                    # Show integers without ".0"
                    val = cell.value
                    cells.append(str(int(val)) if val == int(val) else str(val))
                else:
                    cells.append(str(cell.value))
            if any(cells):
                rows.append("\t".join(cells))
        if rows:
            sheets.append(f"--- {ws.name} ---\n" + "\n".join(rows))
    return "\n\n".join(sheets)


# ---------------------------------------------------------------------------
# Extractor registry
# ---------------------------------------------------------------------------
EXTRACTORS = {
    ".pdf": extract_pdf,
    ".doc": extract_doc,
    ".docx": extract_docx,
    ".xlsx": extract_xlsx,
    ".xls": extract_xls,
}


# ---------------------------------------------------------------------------
# File‑type detection
# ---------------------------------------------------------------------------
def guess_from_magic(data: bytes) -> str | None:
    """Detect file type by magic bytes / internal structure."""
    if data[:4] == b"%PDF":
        return ".pdf"

    # DOCX / XLSX — both are ZIP archives
    if data[:4] == b"PK\x03\x04":
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as zf:
                names = zf.namelist()
                if any(n.startswith("word/") for n in names):
                    return ".docx"
                if any(n.startswith("xl/") for n in names):
                    return ".xlsx"
        except zipfile.BadZipFile:
            pass

    # OLE2 compound document — distinguish .doc from .xls
    if data[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        try:
            ole = olefile.OleFileIO(io.BytesIO(data))
            try:
                if ole.exists("WordDocument"):
                    return ".doc"
                if ole.exists("Workbook") or ole.exists("Book"):
                    return ".xls"
            finally:
                ole.close()
        except Exception:
            pass
        return ".xls"  # safe fallback for OLE2

    return None


def guess_extension(url: str, content_type: str | None, data: bytes) -> str | None:
    # 1. URL path
    parsed = urlparse(url)
    path = parsed.path.lower()
    for ext in SUPPORTED_EXTENSIONS:
        if path.endswith(ext):
            return ext

    # 2. Content‑Type header
    ct_map = {
        "application/pdf": ".pdf",
        "application/msword": ".doc",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document": ".docx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": ".xlsx",
        "application/vnd.ms-excel": ".xls",
    }
    if content_type:
        for mime, ext in ct_map.items():
            if mime in content_type:
                return ext

    # 3. Magic bytes
    return guess_from_magic(data)


# ---------------------------------------------------------------------------
# API endpoint
# ---------------------------------------------------------------------------
@app.get("/extract")
async def extract_text(
    url: str = Query(..., description="URL of the document to download and extract text from"),
):
    """
    Download a document (PDF, DOC, DOCX, XLS, XLSX) by URL and return its text content.
    """
    try:
        async with httpx.AsyncClient(follow_redirects=True, timeout=30.0) as client:
            resp = await client.get(url)
            resp.raise_for_status()
    except httpx.HTTPStatusError as e:
        raise HTTPException(status_code=502, detail=f"Remote server returned {e.response.status_code}")
    except httpx.RequestError as e:
        raise HTTPException(status_code=502, detail=f"Failed to download document: {e}")

    if len(resp.content) > MAX_DOWNLOAD_SIZE:
        raise HTTPException(status_code=413, detail="Document too large (max 50 MB)")

    content_type = resp.headers.get("content-type", "")
    ext = guess_extension(url, content_type, resp.content)

    if ext is None:
        raise HTTPException(
            status_code=400,
            detail=f"Cannot determine document type. Supported: {', '.join(sorted(SUPPORTED_EXTENSIONS))}",
        )

    extractor = EXTRACTORS.get(ext)
    if extractor is None:
        raise HTTPException(status_code=400, detail=f"Unsupported file type: {ext}")

    try:
        text = extractor(resp.content)
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"Failed to extract text: {e}")

    return {
        "ok": True,
        "source_url": url,
        "file_type": ext.lstrip("."),
        "text": text,
    }
