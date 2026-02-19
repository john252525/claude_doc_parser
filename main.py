import io
from urllib.parse import urlparse

import httpx
from PyPDF2 import PdfReader
from docx import Document as DocxDocument
from fastapi import FastAPI, Query, HTTPException
from openpyxl import load_workbook

app = FastAPI(title="Document Text Extractor")

SUPPORTED_EXTENSIONS = {".pdf", ".docx", ".xlsx", ".xls"}
MAX_DOWNLOAD_SIZE = 50 * 1024 * 1024  # 50 MB


def extract_pdf(data: bytes) -> str:
    reader = PdfReader(io.BytesIO(data))
    pages = []
    for page in reader.pages:
        text = page.extract_text()
        if text:
            pages.append(text)
    return "\n\n".join(pages)


def extract_docx(data: bytes) -> str:
    doc = DocxDocument(io.BytesIO(data))
    paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
    return "\n".join(paragraphs)


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


EXTRACTORS = {
    ".pdf": extract_pdf,
    ".docx": extract_docx,
    ".xlsx": extract_xlsx,
    ".xls": extract_xlsx,
}


def guess_extension(url: str, content_type: str | None) -> str | None:
    # Try from URL path first
    parsed = urlparse(url)
    path = parsed.path.lower()
    for ext in SUPPORTED_EXTENSIONS:
        if path.endswith(ext):
            return ext

    # Fallback to Content-Type header
    ct_map = {
        "application/pdf": ".pdf",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document": ".docx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": ".xlsx",
        "application/vnd.ms-excel": ".xls",
    }
    if content_type:
        for mime, ext in ct_map.items():
            if mime in content_type:
                return ext

    return None


@app.get("/extract")
async def extract_text(
    url: str = Query(..., description="URL of the document to download and extract text from"),
):
    """
    Download a document (PDF, DOCX, XLSX) by URL and return its text content.
    """
    # Download
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
    ext = guess_extension(url, content_type)

    if ext is None:
        raise HTTPException(
            status_code=400,
            detail=f"Cannot determine document type. Supported: {', '.join(SUPPORTED_EXTENSIONS)}",
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
