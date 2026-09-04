"""Document Extraction (No LLMs).

Extracts metadata, structure (headings/paragraphs), tables and text from a PDF or
DOCX using pdfplumber, PyMuPDF and python-docx only. No LLM, no API, no network.

Usage:  python extract_document.py <file.pdf|file.docx> [outdir]
Writes: <outdir>/<name>.json  and  <outdir>/<name>.md   (default outdir: output)
"""

import json
import logging
import re
import sys
from collections import Counter
from pathlib import Path

import pdfplumber
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

try:
    import pymupdf                  # PyMuPDF 1.24+
except ImportError:
    import fitz as pymupdf          # older name

logging.getLogger("pdfminer").setLevel(logging.ERROR)   # hide font warnings

MAX_HEADING = 120                   # a heading is short; longer lines are prose
CLEAN = lambda v: re.sub(r"\s+", " ", v or "").strip()


def pdf_read(path):
    """Return (metadata, [(page, text, font_size), ...]) for every text line."""
    lines = []
    with pymupdf.open(path) as doc:
        meta = {"file": path.name, "type": "PDF", "pages": doc.page_count,
                **{k: v for k, v in (doc.metadata or {}).items() if v}}
        for page_no, page in enumerate(doc, 1):
            for block in page.get_text("dict")["blocks"]:
                for line in block.get("lines", []):          # image blocks have none
                    text = CLEAN("".join(s["text"] for s in line["spans"]))
                    if text:
                        size = round(max(s["size"] for s in line["spans"]), 1)
                        lines.append((page_no, text, size))
    return meta, lines


def pdf_structure(lines):
    """Classify lines into headings and paragraphs by measured font size."""
    volume = Counter()
    for _, text, size in lines:
        volume[size] += len(text)                            # body = most characters
    body = volume.most_common(1)[0][0]
    big = sorted({s for _, t, s in lines
                  if s >= body * 1.15 and len(t) <= MAX_HEADING}, reverse=True)[:4]
    level = {size: n for n, size in enumerate(big, 1)}       # biggest size = level 1

    blocks, seen = [], Counter(CLEAN(re.sub(r"\d", "#", t)) for _, t, _ in lines)
    for page_no, text, size in lines:
        if seen[CLEAN(re.sub(r"\d", "#", text))] > 3:
            continue                                         # running header/footer
        if size in level and len(text) <= MAX_HEADING and text[-1] not in ".;,":
            blocks.append({"type": "heading", "level": level[size],
                           "page": page_no, "text": text})
        elif blocks and blocks[-1]["type"] == "paragraph" and blocks[-1]["page"] == page_no:
            blocks[-1]["text"] += " " + text                 # same paragraph
        else:
            blocks.append({"type": "paragraph", "page": page_no, "text": text})
    return blocks


def pdf_tables(path):
    """Extract tables with pdfplumber, dropping rows that are not tabular."""
    tables = []
    with pdfplumber.open(path) as pdf:
        for page_no, page in enumerate(pdf.pages, 1):
            for raw in page.extract_tables():
                rows = [[CLEAN(c) for c in row] for row in raw]
                rows = [r for r in rows if sum(1 for c in r if c) >= 2]
                if len(rows) >= 2:                           # else a layout box
                    tables.append({"page": page_no, "rows": len(rows),
                                   "columns": max(map(len, rows)), "data": rows})
    return tables


def docx_read(path):
    """Return (metadata, blocks, tables) walking the body in document order."""
    doc = Document(path)
    prop = doc.core_properties
    meta = {"file": path.name, "type": "DOCX", "pages": None,
            "title": prop.title, "author": prop.author, "subject": prop.subject,
            "keywords": prop.keywords, "revision": prop.revision,
            "created": str(prop.created), "modified": str(prop.modified)}
    blocks, tables = [], []
    for child in doc.element.body.iterchildren():            # keeps table position
        if child.tag.endswith("}p"):
            para = Paragraph(child, doc)
            text = CLEAN(para.text)
            if not text:
                continue
            found = re.match(r"heading (\d+)|(title)", (para.style.name or "").lower())
            if found:
                blocks.append({"type": "heading",
                               "level": int(found.group(1) or 1), "text": text})
            else:
                blocks.append({"type": "paragraph", "text": text})
        elif child.tag.endswith("}tbl"):
            rows = [[CLEAN(c.text) for c in r.cells] for r in Table(child, doc).rows]
            tables.append({"page": None, "rows": len(rows),
                           "columns": max(map(len, rows)), "data": rows})
            blocks.append({"type": "table", "index": len(tables)})
    return meta, blocks, tables


def write(result, outdir, stem):
    """Write the JSON result plus a readable Markdown report."""
    outdir.mkdir(parents=True, exist_ok=True)
    (outdir / f"{stem}.json").write_text(
        json.dumps(result, indent=2, ensure_ascii=False), encoding="utf-8")

    out = [f"# {result['metadata']['file']}", ""]
    out += [f"- **{k}**: {v}" for k, v in result["metadata"].items() if v] + [""]
    for block in result["structure"]:
        if block["type"] == "heading":
            out += ["#" * (block["level"] + 1) + " " + block["text"], ""]
        elif block["type"] == "paragraph":
            out += [block["text"], ""]
    for n, table in enumerate(result["tables"], 1):
        width = table["columns"]
        rows = [r + [""] * (width - len(r)) for r in table["data"]]
        out += [f"### Table {n} (page {table['page']}) - "
                f"{table['rows']}x{width}", "",
                "| " + " | ".join(rows[0]) + " |",
                "|" + "|".join([" --- "] * width) + "|"]
        out += ["| " + " | ".join(r) + " |" for r in rows[1:]] + [""]
    (outdir / f"{stem}.md").write_text("\n".join(out), encoding="utf-8")


def main():
    if len(sys.argv) < 2:
        sys.exit("usage: python extract_document.py <file.pdf|file.docx> [outdir]")
    path = Path(sys.argv[1])
    outdir = Path(sys.argv[2] if len(sys.argv) > 2 else "output")
    if not path.exists():
        sys.exit(f"file not found: {path}")

    if path.suffix.lower() == ".pdf":
        meta, lines = pdf_read(path)
        blocks, tables = pdf_structure(lines), pdf_tables(path)
    elif path.suffix.lower() == ".docx":
        meta, blocks, tables = docx_read(path)
    else:
        sys.exit(f"unsupported type '{path.suffix}' - use .pdf or .docx")

    result = {"metadata": meta, "structure": blocks, "tables": tables,
              "text": "\n".join(b["text"] for b in blocks if "text" in b)}
    write(result, outdir, path.stem.replace(" ", "_"))

    kinds = Counter(b["type"] for b in blocks)
    print(f"{meta['file']}: {kinds['heading']} headings, {kinds['paragraph']} "
          f"paragraphs, {len(tables)} tables, {len(result['text'])} chars "
          f"-> {outdir}/{path.stem.replace(' ', '_')}.json|.md")


if __name__ == "__main__":
    main()