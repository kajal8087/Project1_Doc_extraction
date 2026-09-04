Document Extraction (No LLMs)

1. Setup (Windows PowerShell)
```powershell
cd C:\Projects\doc-extraction
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

---
2. Running it
```powershell
python extract_document.py "C:\Docs\design-document.pdf"
python extract_document.py report.docx
python extract_document.py report.pdf results     # custom output folder
```
No sample file handy? `python make_sample_docx.py` builds one with a title, three
heading levels, a bullet list and a table.


4. Libraries
Library	Used for
`pymupdf`	PDF metadata, and the font size of every line (needed to find headings)
`pdfplumber`	PDF tables
`python-docx`	DOCX paragraphs, styles, tables, core properties

---
5. How it works
```
.pdf  -> pymupdf reads every line with its font size
         -> the body size is measured, larger short lines become headings
         -> pdfplumber extracts the tables
.docx -> python-docx reads core properties
         -> Word style names ("Heading 2") give the level directly
         -> the XML body is walked so tables keep their position in the text
      -> one dictionary -> JSON + Markdown
```


