# Document Extraction

This is a small Python tool which reads a PDF or Word file and takes out the headings, paragraphs, tables and file details, without using any AI. In a Word file the headings are already marked, so it just reads them. In a PDF nothing is marked, so the tool checks the font size of each line and whichever line is bigger than the normal text, that one is taken as a heading. The output comes as one JSON file and one readable report. I tested it on a 10 -page document and it picked up 15 heading and 4 tables correctly.


---


## 1. How to Run

python extract_document.py "C:\Docs\design-document.pdf"
python extract_document.py report.docx
python extract_document.py report.pdf results


The third one writes the output to a folder named `results` instead of the default
`output` folder.

No sample file handy? `python make_sample_docx.py` builds one with a title, three
heading levels, a bullet list and a table.

---

## 2. Output

Two files are created for every run:

| File | Contents |
| --- | --- |
| `<name>.json` | metadata, headings and paragraphs in order, every table, full text |
| `<name>.md` | the same thing as a readable report |

One line is printed on the screen after each run:

```
design-document.pdf: 15 headings, 4 tables -> output/design-document.json|.md
```

---

## 3. Libraries

| Library | Used for |
| --- | --- |
| `pymupdf` | PDF metadata, and the font size of every line (needed to find headings) |
| `pdfplumber` | PDF tables |
| `python-docx` | DOCX paragraphs, styles, tables, core properties |

---

## 4. How it works

```
.pdf  -> pymupdf reads every line with its font size
         -> the body size is measured, larger short lines become headings
         -> pdfplumber extracts the tables
.docx -> python-docx reads core properties
         -> Word style names ("Heading 2") give the level directly
         -> the XML body is walked so tables keep their position in the text
      -> one dictionary -> JSON + Markdown
```

