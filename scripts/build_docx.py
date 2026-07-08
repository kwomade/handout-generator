#!/usr/bin/env python3
"""Wandelt ein (Teilmengen-)Markdown in eine echte .docx um - ohne Microsoft Word.

Aufruf:
    python build_docx.py <input.md> <output.docx>

Unterstuetztes Markdown:
    # .. ####     Ueberschriften (Ebene 1-4)
    Absaetze      Fliesstext mit **fett** und `code`
    - / *         Aufzaehlung
    | a | b |     Tabellen (erste Zeile = Kopf, Trennzeile ---|--- wird ignoriert)
    ```           Codebloecke (Monospace, hinterlegt)
    >             Zitat / "Merke"-Kasten (Rahmen links)
    ---           Trennlinie

Exit-Codes: 0 = ok, 1 = falscher Aufruf, 3 = python-docx fehlt.
"""
import sys
import re
import pathlib

try:
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError:
    sys.stderr.write("DOCX_MISSING: 'python-docx' ist nicht installiert. "
                     "Bitte 'pip install --user python-docx' ausfuehren.\n")
    sys.exit(3)


INLINE_RE = re.compile(r'(\*\*.+?\*\*|`[^`]+`)')
SEP_ROW_RE = re.compile(r'^\s*\|[\s:\-\|]+\|?\s*$')
HEADING_RE = re.compile(r'^(#{1,4})\s+(.*)')
QUOTE_RE = re.compile(r'^>\s?')
BULLET_RE = re.compile(r'^\s*[-*]\s+(.*)')
RULE_RE = re.compile(r'^\s*---\s*$')


def add_inline(paragraph, text):
    """Fuegt Text mit **fett** und `code` als Runs hinzu."""
    pos = 0
    for m in INLINE_RE.finditer(text):
        if m.start() > pos:
            paragraph.add_run(text[pos:m.start()])
        tok = m.group(0)
        if tok.startswith("**"):
            run = paragraph.add_run(tok[2:-2])
            run.bold = True
        else:
            run = paragraph.add_run(tok[1:-1])
            run.font.name = "Consolas"
        pos = m.end()
    if pos < len(text):
        paragraph.add_run(text[pos:])


def _shade(pPr, fill):
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:fill'), fill)
    pPr.append(shd)


def shade_paragraph(paragraph, fill):
    _shade(paragraph._p.get_or_add_pPr(), fill)


def shade_cell(cell, fill):
    _shade(cell._tc.get_or_add_tcPr(), fill)


def add_border(paragraph, side, color="AAAAAA", sz="6", space="1"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = pPr.find(qn('w:pBdr'))
    if pBdr is None:
        pBdr = OxmlElement('w:pBdr')
        pPr.append(pBdr)
    edge = OxmlElement('w:' + side)
    edge.set(qn('w:val'), 'single')
    edge.set(qn('w:sz'), sz)
    edge.set(qn('w:space'), space)
    edge.set(qn('w:color'), color)
    pBdr.append(edge)


def add_table(doc, rows):
    if not rows:
        return
    ncols = max(len(r) for r in rows)
    table = doc.add_table(rows=len(rows), cols=ncols)
    table.style = 'Table Grid'
    for ri, row in enumerate(rows):
        for ci in range(ncols):
            cell = table.cell(ri, ci)
            cell.text = ""
            para = cell.paragraphs[0]
            val = row[ci] if ci < len(row) else ""
            if ri == 0:
                run = para.add_run(val)
                run.bold = True
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                shade_cell(cell, "4472C4")
            else:
                add_inline(para, val)
    doc.add_paragraph()  # Abstand nach der Tabelle


def build(md_path, docx_path):
    lines = md_path.read_text(encoding="utf-8").splitlines()
    doc = Document()

    # Grundschrift etwas groesser fuer Lesbarkeit
    normal = doc.styles['Normal']
    normal.font.size = Pt(11)

    in_code = False
    table_rows = []

    def flush_table():
        if table_rows:
            add_table(doc, table_rows)
            table_rows.clear()

    for raw in lines:
        line = raw

        # Zitatpraefix erkennen und entfernen (erlaubt auch Code/Listen im Zitat)
        quoted = False
        if QUOTE_RE.match(line):
            quoted = True
            line = QUOTE_RE.sub('', line, count=1)

        # Codeblock-Umschaltung
        if line.strip().startswith("```"):
            in_code = not in_code
            continue
        if in_code:
            para = doc.add_paragraph()
            run = para.add_run(line)
            run.font.name = "Consolas"
            run.font.size = Pt(9)
            shade_paragraph(para, "F2F2F2")
            para.paragraph_format.space_after = Pt(0)
            if quoted:
                para.paragraph_format.left_indent = Pt(18)
            continue

        # Tabellenzeilen
        if line.strip().startswith("|"):
            if SEP_ROW_RE.match(line):
                continue
            cells = [c.strip() for c in line.strip().strip('|').split('|')]
            table_rows.append(cells)
            continue
        else:
            flush_table()

        # Leerzeile
        if line.strip() == "":
            doc.add_paragraph()
            continue

        # Trennlinie
        if RULE_RE.match(line):
            para = doc.add_paragraph()
            add_border(para, 'bottom')
            continue

        # Ueberschrift
        m = HEADING_RE.match(line)
        if m:
            level = len(m.group(1))
            doc.add_heading(m.group(2), level=level)
            continue

        # Aufzaehlung
        m = BULLET_RE.match(line)
        if m:
            para = doc.add_paragraph(style='List Bullet')
            add_inline(para, m.group(1))
            if quoted:
                para.paragraph_format.left_indent = Pt(36)
            continue

        # Normaler Absatz (ggf. als Merke-Kasten)
        para = doc.add_paragraph()
        add_inline(para, line)
        if quoted:
            para.paragraph_format.left_indent = Pt(18)
            add_border(para, 'left', color="8888AA", sz="18", space="8")

    flush_table()
    doc.save(str(docx_path))
    print("OK\t%s" % docx_path)


def main():
    if len(sys.argv) < 3:
        sys.stderr.write("usage: build_docx.py <input.md> <output.docx>\n")
        sys.exit(1)
    md_path = pathlib.Path(sys.argv[1])
    docx_path = pathlib.Path(sys.argv[2])
    if not md_path.is_file():
        sys.stderr.write("Eingabedatei nicht gefunden: %s\n" % md_path)
        sys.exit(1)
    build(md_path, docx_path)


if __name__ == "__main__":
    main()
