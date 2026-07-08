#!/usr/bin/env python3
"""Wandelt ein (Teilmengen-)Markdown in eine echte .docx um - ohne Microsoft Word.

Aufruf:
    python build_docx.py <input.md> <output.docx>

Unterstuetztes Markdown:
    # .. ####     Ueberschriften (Ebene 1-4)
    Absaetze      Fliesstext mit **fett** und `code`
    - / *         Aufzaehlung
    | a | b |     Tabellen (erste Zeile = Kopf, Trennzeile ---|--- wird ignoriert)
    ```lang       Codebloecke mit Syntax-Highlighting (Java, Python, SQL, JSON, XML, JS)
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


# ---------------------------------------------------------------------------
# Regex-Bausteine
# ---------------------------------------------------------------------------
INLINE_RE = re.compile(r'(\*\*.+?\*\*|`[^`]+`)')
SEP_ROW_RE = re.compile(r'^\s*\|[\s:\-\|]+\|?\s*$')
HEADING_RE = re.compile(r'^(#{1,4})\s+(.*)')
QUOTE_RE = re.compile(r'^>\s?')
BULLET_RE = re.compile(r'^\s*[-*]\s+(.*)')
RULE_RE = re.compile(r'^\s*---\s*$')
WORD_RE = re.compile(r'@?\w+')
NUM_RE = re.compile(r'\d+(\.\d+)?')

# ---------------------------------------------------------------------------
# Farben fuer Syntax-Highlighting (helles, druckfreundliches Schema)
# ---------------------------------------------------------------------------
C_BG = "F6F8FA"       # Hintergrund der Code-Box
C_BORDER = "D0D7DE"   # Rahmen
C_TEXT = "24292E"     # normaler Code-Text
C_KEYWORD = "0033B3"  # Schluesselwoerter (blau)
C_STRING = "067D17"   # Zeichenketten (gruen)
C_COMMENT = "8C8C8C"  # Kommentare (grau)
C_NUMBER = "1750EB"   # Zahlen
C_TAG = "0033B3"      # XML/HTML-Tags
C_LABEL = "6E7781"    # Sprach-Label

# ---------------------------------------------------------------------------
# Schluesselwoerter je Sprache
# ---------------------------------------------------------------------------
KW = {
    "java": set("""abstract assert boolean break byte case catch char class const
        continue default do double else enum extends final finally float for goto if
        implements import instanceof int interface long native new package private
        protected public return short static strictfp super switch synchronized this
        throw throws transient try void volatile while true false null var""".split()),
    "python": set("""def class return if elif else for while import from as try except
        finally with lambda None True False and or not in is pass break continue global
        nonlocal yield raise del assert async await elif print self""".split()),
    "sql": set("""select from where insert into values update set delete create table
        alter drop join inner left right outer full cross on group by order having
        primary key foreign references not null unique and or as count sum avg min max
        distinct limit offset like between in is asc desc auto_increment varchar int
        default union all exists""".split()),
    "json": set("true false null".split()),
    "js": set("""var let const function return if else for while do switch case break
        continue new class extends this typeof instanceof null true false undefined
        import export from default async await try catch finally throw of in""".split()),
}
KW["javascript"] = KW["js"]
KW["python3"] = KW["python"]

# Kommentar-Marker je Sprache: (Zeilenkommentar, Blockstart, Blockende)
COMMENTS = {
    "java": ("//", "/*", "*/"),
    "js": ("//", "/*", "*/"),
    "javascript": ("//", "/*", "*/"),
    "sql": ("--", "/*", "*/"),
    "python": ("#", None, None),
    "python3": ("#", None, None),
    "json": (None, None, None),
    "xml": (None, "<!--", "-->"),
    "html": (None, "<!--", "-->"),
}


def _kw_hit(word, lang):
    kws = KW.get(lang)
    if not kws:
        return False
    if lang == "sql":
        return word.lower() in kws
    return word in kws


def tokenize_code(line, lang, state):
    """Zerlegt eine Codezeile in (text, farbe|None)-Paare fuer das Highlighting."""
    out = []
    i, n = 0, len(line)
    line_c, block_o, block_c = COMMENTS.get(lang, (None, None, None))

    # XML/HTML: eigener, einfacher Modus (Tags + Strings + Kommentare)
    if lang in ("xml", "html"):
        return _tokenize_xml(line, state)

    while i < n:
        # laufender Blockkommentar
        if state.get("block"):
            end = line.find(block_c, i) if block_c else -1
            if end == -1:
                out.append((line[i:], C_COMMENT)); i = n
            else:
                out.append((line[i:end + len(block_c)], C_COMMENT))
                i = end + len(block_c); state["block"] = False
            continue
        # Blockkommentar-Beginn
        if block_o and line.startswith(block_o, i):
            end = line.find(block_c, i + len(block_o))
            if end == -1:
                out.append((line[i:], C_COMMENT)); i = n; state["block"] = True
            else:
                out.append((line[i:end + len(block_c)], C_COMMENT)); i = end + len(block_c)
            continue
        # Zeilenkommentar
        if line_c and line.startswith(line_c, i):
            out.append((line[i:], C_COMMENT)); i = n
            continue
        ch = line[i]
        # Zeichenkette
        if ch == '"' or ch == "'":
            j = i + 1
            while j < n and line[j] != ch:
                if line[j] == "\\":
                    j += 1
                j += 1
            j = min(j + 1, n)
            out.append((line[i:j], C_STRING)); i = j
            continue
        # Wort (Schluesselwort / Zahl / normal)
        m = WORD_RE.match(line, i)
        if m:
            w = m.group(0)
            if NUM_RE.fullmatch(w):
                color = C_NUMBER
            elif _kw_hit(w, lang):
                color = C_KEYWORD
            else:
                color = None
            out.append((w, color)); i = m.end()
            continue
        # sonstiges Einzelzeichen
        out.append((ch, None)); i += 1
    return out


def _tokenize_xml(line, state):
    out = []
    i, n = 0, len(line)
    while i < n:
        if state.get("block"):
            end = line.find("-->", i)
            if end == -1:
                out.append((line[i:], C_COMMENT)); i = n
            else:
                out.append((line[i:end + 3], C_COMMENT)); i = end + 3; state["block"] = False
            continue
        if line.startswith("<!--", i):
            end = line.find("-->", i + 4)
            if end == -1:
                out.append((line[i:], C_COMMENT)); i = n; state["block"] = True
            else:
                out.append((line[i:end + 3], C_COMMENT)); i = end + 3
            continue
        ch = line[i]
        if ch == '"' or ch == "'":
            j = i + 1
            while j < n and line[j] != ch:
                j += 1
            j = min(j + 1, n)
            out.append((line[i:j], C_STRING)); i = j
            continue
        if ch == "<":
            m = re.match(r'</?[\w:.-]+', line[i:])
            if m:
                out.append((m.group(0), C_TAG)); i += len(m.group(0)); continue
        if ch == ">" or (ch == "/" and i + 1 < n and line[i + 1] == ">"):
            out.append((ch, C_TAG)); i += 1
            continue
        m = WORD_RE.match(line, i)
        if m:
            out.append((m.group(0), None)); i = m.end()
            continue
        out.append((ch, None)); i += 1
    return out


# ---------------------------------------------------------------------------
# Inline-Formatierung fuer normalen Text
# ---------------------------------------------------------------------------
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
            run.font.color.rgb = RGBColor(0xB0, 0x30, 0x60)
        pos = m.end()
    if pos < len(text):
        paragraph.add_run(text[pos:])


# ---------------------------------------------------------------------------
# OOXML-Hilfsfunktionen (Schattierung, Rahmen)
# ---------------------------------------------------------------------------
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


def set_cell_border(cell, color, sz="4"):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        e = OxmlElement('w:' + edge)
        e.set(qn('w:val'), 'single')
        e.set(qn('w:sz'), sz)
        e.set(qn('w:space'), '0')
        e.set(qn('w:color'), color)
        borders.append(e)
    tcPr.append(borders)


def _preserve_space(run):
    t = run._r.find(qn('w:t'))
    if t is not None:
        t.set(qn('xml:space'), 'preserve')


# ---------------------------------------------------------------------------
# Bloecke: Tabelle und Code-Box
# ---------------------------------------------------------------------------
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
    doc.add_paragraph()


def add_code_block(doc, code_lines, lang):
    """Rendert einen Codeblock als Code-Box mit Syntax-Highlighting."""
    # Fuehrende/abschliessende Leerzeilen entfernen
    while code_lines and code_lines[0].strip() == "":
        code_lines.pop(0)
    while code_lines and code_lines[-1].strip() == "":
        code_lines.pop()
    if not code_lines:
        return

    lang = (lang or "").lower()

    # Optionales Sprach-Label ueber der Box
    if lang and lang not in ("text", "plain", "output"):
        cap = doc.add_paragraph()
        cap.paragraph_format.space_before = Pt(6)
        cap.paragraph_format.space_after = Pt(0)
        run = cap.add_run(lang.upper())
        run.bold = True
        run.font.size = Pt(7)
        run.font.name = "Consolas"
        run.font.color.rgb = RGBColor(0x6E, 0x77, 0x81)

    table = doc.add_table(rows=1, cols=1)
    table.autofit = True
    cell = table.cell(0, 0)
    cell.text = ""
    shade_cell(cell, C_BG)
    set_cell_border(cell, C_BORDER)

    state = {"block": False}
    first = True
    for line in code_lines:
        line = line.replace("\t", "    ")
        para = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False
        para.paragraph_format.space_after = Pt(0)
        para.paragraph_format.space_before = Pt(0)
        if line.strip() == "":
            continue
        for text, color in tokenize_code(line, lang, state):
            run = para.add_run(text)
            run.font.name = "Consolas"
            run.font.size = Pt(9)
            run.font.color.rgb = RGBColor.from_string(color if color else C_TEXT)
            _preserve_space(run)
    doc.add_paragraph()


# ---------------------------------------------------------------------------
# Hauptparser
# ---------------------------------------------------------------------------
def build(md_path, docx_path):
    lines = md_path.read_text(encoding="utf-8").splitlines()
    doc = Document()

    normal = doc.styles['Normal']
    normal.font.size = Pt(11)

    in_code = False
    code_buf = []
    code_lang = ""
    table_rows = []

    def flush_table():
        if table_rows:
            add_table(doc, table_rows)
            table_rows.clear()

    for raw in lines:
        line = raw

        quoted = False
        if QUOTE_RE.match(line):
            quoted = True
            line = QUOTE_RE.sub('', line, count=1)

        # Codeblock-Grenzen
        if line.strip().startswith("```"):
            if not in_code:
                in_code = True
                code_lang = line.strip()[3:].strip()
                code_buf = []
            else:
                in_code = False
                add_code_block(doc, code_buf, code_lang)
                code_buf = []
            continue
        if in_code:
            code_buf.append(line)
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

        if line.strip() == "":
            doc.add_paragraph()
            continue

        if RULE_RE.match(line):
            para = doc.add_paragraph()
            add_border(para, 'bottom')
            continue

        m = HEADING_RE.match(line)
        if m:
            doc.add_heading(m.group(2), level=len(m.group(1)))
            continue

        m = BULLET_RE.match(line)
        if m:
            para = doc.add_paragraph(style='List Bullet')
            add_inline(para, m.group(1))
            if quoted:
                para.paragraph_format.left_indent = Pt(36)
            continue

        para = doc.add_paragraph()
        add_inline(para, line)
        if quoted:
            para.paragraph_format.left_indent = Pt(18)
            add_border(para, 'left', color="8888AA", sz="18", space="8")

    # offene Bloecke schliessen
    if in_code and code_buf:
        add_code_block(doc, code_buf, code_lang)
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
