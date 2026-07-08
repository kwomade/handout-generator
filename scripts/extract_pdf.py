#!/usr/bin/env python3
"""Extrahiert Text aus PDF-Dateien mit pypdf.

Aufruf:
    python extract_pdf.py <pdf-datei-oder-ordner> [ausgabeordner]

- Ohne Ausgabeordner: Text wird auf stdout gedruckt.
- Mit Ausgabeordner: pro PDF eine .txt-Datei (UTF-8); eine Zusammenfassung geht auf stdout.
- Der Text enthaelt "--- Seite N ---"-Marker fuer die Orientierung.
- PDFs ohne Textebene (Scans) werden mit WARN_NO_TEXT gemeldet.

Exit-Codes: 0 = ok, 1 = falscher Aufruf, 2 = keine Datei lieferte nutzbaren Text,
            3 = pypdf fehlt.
"""
import sys
import pathlib

# UTF-8-Ausgabe erzwingen (Windows-Konsole ist oft cp1252 -> Mojibake)
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except AttributeError:
    pass

try:
    from pypdf import PdfReader
except ImportError:
    sys.stderr.write("PYPDF_MISSING: 'pypdf' ist nicht installiert. "
                     "Bitte 'pip install --user pypdf' ausfuehren.\n")
    sys.exit(3)

# Unter diesem Schnitt (Zeichen pro Seite) gehen wir von einem Scan ohne
# Textebene aus und warnen, statt stillschweigend leeren Text zu liefern.
MIN_CHARS_PER_PAGE = 25


def extract(pdf_path):
    reader = PdfReader(str(pdf_path))
    parts = []
    chars = 0
    for num, page in enumerate(reader.pages, start=1):
        text = page.extract_text() or ""
        chars += len(text.strip())
        parts.append("--- Seite %d ---" % num)
        parts.append(text)
    return "\n".join(parts), len(reader.pages), chars


def main():
    if len(sys.argv) < 2:
        sys.stderr.write("usage: extract_pdf.py <pdf-or-folder> [out-dir]\n")
        sys.exit(1)

    src = pathlib.Path(sys.argv[1])
    out_dir = pathlib.Path(sys.argv[2]) if len(sys.argv) > 2 else None
    if out_dir is not None:
        out_dir.mkdir(parents=True, exist_ok=True)

    if src.is_file():
        pdfs = [src]
    else:
        pdfs = sorted(src.glob("*.pdf"))

    if not pdfs:
        sys.stderr.write("Keine PDF-Dateien gefunden unter: %s\n" % src)
        sys.exit(0)

    usable = 0
    for pdf in pdfs:
        try:
            content, npages, chars = extract(pdf)
        except Exception as exc:  # robuste Einzelfehler, weitermachen
            sys.stderr.write("ERROR\t%s\t%s\n" % (pdf.name, exc))
            continue

        if out_dir is not None:
            target = out_dir / (pdf.stem + ".txt")
            target.write_text(content, encoding="utf-8")

        if chars < MIN_CHARS_PER_PAGE * max(1, npages):
            print("WARN_NO_TEXT\t%s\tnur %d Zeichen auf %d Seiten "
                  "(Scan ohne Textebene? -> Fallback nutzen)" % (pdf.name, chars, npages))
            continue

        usable += 1
        if out_dir is not None:
            print("OK\t%s\t%d Seiten, %d chars -> %s"
                  % (pdf.name, npages, chars, target.name))
        else:
            print(content)

    if usable == 0:
        sys.stderr.write("KEINE nutzbare Textquelle unter den PDFs.\n")
        sys.exit(2)


if __name__ == "__main__":
    main()
