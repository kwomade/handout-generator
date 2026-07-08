#!/usr/bin/env python3
"""Extrahiert Text aus PDF-Dateien mit pypdf.

Aufruf:
    python extract_pdf.py <pdf-datei-oder-ordner> [ausgabeordner]

- Ohne Ausgabeordner: Text wird auf stdout gedruckt.
- Mit Ausgabeordner: pro PDF eine .txt-Datei (UTF-8); eine Zusammenfassung geht auf stdout.

Exit-Codes: 0 = ok, 1 = falscher Aufruf, 3 = pypdf fehlt.
"""
import sys
import pathlib

try:
    from pypdf import PdfReader
except ImportError:
    sys.stderr.write("PYPDF_MISSING: 'pypdf' ist nicht installiert. "
                     "Bitte 'pip install --user pypdf' ausfuehren.\n")
    sys.exit(3)


def extract(pdf_path):
    reader = PdfReader(str(pdf_path))
    parts = []
    for page in reader.pages:
        parts.append(page.extract_text() or "")
    return "\n".join(parts)


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

    for pdf in pdfs:
        try:
            content = extract(pdf)
        except Exception as exc:  # robuste Einzelfehler, weitermachen
            sys.stderr.write("ERROR\t%s\t%s\n" % (pdf.name, exc))
            continue
        if out_dir is not None:
            target = out_dir / (pdf.stem + ".txt")
            target.write_text(content, encoding="utf-8")
            print("OK\t%s\t%d chars -> %s" % (pdf.name, len(content), target.name))
        else:
            print(content)


if __name__ == "__main__":
    main()
