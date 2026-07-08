---
name: handout-generator
description: Liest alle Quelldateien (PDF, Bilder/Screenshots, Text/Markdown) in einem Ordner und erstellt daraus ein didaktisches Lern-Handout als .docx. Verwenden, wenn der Nutzer Kursunterlagen/Folien/Skripte in ein erklärendes Lern- oder Prüfungs-Handout bzw. eine Lernzusammenfassung umwandeln möchte. Cross-platform (Windows/macOS/Linux), Python-basiert.
---

# Handout-Generator

Wandelt Kursmaterial (PDFs, Folien, Screenshots, Textdateien) in ein **erklärendes Lern-Handout** als `.docx` um. Ziel ist kein Stichwort-Spickzettel, sondern ein Dokument, das die Themen *beibringt*.

Zielordner = der aktuelle Arbeitsordner, sofern der Nutzer nichts anderes nennt.

---

## Schritt 0 – Preflight: Python-Werkzeuge sicherstellen

Ziel: einen funktionierenden Python-Interpreter mit `pypdf` (PDF-Text) und `python-docx` (docx-Erzeugung) bereitstellen. **Idempotent** – nur installieren, was fehlt. **Niemals blockieren** – bei Fehlschlag sauber auf den Fallback (Schritt 2b) wechseln.

**0.1 Interpreter finden.** Nacheinander testen, den ersten nehmen, der *echten Code ausführt* (das filtert automatisch den Windows-„Microsoft-Store-Platzhalter" heraus, weil dieser den Code nicht ausführt):
```bash
for c in python3 python py; do "$c" -c "import sys; print(sys.executable)" && { PY="$c"; break; }; done
```
Merke dir den gefundenen Befehl als `$PY`.

**0.2 Wenn kein Interpreter läuft → einmalig installieren** (nach Betriebssystem; `uname` bzw. Umgebung erkennen):
- **Windows:** `winget install --id Python.Python.3.12 -e --source winget`
- **macOS:** `brew install python`  (falls Homebrew fehlt: kurz auf https://brew.sh hinweisen)
- **Linux (Debian/Ubuntu):** `sudo apt-get update && sudo apt-get install -y python3 python3-pip`
  (Fedora: `sudo dnf install -y python3 python3-pip`; Arch: `sudo pacman -S --noconfirm python python-pip`)

Sage dem Nutzer *vorher* in einem Satz, dass du Python einmalig installierst.

**0.3 Bibliotheken sicherstellen:**
```bash
"$PY" -c "import pypdf, docx" || "$PY" -m pip install --user pypdf python-docx
```
Für PowerPoint-Quellen zusätzlich `python-pptx`. Tipp: `pip install -r "<SKILL_DIR>/requirements.txt"` installiert alles auf einmal.

**0.3b Fallback bei „externally-managed" (PEP 668, modernes Linux/Homebrew).**
Schlägt `pip install --user` mit *„externally managed environment"* fehl, ein **venv im Skill-Ordner** anlegen und dessen Python als `$PY` verwenden:
```bash
"$PY" -m venv "<SKILL_DIR>/.venv"
PY="<SKILL_DIR>/.venv/bin/python"   # Windows: <SKILL_DIR>/.venv/Scripts/python.exe
"$PY" -m pip install -r "<SKILL_DIR>/requirements.txt"
```
Das venv ist in `.gitignore` – es bleibt lokal und wird nicht eingecheckt.

**0.4 Fehlerfall (kein Adminrecht, gesperrtes EDU-/Firmengerät, kein Internet):**
Nicht hängen bleiben, nicht endlos wiederholen. Dem Nutzer klar sagen:
> „Python/Tools konnten nicht installiert werden (Grund: …). Ich nutze stattdessen vorhandene Screenshots/Textdateien und bitte dich ggf., den Folientext zu exportieren."
Dann mit Schritt 2b weitermachen.

---

## Schritt 1 – Quellen sichten
Liste die Dateien im Zielordner (`*.pdf`, `*.png/*.jpg`, `*.pptx`, `*.md/*.txt`). Sehr große PDFs (> ~50 MB, z. B. komplette Lehrbücher) nur bei Bedarf/stichprobenartig verarbeiten, nicht komplett – das kostet unnötig Zeit.

## Schritt 2a – Inhalte extrahieren (Standardweg)
Text mit den beiliegenden Skripten in einen Arbeitsordner schreiben (`<SKILL_DIR>` = Ordner dieser SKILL.md; `<ARBEITSORDNER>` = Scratchpad/Temp, **nicht** der Projektordner):
```bash
"$PY" "<SKILL_DIR>/scripts/extract_pdf.py"  "<QUELLORDNER>" "<ARBEITSORDNER>/txt"
"$PY" "<SKILL_DIR>/scripts/extract_pptx.py" "<QUELLORDNER>" "<ARBEITSORDNER>/txt"   # nur bei .pptx
```
Die `.txt` enthalten `--- Seite N ---`-Marker (hilfreich zum Zitieren).

**Auf die Ausgabe achten – nicht blind alles lesen:**
- `WARN_NO_TEXT` bei einer Datei = PDF ohne Textebene (**Scan**). Für *diese* Datei auf den Bild-Fallback (2b) wechseln.
- Exit-Code `2` = *keine* Datei lieferte nutzbaren Text → komplett auf 2b umsteigen.

**Kontext-Strategie (wichtig bei viel Material!):** Extrahierter Text kann riesig sein (ein Lehrbuch = mehrere 100 000 Zeichen) und sprengt sonst den Kontext. Deshalb **datei-/quellenweise** vorgehen: eine Quelle lesen → deren Kernaussagen als kurze Notizen sichern → nächste Quelle. Erst wenn alle Quellen verdichtet sind, das Handout aus den Notizen komponieren. Bei sehr vielen/großen Dateien je Quelle einen Subagenten zusammenfassen lassen und nur die Zusammenfassungen einsammeln. Niemals mehrere volle Rohtexte gleichzeitig in den Kontext ziehen.

## Schritt 2b – Fallback ohne (brauchbare) Extraktion
- **Bilder/Screenshots** direkt mit dem Read-Tool ansehen (das geht ohne Python).
- **`.md/.txt`** direkt lesen.
- **Gescannte PDFs** (`WARN_NO_TEXT`): Wenn möglich Seiten als Bild ansehen, sonst den Nutzer um eine Textversion bitten.
- Fehlt Text ganz, den Nutzer bitten, ihn zu exportieren – und das Handout aus dem verfügbaren Material + Fachwissen erstellen. Transparent kennzeichnen, worauf es beruht.

## Schritt 3 – Handout schreiben (didaktische Vorlage)
Erstelle `<Ordnername_oder_Thema>_Handout.md`. **Erklären, nicht auflisten.** Jedes Kapitel nach diesem Muster:
1. **„Worum geht es?"** – Thema zuerst in einfacher Alltagssprache.
2. **Das Warum** – wozu man das braucht.
3. **Alltagsvergleich/Analogie** – ein konkretes Bild.
4. **Durchgerechnetes Beispiel** – Code/SQL/Diagramm Schritt für Schritt erklärt.
5. **„Merke"-Kasten** (als `>`-Blockquote) mit Eselsbrücke/Prüfungstipp.

Weitere Regeln:
- Deutsch, per „du", Fließtext statt Stichpunkt-Wüste.
- Fachbegriff beim ersten Mal erklären (deutsch **und** englisch).
- Am Ende ein **Selbsttest** mit ~10–15 Fragen und Antworten zum Abdecken.
- Markdown-Umfang, den der Konverter versteht: Überschriften `#`–`####`, Absätze, Aufzählungen `-`, Tabellen `| … |`, Codeblöcke ``` ```lang ```, `**fett**`, `` `code` ``, Zitate/Merke-Kästen `>`.

### Code richtig einsetzen (wichtig)
Der Konverter rendert Codeblöcke als hervorgehobene **Code-Box** mit Syntax-Highlighting. Damit das greift:
- **Immer die Sprache an den öffnenden Zaun schreiben:** ``` ```java ```, ``` ```python ```, ``` ```sql ```, ``` ```json ```, ``` ```xml ```, ``` ```js ``` (unterstützt: java, python, sql, json, xml/html, js). Ohne Sprache gibt es keine Farben.
- **Korrekte, vollständige, lauffähige** Beispiele schreiben (richtige Syntax, passende Klammern/Semikolons). Lieber ein kurzer, *fehlerfreier* Ausschnitt als ein langer.
- **Kurz und fokussiert:** nur das zeigen, worum es geht (5–20 Zeilen). Große Dumps vermeiden.
- **Konsistent mit Leerzeichen einrücken** (keine Tabs); der Konverter erhält Einrückung und Leerzeilen.
- **Erklärung mitliefern:** ein, zwei Sätze *vor* dem Block (was passiert hier?) und ggf. **Kommentare im Code** (`//`, `#`, `--` werden grau hervorgehoben).
- **Im Fließtext** einzelne Bezeichner, Schlüsselwörter, Datei- oder Wertangaben als `` `inline code` `` schreiben – das wird ebenfalls hervorgehoben.
- Ausgaben/Konsolentexte mit ``` ```text ``` kennzeichnen (kein Highlighting, nur Monospace).

## Schritt 4 – .docx erzeugen
```bash
"$PY" "<SKILL_DIR>/scripts/build_docx.py" "<Thema>_Handout.md" "<Thema>_Handout.docx"
```
Das Skript baut eine echte `.docx` – **ohne** Microsoft Word – mit professionellem Layout: Inhaltsverzeichnis, Seitenzahlen (Fußzeile), gestylte Tabellen (blauer Kopf, Zebra-Streifen), Merke-Kästen und **dunklen Code-Boxen mit Syntax-Highlighting**.

Optionen:
- `--light` → helles Code-Theme statt dunkel. **Frag den Nutzer bzw. wähle sinnvoll:** dunkel = am Bildschirm lesen; hell = zum **Ausdrucken** (spart Toner, besser auf S/W-Druckern).
- `--no-toc` → ohne Inhaltsverzeichnis.

Falls die Zieldatei in Word geöffnet ist, schlägt das Speichern mit *Permission denied* fehl → unter anderem Namen speichern oder den Nutzer bitten, die Datei zu schließen.

## Schritt 5 – Abschluss
Dem Nutzer den Pfad zur `.docx` nennen, kurz den Aufbau zusammenfassen und offen kennzeichnen, falls Teile auf Fallback-Material beruhen. Anbieten, einzelne Kapitel zu vertiefen oder Übungsaufgaben zu ergänzen.

---

### Hinweise
- Die Skripte sind reine Standard-Python-Skripte (`pypdf`, `python-docx`, optional `python-pptx`) – identisch auf Windows, macOS, Linux; Abhängigkeiten in `requirements.txt`.
- Temporäre Dateien (extrahierter Text) gehören ins Temp-/Scratchpad-Verzeichnis, nicht in den Projektordner.
- Installationsschritte sind Systemeingriffe: einmalig, transparent, mit Fallback.
- Das Inhaltsverzeichnis füllt Word beim Öffnen selbst (Felder werden automatisch aktualisiert; ggf. „Ja" bestätigen oder F9).
