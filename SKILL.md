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
(Optional für PowerPoint-Quellen zusätzlich `python-pptx`.)

**0.4 Fehlerfall (kein Adminrecht, gesperrtes EDU-/Firmengerät, kein Internet):**
Nicht hängen bleiben, nicht endlos wiederholen. Dem Nutzer klar sagen:
> „Python/Tools konnten nicht installiert werden (Grund: …). Ich nutze stattdessen vorhandene Screenshots/Textdateien und bitte dich ggf., den Folientext zu exportieren."
Dann mit Schritt 2b weitermachen.

---

## Schritt 1 – Quellen sichten
Liste die Dateien im Zielordner (`*.pdf`, `*.png/*.jpg`, `*.pptx`, `*.md/*.txt`). Sehr große PDFs (> ~50 MB, z. B. komplette Lehrbücher) nur bei Bedarf/stichprobenartig verarbeiten, nicht komplett – das kostet unnötig Zeit.

## Schritt 2a – Inhalte extrahieren (Standardweg)
PDF-Text mit dem beiliegenden Skript in einen Arbeitsordner schreiben:
```bash
"$PY" "<SKILL_DIR>/scripts/extract_pdf.py" "<QUELLORDNER>" "<ARBEITSORDNER>/txt"
```
Danach die erzeugten `.txt`-Dateien lesen. `<SKILL_DIR>` ist der Ordner dieser SKILL.md; `<ARBEITSORDNER>` liegt im Scratchpad/Temp, nicht im Projekt.

## Schritt 2b – Fallback ohne PDF-Extraktion
- **Bilder/Screenshots** direkt mit dem Read-Tool ansehen (das geht ohne Python).
- **`.md/.txt`** direkt lesen.
- Fehlt Folientext ganz, den Nutzer bitten, ihn zu exportieren – und das Handout aus dem verfügbaren Material + Fachwissen erstellen. Transparent kennzeichnen, worauf es beruht.

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
- Markdown-Umfang, den der Konverter versteht: Überschriften `#`–`####`, Absätze, Aufzählungen `-`, Tabellen `| … |`, Codeblöcke ``` ``` ```, `**fett**`, `` `code` ``, Zitate/Merke-Kästen `>`.

## Schritt 4 – .docx erzeugen
```bash
"$PY" "<SKILL_DIR>/scripts/build_docx.py" "<Thema>_Handout.md" "<Thema>_Handout.docx"
```
Das Skript baut eine echte `.docx` (Überschriften, Tabellen, Codeblöcke, Merke-Kästen) – **ohne** Microsoft Word.

## Schritt 5 – Abschluss
Dem Nutzer den Pfad zur `.docx` nennen, kurz den Aufbau zusammenfassen und offen kennzeichnen, falls Teile auf Fallback-Material beruhen. Anbieten, einzelne Kapitel zu vertiefen oder Übungsaufgaben zu ergänzen.

---

### Hinweise
- Die Skripte sind reine Standard-Python-Skripte + `pypdf`/`python-docx` – identisch auf Windows, macOS, Linux.
- Temporäre Dateien (extrahierter Text) gehören ins Temp-/Scratchpad-Verzeichnis, nicht in den Projektordner.
- Installationsschritte sind Systemeingriffe: einmalig, transparent, mit Fallback.
