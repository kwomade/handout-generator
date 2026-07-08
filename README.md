<div align="center">

# 📚✨ Handout-Generator

### Aus langweiligen Folien wird ein Lern-Handout, das dir die Themen *wirklich erklärt*.

Ein [Claude Code](https://claude.com/claude-code) **Skill**, der deine Kursunterlagen (PDFs, Folien, Screenshots) einliest und daraus ein didaktisches Lern-Handout als fertige **`.docx`** baut – ganz ohne Microsoft Word.

![Made for students](https://img.shields.io/badge/made%20for-Schüler%20%26%20Azubis-6c47ff?style=for-the-badge)
![Platform](https://img.shields.io/badge/Windows%20·%20macOS%20·%20Linux-informational?style=for-the-badge)
![Python](https://img.shields.io/badge/Python-3.x-3776AB?style=for-the-badge&logo=python&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)

</div>

---

## 🤔 Das Problem

Du kennst das: 15 PDF-Präsentationen, ein 200-Seiten-Buch, und in zwei Wochen ist Prüfung.
Klassische Zusammenfassungen sind oft nur **Stichpunkt-Wüsten** – nett zum Wiederholen, aber sie *erklären* dir nichts. Genau da setzt dieses Tool an.

## 💡 Die Lösung

Der Handout-Generator verwandelt dein Material in ein Dokument, das aufgebaut ist wie ein guter Lehrer erklärt:

> **Worum geht es?** → **Warum ist das wichtig?** → **Ein Alltagsvergleich** → **Ein durchgerechnetes Beispiel** → **🎯 Merke-Kasten** für die Prüfung.

Am Ende gibt's einen **Selbsttest** mit Fragen und Antworten zum Abdecken.

---

## ✨ Features

- 🧠 **Erklärt statt aufzulisten** – Analogien, Schritt-für-Schritt-Beispiele, Merke-Kästen
- 📄 **Echte `.docx`** – saubere Überschriften, Tabellen, Code-Blöcke, farbige Kästen (kein Word nötig!)
- 🗂️ **Frisst fast alles** – PDF, Screenshots/Bilder, Text- und Markdown-Dateien
- 🖥️ **Läuft überall** – Windows, macOS & Linux mit demselben Code
- 🛠️ **Selbst-Setup** – prüft beim ersten Lauf die nötigen Tools und installiert sie bei Bedarf
- 📝 **Selbsttest inklusive** – Prüfungsfragen mit Lösungen am Ende jedes Handouts

---

## 🚀 Installation

Der Handout-Generator ist ein **Skill** für [Claude Code](https://claude.com/claude-code). Lege ihn dort ab, wo Claude Code seine Skills sucht:

```bash
# Repo direkt in den Skills-Ordner klonen
git clone https://github.com/kwomade/handout-generator.git ~/.claude/skills/handout-generator
```

> 💡 **Auf mehreren Rechnern?** Einfach auf jedem Gerät in `~/.claude/skills/` klonen.
> Unter Windows ist das `C:\Users\<DeinName>\.claude\skills\`.

Beim ersten Start richtet der Skill automatisch **Python 3** samt `pypdf` und `python-docx` ein (über `winget` / `brew` / `apt`, je nach System).

---

## 🎓 Benutzung

Öffne Claude Code im Ordner mit deinen Unterlagen und sag einfach:

```
/handout-generator
```

…oder ganz natürlich:

> „Erstell mir aus den Dateien in diesem Ordner ein Lern-Handout.“

Das war's. Nach ein paar Sekunden liegt eine `<Thema>_Handout.docx` im Ordner. 🎉

---

## ⚙️ Wie es funktioniert

```
  deine Dateien           dieser Skill                  dein Handout
 ┌───────────────┐      ┌──────────────────┐          ┌──────────────┐
 │ 📊 Folien.pdf │      │ 1. Text auslesen │          │ 📘 Handout   │
 │ 🖼️ Screens.png│  ─▶  │    (pypdf)       │   ─────▶  │    .docx     │
 │ 📝 Notizen.md │      │ 2. didaktisch    │          │  erklärt +   │
 └───────────────┘      │    aufbereiten   │          │  Selbsttest  │
                        │ 3. .docx bauen   │          └──────────────┘
                        │    (python-docx) │
                        └──────────────────┘
```

| Datei | Aufgabe |
|-------|---------|
| `SKILL.md` | Anleitung für Claude: Ablauf, Handout-Vorlage, Setup-Logik |
| `scripts/extract_pdf.py` | Zieht Text aus PDFs (mit `pypdf`) |
| `scripts/build_docx.py` | Baut aus Markdown eine echte `.docx` (mit `python-docx`) |

Die beiden Skripte sind eigenständig und laufen auch ohne Claude:

```bash
python scripts/extract_pdf.py  "ordner-mit-pdfs"  "ausgabe/txt"
python scripts/build_docx.py   "mein_handout.md"  "mein_handout.docx"
```

> ℹ️ **Windows-Tipp:** Nutze `py` statt `python`, falls `python` nur den Microsoft-Store öffnet.

---

## 🧩 Voraussetzungen

- [Claude Code](https://claude.com/claude-code)
- Python 3.x mit `pypdf` und `python-docx` *(richtet der Skill bei Bedarf selbst ein)*

---

## 🤝 Mitmachen

Ideen, Verbesserungen oder eine schönere Handout-Vorlage? Issues und Pull Requests sind willkommen!

## 📄 Lizenz

MIT – nutze, verändere und teile es frei. Viel Erfolg bei der Prüfung! 🍀

<div align="center">
<sub>Gebaut mit 🤖 <a href="https://claude.com/claude-code">Claude Code</a> · für alle, die lieber verstehen als auswendig lernen.</sub>
</div>
