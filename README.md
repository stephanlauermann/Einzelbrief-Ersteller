# README – Einzelbrief-Ersteller (Python)

GUI-basierter **Einzelbrief-Ersteller** mit **DOCX-Export** (DIN-ähnlicher Fensterbrief), **Arial**, **Falz- & Lochmarken**.

---

## Übersicht

Dieses Programm ist ein **GUI-basierter Einzelbrief-Ersteller**. Es erzeugt **Word-Dokumente (DOCX)** in einer festen Struktur mit:

- Fensterkuvert-Adresse (Absender klein + Empfänger im Fenster)
- Datum rechts, Betreff, Anrede, Text, Grußformel, Signatur
- **Falzmarke** und **Lochmarke** links
- Schrift: **Arial**

---

## Voraussetzungen

### 1) Python

- **Python ≥ 3.10** empfohlen
- Getestet mit Python **3.10 / 3.11 / 3.12**

Download: https://www.python.org/downloads/

**Windows-Hinweis:** Bei der Installation bitte **„Add Python to PATH“** aktivieren.

Prüfen:
```bash
python --version
```

---

### 2) Benötigte Python-Pakete

Das Programm benötigt **ein externes Paket**:

```bash
pip install python-docx
```

Verwendete Standardbibliotheken (bereits enthalten):

- tkinter (GUI)
- datetime
- os

Test:
```bash
python -c "import docx; print('python-docx OK')"
```

---

## Projektstruktur (empfohlen)

```text
einzelbrief/
│
├─ einzelbrief.py
├─ README.md
└─ (optional) default.docx   ← nur als Referenz, nicht notwendig
```

---

## Programm starten

### Windows
```bash
python einzelbrief.py
```

### macOS / Linux
```bash
python3 einzelbrief.py
```

---

## Ausgabe

Das Programm exportiert **DOCX-Dateien**, kompatibel mit:

- Microsoft Word
- LibreOffice Writer
- OnlyOffice

---

## Hinweise

- Falz- und Lochmarken sind als **VML-Objekte im Header** eingebettet.
- Darstellung kann je nach Zoom leicht variieren – **Druck ist stabil**.
- Positionen sind **mm-genau im Skript einstellbar**.

**Typische Anpassungen:**
- „Fenster 3 mm zu tief“ → entsprechenden mm-Wert anpassen
- „Alles 2 mm weiter rechts“ → linken Offset erhöhen

---

## Optional / Ausblick

Aktuell **nicht enthalten**, aber möglich:

- PDF-Export (z.B. via LibreOffice oder ReportLab)
- Adress-Autovervollständigung (OpenStreetMap / Photon)
- Serienbrief-Funktion

---

## Lizenz / Nutzung

- Private und kommerzielle Nutzung erlaubt
- Keine Garantie oder Haftung

---

## Autor

**Stephan Lauermann**  , AI assisted by ChatGPT

Lauermann IT

