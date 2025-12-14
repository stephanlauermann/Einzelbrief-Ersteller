CHANGELOG – Einzelbrief-Ersteller

Version 1.1 – 2025-12-14
------------------------------------------------------------
NEU
- OSM / Nominatim Adress-Autovervollständigung (Debounce + Cache)
- PDF-Export über ReportLab
- ODT-Export über LibreOffice (soffice)
- Zusätzlicher Export-Button für DOCX, PDF und ODT
- Falzmarke (Drittel-Falz) und Lochmarke (A4 Mitte) links im Dokument
- Einheitliche Einrückung von Betreff, Anrede, Text, Grußformel und Signatur
  auf Höhe des Fensterkuvert-Inhalts

VERBESSERT
- Fensterkuvert-Positionierung stabil über mm-basierte Tabellenstruktur
- DIN-ähnliches Layout für Amts- und Behördenbriefe
- Saubere Fehlerbehandlung beim Export (DOCX / PDF / ODT)
- Robuste Schriftbehandlung (Arial, Fallback bei PDF)
- GUI klarer und übersichtlicher gestaltet

GEÄNDERT
- Serienbrief-Funktion entfernt (nur Einzelbrief)
- Rechter Absenderblock vollständig entfernt
- Fußzeile vollständig entfernt
- Export erfolgt ohne externe Word-Vorlage

FIXES
- Doppeltes „Hallo,“ im Text verhindert
- Fehler bei VML-Namespace für Falz-/Lochmarken behoben
- Stabilere Druckposition der Falz- und Lochmarken
- Verbesserte Adresszuordnung bei OSM-Ergebnissen

HINWEISE
- Für ODT-Export ist LibreOffice erforderlich
- Für PDF-Export wird das Python-Modul `reportlab` benötigt
- Für DOCX-Export wird das Python-Modul `python-docx` benötigt
- Fenster- und Markenpositionen sind mm-genau im Code konfigurierbar

------------------------------------------------------------

Version 1.0 – 2025-12-12
------------------------------------------------------------
- Erste stabile Version des Einzelbrief-Erstellers
- GUI-basierte Einzelbrief-Erstellung
- DOCX-Export im Fensterkuvert-Layout
- Absenderzeile im Fenster (klein), Empfängerblock korrekt positioniert
- Datumszeile rechts, Betreff, Text, Grußformel, Signatur
- Arial als Standardschrift
