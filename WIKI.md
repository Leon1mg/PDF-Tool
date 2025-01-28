Hier ist der gesamte Inhalt ausschließlich im Markdown-Format:

```markdown
# PDF-Tool Dokumentation

Willkommen zur Dokumentation des **PDF-Tools**. Dieses Tool wurde mit Python und Bibliotheken wie `tkinter`, `PyPDF2`, `Pillow`, `fitz (PyMuPDF)` und anderen erstellt. Es bietet eine benutzerfreundliche Oberfläche zur Bearbeitung von PDF-Dateien und anderen Dokumentformaten. Die Hauptfunktionen umfassen PDF-Zusammenführung, -Verschlüsselung, -Seitenlöschung und Dateikonvertierung.

---

## Inhaltsverzeichnis
- [Installation](#installation)
- [Projektstruktur](#projektstruktur)
- [Hauptfunktionen](#hauptfunktionen)
  - [Startseite](#startseite)
  - [PDFs zusammenfügen](#pdfs-zusammenfügen)
  - [PDF-Seiten löschen](#pdf-seiten-löschen)
  - [PDF verschlüsseln und entschlüsseln](#pdf-verschlüsseln-und-entschlüsseln)
  - [Dateien in PDF konvertieren](#dateien-in-pdf-konvertieren)
- [Menüleiste](#menüleiste)
- [Erweiterungsmöglichkeiten](#erweiterungsmöglichkeiten)
- [Fehlerbehebung](#fehlerbehebung)
- [Lizenz](#lizenz)

---

## Installation

### Voraussetzungen
- Python 3.x
- Folgende Python-Bibliotheken:
  - `tkinter`
  - `Pillow`
  - `PyPDF2`
  - `fitz` (PyMuPDF)
  - `pywin32`

### Schritte
1. Klonen oder laden Sie das Projekt herunter.
2. Installieren Sie die benötigten Bibliotheken:
   ```bash
   pip install pillow pymupdf pywin32 PyPDF2
   ```
3. Platzieren Sie die Bilder für die Buttons (z. B. `kombinieren.png`, `löschen.png`, usw.) im Ordner `Images/`.
4. Führen Sie das Skript aus:
   ```bash
   python main.py
   ```

---

## Projektstruktur

Das Projekt besteht aus einer zentralen Datei `main.py`, die die gesamte GUI und die Funktionalität definiert. Die wichtigsten Abschnitte der Anwendung sind:

- **Menüleiste**: Globale Navigation und Hilfsoptionen.
- **Funktionale Seiten**: Separate Module für verschiedene Aufgaben (z. B. Zusammenfügen, Verschlüsselung).
- **UI-Komponenten**: Tkinter-Widgets für die Benutzeroberfläche.

---

## Hauptfunktionen

### Startseite
Die Startseite ist der zentrale Einstiegspunkt. Sie enthält große, klickbare Buttons, die zu den verschiedenen Funktionen führen:
- PDFs zusammenfügen
- PDF-Seiten löschen
- PDF verschlüsseln
- Dateien in PDF konvertieren

---

### PDFs zusammenfügen

#### Beschreibung
Mit dieser Funktion können mehrere PDF-Dateien zu einer einzigen Datei kombiniert werden.

#### Workflow
1. Klicken Sie auf **PDFs zusammenfügen**.
2. Wählen Sie die Dateien aus, die zusammengefügt werden sollen.
3. Überprüfen Sie die Dateiliste in der GUI.
4. Speichern Sie die kombinierte Datei über den Button **PDFs zusammenfügen**.

#### Code-Details
- Verwendet `PdfMerger` aus der Bibliothek `PyPDF2`.
- Unterstützt mehrere Dateien.
- Prüft, ob mindestens zwei PDFs hinzugefügt wurden, bevor der Zusammenführungsprozess beginnt.

---

### PDF-Seiten löschen

#### Beschreibung
Löscht ausgewählte Seiten aus einer PDF-Datei.

#### Workflow
1. Laden Sie eine PDF-Datei.
2. Navigieren Sie mit den **Weiter** und **Zurück**-Buttons durch die Seiten.
3. Geben Sie die Seitenzahlen ein, die gelöscht werden sollen (z. B. `1,3,5`).
4. Speichern Sie die bearbeitete Datei.

#### Code-Details
- Verwendet `fitz` (PyMuPDF), um die Seiten als Vorschau anzuzeigen.
- Löscht die Seiten basierend auf Benutzerangaben und erstellt eine neue PDF-Datei.

---

### PDF verschlüsseln und entschlüsseln

#### Beschreibung
Schützt oder entschlüsselt PDF-Dateien mit einem Passwort.

#### Workflow
1. Wählen Sie die Funktion **PDF verschlüsseln** oder **PDF entschlüsseln**.
2. Laden Sie die Datei.
3. Geben Sie ein Passwort ein.
4. Speichern Sie die bearbeitete Datei.

#### Code-Details
- Verschlüsselung erfolgt über die `PdfWriter`-Klasse von `PyPDF2`.
- Entschlüsselung überprüft, ob die PDF passwortgeschützt ist, bevor der Prozess gestartet wird.

---

### Dateien in PDF konvertieren

#### Beschreibung
Konvertiert gängige Dokumentformate (Word, Excel, PowerPoint, Bilder) in PDF.

#### Unterstützte Formate
- **Dokumente**: `.doc`, `.docx`
- **Tabellen**: `.xls`, `.xlsx`
- **Präsentationen**: `.ppt`, `.pptx`
- **Bilder**: `.png`, `.jpeg`, `.jpg`

#### Workflow
1. Wählen Sie eine Datei aus den unterstützten Formaten aus.
2. Speichern Sie die konvertierte PDF-Datei.

#### Code-Details
- Verwendet `pywin32`, um Office-Dokumente in PDF zu konvertieren.
- Bilder werden mit `Pillow` verarbeitet.

---

## Menüleiste

Die Menüleiste bietet folgende Optionen:
- **File**: Anwendung beenden.
- **View**: Wechsel zwischen Vollbild- und Fenstermodus.
- **Help**: Zeigt Informationen zur Anwendung und den Link zur Dokumentation.

---

## Erweiterungsmöglichkeiten

- **Zusätzliche Funktionen**: Weitere Tools für PDFs, wie Kommentarfunktionen oder PDF-Splitting.
- **Mehrsprachigkeit**: Übersetzungen für andere Sprachen.
- **Cloud-Unterstützung**: Dateien direkt in Cloud-Speichern wie Google Drive oder OneDrive bearbeiten.

---

## Fehlerbehebung

- **Fehler beim Konvertieren**: Stellen Sie sicher, dass Microsoft Office installiert ist, wenn Sie `.docx`, `.xls` oder `.pptx` konvertieren möchten.
- **Fehler bei der Verschlüsselung**: Überprüfen Sie, ob die PDF-Datei gültig und nicht beschädigt ist.
- **GUI-Probleme**: Stellen Sie sicher, dass alle benötigten Bilddateien im Ordner `Images` vorhanden sind.

---

## Lizenz

Dieses Tool ist unter der **MIT-Lizenz** lizenziert. Sie dürfen den Code verwenden, modifizieren und verbreiten, solange die ursprünglichen Lizenzhinweise erhalten bleiben.
```
