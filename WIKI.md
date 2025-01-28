Hier ist ein Vorschlag für ein umfassendes Wiki zu deinem Python-PDF-Tool. Es deckt alle wichtigen Funktionen und Details ab, um den Benutzern das Verständnis und die Bedienung zu erleichtern:

---

# 📚 **PDF-Tool Wiki**

## Inhaltsverzeichnis

1. [Einführung](#einführung)  
2. [Hauptmerkmale](#hauptmerkmale)  
3. [Installation](#installation)  
4. [Anwendung starten](#anwendung-starten)  
5. [Funktionen](#funktionen)  
    - [Startseite](#startseite)  
    - [PDFs zusammenfügen](#pdfs-zusammenfügen)  
    - [PDF-Seiten löschen](#pdf-seiten-löschen)  
    - [PDF verschlüsseln und entschlüsseln](#pdf-verschlüsseln-und-entschlüsseln)  
    - [Dateikonvertierung nach PDF](#dateikonvertierung-nach-pdf)  
6. [Fehlerbehebung](#fehlerbehebung)  
7. [FAQs](#faqs)  

---

## Einführung

Das **PDF-Tool** ist ein vielseitiges, benutzerfreundliches Desktop-Tool, mit dem Sie gängige Aufgaben im Umgang mit PDFs erledigen können. Die intuitive Oberfläche ermöglicht es, PDFs zu kombinieren, Seiten zu löschen, PDFs zu verschlüsseln oder zu entschlüsseln sowie Dateien in PDF zu konvertieren. 

---

## Hauptmerkmale

- PDFs zusammenfügen: Kombinieren Sie mehrere PDFs in einer einzigen Datei.
- Seiten löschen: Entfernen Sie gezielt Seiten aus einer PDF.
- Verschlüsseln und Entschlüsseln: Schützen Sie PDFs mit einem Passwort oder entfernen Sie bestehende Passwörter.
- Konvertierung: Konvertieren Sie Dateien wie Word-, Excel-, PowerPoint- und Bilddateien in PDF.
- Einfache Navigation: Eine klar strukturierte Benutzeroberfläche mit ansprechenden Grafiken.
- Unterstützung mehrerer Formate: Unterstützt `.doc`, `.docx`, `.xls`, `.xlsx`, `.ppt`, `.pptx`, `.png`, `.jpg` und `.jpeg`.

---

## Installation

### Voraussetzungen

1. **Python-Version**: Python 3.7 oder höher.  
2. **Zusätzliche Bibliotheken**:
   - `tkinter`  
   - `Pillow`  
   - `PyMuPDF`  
   - `PyPDF2`  
   - `pywin32`  

   Installieren Sie diese mit folgendem Befehl:
   ```bash
   pip install tk Pillow PyMuPDF PyPDF2 pywin32
   ```

3. **Windows-Nutzer**: Stellen Sie sicher, dass Microsoft Word, Excel und PowerPoint installiert sind, um die Konvertierungsfunktion nutzen zu können.

### Projektdateien

1. Laden Sie die Projektdateien herunter und stellen Sie sicher, dass die folgende Ordnerstruktur vorhanden ist:
   ```
   PDF-Tool/
   ├── main.py
   ├── Images/
   │   ├── kombinieren.png
   │   ├── löschen.png
   │   ├── Verschlüsselung.png
   │   ├── Converter.png
   │   └── icon.ico
   ```

2. Starten Sie das Tool, indem Sie den folgenden Befehl in Ihrem Terminal oder Ihrer Kommandozeile ausführen:
   ```bash
   python main.py
   ```

---

## Anwendung starten

Sobald das Tool gestartet ist, sehen Sie die Startseite mit verschiedenen Optionen:  

- **PDFs zusammenfügen**  
- **PDF-Seiten löschen**  
- **PDF verschlüsseln**  
- **Konverter**  

Sie können jede Funktion über die entsprechenden Buttons aufrufen.  

---

## Funktionen

### Startseite

Die Startseite dient als zentrale Navigation für die Hauptfunktionen. Von hier aus können Sie zwischen den verschiedenen Werkzeugen wechseln.  

---

### PDFs zusammenfügen

1. **Funktion aufrufen**: Wählen Sie `PDFs zusammenfügen` auf der Startseite.  
2. **PDF-Dateien hinzufügen**:  
   - Klicken Sie auf `PDFs hinzufügen`, um mehrere PDFs auszuwählen.  
   - Die ausgewählten PDFs werden in einer Liste angezeigt.  
3. **PDFs zusammenfügen**:  
   - Klicken Sie auf `PDFs zusammenfügen`, wählen Sie einen Speicherort und geben Sie einen Namen für die zusammengefügte Datei ein.  
4. **Liste leeren**: Mit dem Button `Liste leeren` können Sie die PDF-Liste zurücksetzen.  

---

### PDF-Seiten löschen

1. **Funktion aufrufen**: Wählen Sie `PDF-Seiten löschen` auf der Startseite.  
2. **PDF laden**: Klicken Sie auf `PDF laden`, um die gewünschte Datei zu öffnen.  
3. **Seitenvorschau**:  
   - Navigieren Sie durch die Seiten mithilfe der Buttons `Zurück` und `Weiter`.  
   - Eine Vorschau der aktuellen Seite wird angezeigt.  
4. **Seiten löschen**:  
   - Geben Sie die Seitenzahlen ein, die Sie löschen möchten (z. B. `1,3,5`).  
   - Klicken Sie auf `Seiten löschen und speichern`, um die bearbeitete PDF zu speichern.  

---

### PDF verschlüsseln und entschlüsseln

1. **Funktion aufrufen**: Wählen Sie `PDF verschlüsseln` auf der Startseite.  
2. **PDF verschlüsseln**:  
   - Wählen Sie die PDF aus, die Sie schützen möchten.  
   - Geben Sie ein Passwort ein und klicken Sie auf `PDF verschlüsseln`.  
3. **PDF entschlüsseln**:  
   - Wählen Sie die PDF aus, die Sie entsperren möchten.  
   - Geben Sie das Passwort ein und klicken Sie auf `PDF entschlüsseln`.  

---

### Dateikonvertierung nach PDF

1. **Funktion aufrufen**: Wählen Sie `Konverter` auf der Startseite.  
2. **Datei auswählen**:  
   - Unterstützte Dateiformate: `.doc`, `.docx`, `.xls`, `.xlsx`, `.ppt`, `.pptx`, `.png`, `.jpg`, `.jpeg`.  
   - Klicken Sie auf `Datei auswählen und konvertieren`, um eine Datei auszuwählen.  
3. **Speichern**: Geben Sie den Speicherort und den Dateinamen für die PDF an.  

---

## Fehlerbehebung

### Häufige Probleme und Lösungen

1. **Fehler: `win32com.client` wird nicht gefunden**  
   - Stellen Sie sicher, dass `pywin32` installiert ist:
     ```bash
     pip install pywin32
     ```

2. **Fehler: `Microsoft Word/Excel/PowerPoint konnte nicht gefunden werden`**  
   - Überprüfen Sie, ob Microsoft Office auf Ihrem Computer installiert ist.  

3. **PDF-Datei kann nicht geladen werden**  
   - Stellen Sie sicher, dass die Datei nicht beschädigt ist und im `.pdf`-Format vorliegt.  

---



Das war's! Viel Spaß beim Nutzen deines PDF-Tools! 🚀
