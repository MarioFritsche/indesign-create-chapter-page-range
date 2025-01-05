# InDesign Create Chapter Page Range

Ein InDesign-Skript, das eine einfache Möglichkeit bietet, Inhaltsverzeichnisse (IHV) oder Kapitelübersichten (KÜ) zu erstellen und zu verwalten. Das Skript ist für die Erstellung einer Ebene ausgelegt und bietet flexible Optionen zur Formatierung und Aktualisierung bestehender Verzeichnisse.

![Screenshot Dialogfenster](https://indesign-kalender.de/github/ihv-gitgub.jpg)

## Funktionen

- **Erstellung von Inhaltsverzeichnissen oder Kapitelübersichten**:
  - Möglichkeit, zwischen einem Inhaltsverzeichnis (IHV) oder einer Kapitelübersicht (KÜ) zu wählen.
  - Frei wählbarer Titel für das Verzeichnis.
  - Auswahl des Absatzformats, das für die Erfassung der Einträge verwendet werden soll.
- **Flexible Formatierung**:

  - Individuelle Absatzformate für den Titel und die Einträge des Verzeichnisses.

- **Verwaltung mehrerer Verzeichnisse**:
  - Es können beliebig viele Inhaltsverzeichnisse oder Kapitelübersichten im selben Dokument erstellt werden.
  - Bestehende Verzeichnisse können jederzeit aktualisiert werden.
  - Auswahl bestehender Verzeichnisse im unteren Bereich des Dialogs zur Aktualisierung.

## Erste Schritte

### Installation

1. Lade das Skript `createChapterPageRange.jsx` herunter.
2. Kopiere die Datei in den InDesign-Skript-Ordner:
   - **Windows**: `C:\Benutzer\<Benutzername>\AppData\Roaming\Adobe\InDesign\<Versionsnummer>\Scripts\Scripts Panel`
   - **macOS**: `~/Library/Preferences/Adobe InDesign/<Versionsnummer>/Scripts/Scripts Panel`
3. Öffne InDesign und rufe das Skript über das Skript-Bedienfeld (`Fenster > Hilfsprogramme > Skripte`) auf.

### Anwendung

#### 1. Erstellen eines neuen Inhaltsverzeichnisses oder einer Kapitelübersicht

- **Vorbereitung**:
  - Erstelle beim ersten Anlegen eines Verzeichnisses einen leeren Textrahmen in deinem Dokument und wähle ihn aus.
- **Skript starten**:
  - Führe das Skript `createChapterPageRange.jsx` aus.
  - Wähle aus, ob ein Inhaltsverzeichnis (IHV) oder eine Kapitelübersicht (KÜ) erstellt werden soll.
  - Gib einen Titel für das Verzeichnis an.
  - Wähle das Absatzformat, das für die Einträge genutzt werden soll.
  - Definiere Absatzformate für den Titel und die Einträge.
  - Bestätige die Eingaben, um das Verzeichnis zu erstellen.

#### 2. Aktualisieren eines bestehenden Verzeichnisses

- Starte das Skript.
- Wenn bereits ein Verzeichnis mit diesem Skript angelegt wurde, muss kein Textrahmen ausgewählt werden.
- Wähle im unteren Bereich des Dialogs das Verzeichnis aus, das aktualisiert werden soll.
- Nimm bei Bedarf Änderungen an den Einstellungen vor und bestätige, um die Aktualisierung durchzuführen.

![Screenshot Dialogfenster Aktualisierung bestehende Kapitelübersicht](https://indesign-kalender.de/github/due-github.jpg)

## Einschränkungen

- Das Skript unterstützt aktuell nur Inhaltsverzeichnisse und Kapitelübersichten mit einer Ebene.
- Beim ersten Anlegen eines Verzeichnisses muss ein Textrahmen ausgewählt werden. Für das Aktualisieren oder Hinzufügen weiterer Verzeichnisse ist dies nicht erforderlich.

## Support

Für Feedback erstelle bitte ein [Issue](https://github.com/MarioFritsche/createChapterPageRange/issues) auf GitHub.

## Lizenz

Dieses Projekt steht unter der [MIT-Lizenz](LICENSE).
