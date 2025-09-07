# Batterie-Optimierung

Dieses Repository enthält zwei Python-Programme zur wirtschaftlichen Optimierung des Einsatzes einer Batterie auf Basis von Strompreisen und Verbrauchsdaten.

## Programme

### 1. batterie01.py
- **Funktion:** Wandelt Stundenpreise in Viertelstundenpreise um (Interpolation/Aufteilung).
- **Eingabe:** Excel-Datei mit Zeitstempel, Preis (stündlich) und Verbrauch viertelstündlich - H0 Profil.
               Im Excelfile kann man den Jahresverbrauch eingeben - dieser wird dann lt. H0 Profil auf 15 Minuten Werte umgewandelt. 
- **Ausgabe:** Excel-Datei mit Zeitstempel, Preis und Verbrauch im 15-Minuten-Raster.
- **Ziel:** Vorbereitung der Daten für die Optimierung. 

### 2. batterie02.py
- **Funktion:** Optimiert die wirtschaftliche Bewirtschaftung der Batterie auf Basis der Viertelstundenpreise und Viertelstunden-Verbrauchsdaten
- **Features:**  - Siehe oben (Excel-Export, Parameter, Summen, API, etc.)
- **Eingabe:** Excel-Datei mit Viertelstundenpreisen und Verbrauch.
- **Ausgabe:** Optimierte Lade-/Entladestrategie als Excel-Datei - pro Durchlauf wird an der Erbegnisdatei ein Datumsstempel angehängt.
- **Ziel:** Minimierung der Stromkosten durch optimale Steuerung der Batterie.

## Anwendung
1. Excel-Datei mit Preisdaten und Verbrauch bereitstellen.
2. Parameter im Programm anpassen oder per Eingabe setzen.
3. Skript ausführen: Die Lade und Entladezeitpunkte werden als Ergebnisse als Excel-Datei exportiert.

## Hinweise
- Die Ergebnisdateien werden automatisch von der Versionskontrolle ausgeschlossen.
- Für große Zeiträume kann die Optimierung (je nach Rechner) mehrere Minuten dauern.
- Die Programme sind für private/experimentelle Zwecke gedacht und nicht für den produktiven Einsatz getestet.

---
Fragen oder Anregungen? Gerne als Issue im Repository melden!
