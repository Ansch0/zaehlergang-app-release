# Zählergang PWA v0.0.2

Statische, installierbare Web-App ohne Build-Schritt. Sie kann direkt auf GitHub Pages oder einem anderen statischen Webhost veröffentlicht werden.

## Enthalten

- `.xlsx` und `.xlsm` lokal öffnen
- Standorte und 10-Uhr-Ablesungen aus `Zaehlerliste_Sortiert` laden
- Zähler aus `Daten_Ausdruck` zuordnen
- Vormonat, Aktualwert und Info anzeigen
- Filter: Alle / Nur offene / Erledigte
- Werteingabe mit Touch-Ziffernblock
- Warnung bei kleinerem Wert oder mehr als 20 % Erhöhung
- Überschreiben vorhandener Werte
- automatische lokale Zwischenspeicherung
- neuer Zählergang: Aktualwert → Vormonat, Aktualwert/Zeitstempel leeren
- Ausgabe als normale `.xlsx`
- Service Worker für Offline-Nutzung nach dem ersten vollständigen Laden

## Datenschutz

Die gewählte Excel-Datei wird ausschließlich im Browser auf dem Gerät verarbeitet. Die App lädt keine Messdaten auf einen Server hoch.

## Veröffentlichung

Alle Dateien aus diesem Ordner in das Veröffentlichungs-Repository hochladen. Danach GitHub Pages auf den Branch `main` und Ordner `/ (root)` einstellen.

## Test

1. Webadresse einmal mit Internetverbindung öffnen.
2. Excel-Datei auswählen.
3. Standorte und Werte prüfen.
4. Einen Testwert eingeben.
5. `Excel sichern` wählen und die erzeugte Datei in Excel prüfen.
6. Danach Flugmodus aktivieren und die App erneut öffnen, um den Offline-Modus zu prüfen.

## Bekannte Grenze

Safari kann eine bereits ausgewählte Excel-Datei nicht direkt überschreiben. Deshalb wird der laufende Stand automatisch im Browser gespeichert und über `Excel sichern` als neue Datei ausgegeben.
