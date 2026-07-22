# Zählergang PWA

## Voraussetzungen für die Entwicklung

- Node.js LTS
- ein moderner Browser

## Start

```bash
npm install
npm run dev
```

Die lokale Adresse aus dem Terminal im Browser öffnen.

## Produktions-Build

```bash
npm run build
npm run preview
```

## Testablauf

1. Nur mit einer Kopie der Excel-Datei testen.
2. `.xlsm` oder `.xlsx` über „Zählergang öffnen“ auswählen.
3. Prüfen, ob Standorte und Zähler korrekt erscheinen.
4. Einen Wert erfassen.
5. „Excel sichern“ verwenden.
6. Die erzeugte `.xlsx` in Excel öffnen und `Daten_Ausdruck`, Spalten B–D prüfen.

## Hinweis

Die App exportiert bewusst `.xlsx`. VBA-Makros werden nicht benötigt, da die PWA deren Funktion ersetzt.
