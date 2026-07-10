# Media Assets Library (SPFx)

Eine SharePoint Framework (SPFx) Anwendung zur Verwaltung von Medieninhalten in einer SharePoint-Dokumentbibliothek.

## Funktionen

- Upload von Bildern, Videos, Audio- und Dokumentdateien
- Vorschau von Medien direkt im Browser
- Bearbeitung von Metadaten
- Bucket-/Ordnerverwaltung
- Suche und Filter
- Download von Dateien
- Löschen von Dateien
- SharePoint REST API Integration
- Automatische Format-Erkennung (Bild, Video, Audio, Dokument)

---

# Voraussetzungen

Vor der Installation müssen folgende Komponenten installiert sein:

- Node.js (zur SPFx-Version passend)
- npm
- Gulp CLI
- Zugriff auf die SharePoint-Zielsite

Versionen prüfen:

```bash
node --version
npm --version
gulp --version
```

Falls Gulp noch nicht installiert ist:

```bash
npm install -g gulp-cli
```

---

# Projekt übernehmen

## 1. Repository forken

Auf GitHub das Repository öffnen und oben rechts auf **Fork** klicken.

Dadurch wird eine eigene Kopie des Projekts im eigenen GitHub-Account erstellt.

---

## 2. Repository klonen

```bash
git clone https://github.com/<github-account>/MediaAssetsLibrary.git
```

Anschließend in das Projektverzeichnis wechseln:

```bash
cd MediaAssetsLibrary
```

---

## 3. Abhängigkeiten installieren

Alle Projektabhängigkeiten installieren:

```bash
npm install
```

Nach erfolgreicher Installation sollte eine Meldung ähnlich der folgenden erscheinen:

```text
up to date, audited xxxx packages
```

Hinweise zu Sicherheitswarnungen können zunächst ignoriert werden, sofern das Projekt erfolgreich startet.

---

## 4. SPFx Entwicklerzertifikat installieren

Für die lokale Entwicklung benötigt SharePoint Framework ein vertrauenswürdiges HTTPS-Zertifikat.

Im Projektordner ausführen:

```bash
gulp trust-dev-cert
```

Falls bereits ein altes oder fehlerhaftes Zertifikat vorhanden ist:

```bash
gulp untrust-dev-cert
gulp trust-dev-cert
```

---

# Wichtige Konfiguration

## SharePoint Workbench URL anpassen

Vor dem ersten Start muss die Datei

```text
config/serve.json
```

angepasst werden.

Aktueller Inhalt:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/spfx-serve.schema.json",
  "port": 4321,
  "https": true,
  "initialPage": "https://caritasberlin.sharepoint.com/sites/Medien_dev/_layouts/15/workbench.aspx"
}
```

### Wichtig

Der Wert von `initialPage` muss auf die SharePoint-Umgebung angepasst werden, in der das WebPart entwickelt oder getestet werden soll.

Beispiel:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/spfx-serve.schema.json",
  "port": 4321,
  "https": true,
  "initialPage": "https://tenant.sharepoint.com/sites/meine-site/_layouts/15/workbench.aspx"
}
```

### Beispiel Workbench URL

```text
https://tenant.sharepoint.com/sites/meine-site/_layouts/15/workbench.aspx
```

---

## Bibliotheksname prüfen

Im Code wird aktuell die SharePoint-Dokumentbibliothek

```typescript
private readonly libraryName = "Medienbibliothek";
```

verwendet.

Sollte die Bibliothek in der Zielumgebung anders heißen, muss dieser Wert angepasst werden.

Datei:

```text
src/webparts/mediaAssetsLib/components/MediaAssetsLib.tsx
```

Beispiel:

```typescript
private readonly libraryName = "MediaLibrary";
```

---

# Projekt starten

Nach erfolgreicher Installation und Konfiguration:

```bash
gulp serve
```

Der Build-Prozess startet anschließend lokal.

In der Konsole erscheint eine Ausgabe ähnlich zu:

```text
Build target: DEBUG

Starting 'serve'...

To load your scripts, use this query string:

?debug=true&noredir=true&debugManifestsFile=https://localhost:4321/temp/manifests.js
```

Danach öffnet sich die konfigurierte SharePoint Workbench.

---

# Ordnerstruktur

```text
MediaAssetsLibrary
│
├── config
│   └── serve.json
│
├── src
│   └── webparts
│       └── mediaAssetsLib
│           ├── MediaAssetsLib.tsx
│           ├── FileCard.tsx
│           ├── UploadModal.tsx
│           ├── EditModal.tsx
│           ├── PreviewModal.tsx
│           └── FilterBar.tsx
│
├── gulpfile.js
├── package.json
└── README.md
```

---

# Deployment

Produktionspaket erstellen:

```bash
gulp bundle --ship
gulp package-solution --ship
```

Das fertige SharePoint-Paket befindet sich anschließend unter:

```text
sharepoint/solution/
```

Beispiel:

```text
sharepoint/solution/media-assets-library.sppkg
```

Dieses Paket kann anschließend in den SharePoint App Catalog hochgeladen werden.

---

# Häufige Fehler

## No development certificate found

Fehlermeldung:

```text
No development certificate found.
Generate a new certificate manually...
```

Lösung:

```bash
gulp trust-dev-cert
```

---

## gulp wird nicht erkannt

Fehlermeldung:

```text
gulp is not recognized as an internal or external command
```

Lösung:

```bash
npm install -g gulp-cli
```

---

## npm Pakete fehlen

Fehlermeldungen wie:

```text
Cannot find module ...
```

Lösung:

```bash
npm install
```

---

## Port 4321 ist belegt

In der Datei

```text
config/serve.json
```

einen anderen Port vergeben:

```json
{
  "port": 4322
}
```

---

# Quick Start

```bash
# Repository klonen
git clone <repository-url>

# Projekt öffnen
cd MediaAssetsLibrary

# Abhängigkeiten installieren
npm install

# Entwicklerzertifikat installieren
gulp trust-dev-cert

# config/serve.json anpassen

# Entwicklungsserver starten
gulp serve
```

---

# Ansprechpartner

Bei Fragen zur Architektur, den verwendeten SharePoint-Listenfeldern oder der Deployment-Konfiguration bitte einen Projektverantwortlichen kontaktieren.

---

**Media Assets Library**
SharePoint Framework (SPFx) Medienverwaltung für die Caritas Berlin.
