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

# SAHREPOINT

- erstelle Medienbibliothek (Vorlage)
- Spalte Format, Type Auswahl (Choice) , Bild , Video, Audio, can add values manually
- Notizen bleibt
- estelle Tags, Auswahl (choice), can add values manually, allow multiple selections
- erstelle Kategorie, Auswahl,erstelle choices:
  Feste, Jubiläen, Kampagne, Demos, Beratung, Portrait, Internes, Sonstige
- estelle Buckets, Auswahl (choice), can add values manually, allow multiple selections
- erstelle Ersteller, single line text
- erstelle Dienste, auswahl (choice), choices:
  Allgemine Soziale Berautng, Integration in Arbeit, Betreuungsverien, Ehrenamt, Kinder, Jugend- und Familienhilfe,
  Frauenhaus, Flucht und Migration, Hospizdienst, Kleiderkammer, Schuldner- und Insolvenzberatung, Schwangerschaftsberatung, Strompar-Check,Suchthilfe, Vormundschaftsverein, Wohnungslosenhilfe, youngcaritas, CFL, Pastoraler Raum, Caritas Gesundheit, Kommunikation, Vorstand

# Voraussetzungen

Vor der Installation müssen folgende Komponenten installiert sein:

- Node.js (zur SPFx-Version passend)
- npm
- Gulp CLI
- Zugriff auf die SharePoint-Zielsite
- install git

Versionen prüfen:

```bash
INSTALL NODE 18!!!!
```

```bash
node --version
npm --version
gulp --version
```

Falls Gulp noch nicht installiert ist:

```bash

npm install -g gulp-cli
```

git installieren

```bash
https://git-scm.com/install/
```

---

# Projekt übernehmen

## 1. Repository forken

Auf GitHub das Repository öffnen und oben rechts auf **Fork** klicken.

Dadurch wird eine eigene Kopie des Projekts im eigenen GitHub-Account erstellt.

---

## 2. Repository klonen

```bash
git clone https://github.com/J4NBI/MediaAssetsLibrary.git
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
(npx) gulp trust-dev-cert
```

Falls bereits ein altes oder fehlerhaftes Zertifikat vorhanden ist:

```bash
(npx) gulp untrust-dev-cert
(npx) gulp trust-dev-cert
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

## Seitenpfad kontolieren

MediaAssestsLibWebPArts.ts

siteUrl:
"https://caritasberlin.sharepoint.com/sites/Medien_dev",

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

```bash
MediaAssetsLibWebPart.manifest.json
```

Version ändern

"title": { "default": "Caritas Media Library V32" },

---

# Projekt starten

Nach erfolgreicher Installation und Konfiguration:

```bash
(npx) gulp serve
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

# Deployment nach SharePoint

## 1. Produktionspaket erstellen

Für das Deployment zunächst das SPFx-Paket erstellen:

```bash
(npx) gulp bundle --ship
(npx) gulp package-solution --ship
```

Das generierte Paket befindet sich anschließend unter:

```text
sharepoint/solution/media-assets-library.sppkg
```

---

## 2. Paket in den SharePoint App Catalog hochladen

Den SharePoint App Catalog öffnen:

```text
https://<tenant>-admin.sharepoint.com
```

Navigation:

```text
More Features
→ Apps
→ Open App Catalog
```

oder direkt:

```text
App Catalog
→ Apps for SharePoint
```

Die Datei

```text
media-assets-library.sppkg
```

hochladen.

Nach dem Upload erscheint ein Bereitstellungsdialog.

Optional kann die Einstellung

```text
Make this solution available to all sites in the organization
```

aktiviert werden, um die Lösung mandantenweit bereitzustellen.

Anschließend auf **Deploy** klicken.

---

## 3. Anwendung auf einer SharePoint Site hinzufügen

Auf der gewünschten SharePoint Site:

```text
Websiteinhalte
→ Neu
→ App
```

Die Anwendung:

```text
Media Assets Library
```

auswählen und hinzufügen.

---

## 4. WebPart auf einer Seite verwenden

Eine moderne SharePoint-Seite öffnen und bearbeiten:

```text
Bearbeiten
→ +
→ Media Assets Library
```

Das WebPart auswählen und die Seite veröffentlichen.

## 5. Push Git

- git add .
- git commit -m "Beschreibung der Änderung"
- git push
