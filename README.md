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

---

# Verwendung in Microsoft Teams

Die Anwendung kann zusätzlich innerhalb von Microsoft Teams genutzt werden.

Da Teams auf die zugehörige SharePoint-Site zugreift, gelten dieselben Berechtigungen wie in SharePoint.

Kanal öffnen
→ +
→ SharePoint

oder

- → Website

oder

- → SharePoint-Seite

Beispiel

## https://caritasberlin.sharepoint.com/sites/Medien_dev/SitePages/Medienverwaltung.aspx

# Bereitstellung in Microsoft Teams

Die Media Assets Library unterstützt :Die Media Assets Library unterstützt Microsoft Teams nativ und kann sowohl als Teams-Tab als auch als persönliche Teams-App verwendet werden.

```text
Make this solution available to all sites in the organization
```

aktivieren, um die Lösung tenantweit bereitzustellen.

---

## 2. Synchronisierung mit Microsoft Teams

Da das WebPart die Hosts

```json
"supportedHosts": [
  "SharePointWebPart",
  "TeamsPersonalApp",
  "TeamsTab",
  "SharePointFullPage"
]
```

unterstützt, wird die Lösung automatisch für Microsoft Teams bereitgestellt.

Die Synchronisierung kann einige Minuten dauern.

---

## 3. Teams Admin Center prüfen

Im Teams Admin Center:

```text
Teams Apps
→ Manage Apps
```

nach

```text
Media Assets Library
```

suchen.

Falls erforderlich, die App auf:

```text
Allowed
```

setzen.

---

## 4. Teams-App installieren

In Microsoft Teams:

```text
Apps
→ Media Assets Library
→ Hinzufügen
```

Die Anwendung kann anschließend als persönliche App oder als Registerkarte in einem Team verwendet werden.

---

## 5. Als Registerkarte in einem Kanal hinzufügen

Gewünschtes Team öffnen:

```text
Team
→ Kanal
→ +
```

Anschließend:

```text
Media Assets Library
```

auswählen und speichern.

Danach erscheint die Anwendung als eigener Tab im Kanal.

---

## 6. Anwendung für Benutzer anheften (optional)

Im Teams Admin Center:

```text
Teams Apps
→ Setup Policies
→ Pinned Apps
```

Die App hinzufügen.

Dadurch erscheint die Anwendung automatisch bei den Benutzern in der linken Teams-Navigation.

---

# Zugriff auf die Medienbibliothek

Die Anwendung verwendet automatisch die SharePoint-Site, auf der sie installiert bzw. eingebunden wurde.

Der Site-Kontext wird durch SPFx bereitgestellt und muss nicht separat in Teams konfiguriert werden.

Beispiel:

```text
https://caritasberlin.sharepoint.com/sites/Medien_dev
```

Im Code ist standardmäßig folgende Bibliothek hinterlegt:

```typescript
private readonly libraryName = "Medienbibliothek";
```

Daraus ergibt sich automatisch:

```text
https://caritasberlin.sharepoint.com/sites/Medien_dev/Medienbibliothek
```

---

# Änderung der Zielbibliothek

Soll eine andere Bibliothek verwendet werden, muss der Wert angepasst werden:

Datei:

```text
src/webparts/mediaAssetsLib/components/MediaAssetsLib.tsx
```

Beispiel:

```typescript
private readonly libraryName = "MediaLibrary";
```

Danach muss die Lösung neu gebaut und erneut bereitgestellt werden:

```bash
gulp bundle --ship
gulp package-solution --ship
```

---

# Berechtigungen

Benutzer benötigen:

```text
- Zugriff auf Microsoft Teams
- Zugriff auf die zugrunde liegende SharePoint-Site
- Berechtigungen auf die Medienbibliothek
```

Für Upload, Bearbeitung und Löschen von Dateien wird mindestens empfohlen:

```text
Mitwirken (Contribute)
```

Für rein lesenden Zugriff genügt:

```text
Lesen (Read)
```

Die Anwendung verwendet die Berechtigungen des aktuell angemeldeten Benutzers und benötigt keine zusätzlichen Microsoft Graph Application Permissions.

## Voraussetzungen

Vor der Bereitstellung muss die SPFx-Lösung erfolgreich im SharePoint App Catalog installiert sein.

Produktionspaket erstellen:

```bash
gulp bundle --ship
gulp package-solution --ship
```

Anschließend die Datei

```text
sharepoint/solution/media-assets-library.sppkg
```

in den SharePoint App Catalog hochladen und deployen.

---

## 1. SharePoint App Catalog

Im SharePoint Admin Center den App Catalog öffnen:

Beispiel:

https://caritasberlin.sharepoint.com/sites/appcatalog/SitePages/Homepage.aspx

```text
Apps for SharePoint
```

Die Datei

```text
media-assets-library.sppkg
```

hochladen.

Im folgenden Dialog:

```text
Deploy
```

auswählen.

## Voraussetzungen

Benutzer benötigen:

- Zugriff auf das entsprechende Team
- Zugriff auf die zugrunde liegende SharePoint Site
- Zugriff auf die verwendete Dokumentbibliothek

---

## Benötigte SharePoint-Berechtigungen

### Nur Lesen

Für die Anzeige von Dateien genügt:

```text
Lesen (Read)
```

---

### Dateien verwalten

Für folgende Funktionen:

- Upload
- Download
- Bearbeitung von Metadaten
- Löschen von Dateien
- Ordnerverwaltung

wird mindestens benötigt:

```text
Mitwirken (Contribute)
```

---

## Dokumentbibliothek

Die Anwendung verwendet standardmäßig die Bibliothek:

```typescript
private readonly libraryName = "Medienbibliothek";
```

Benutzer benötigen entsprechende Berechtigungen auf dieser Bibliothek.

Wird ein anderer Bibliotheksname verwendet, muss dieser im Projekt angepasst werden.

Datei:

```text
src/webparts/mediaAssetsLib/components/MediaAssetsLib.tsx
```

---

## API-Berechtigungen

Die Anwendung verwendet die SharePoint REST API im Kontext des aktuell angemeldeten Benutzers.

Es werden daher keine zusätzlichen Microsoft Graph Application Permissions benötigt.

Die effektiven Berechtigungen ergeben sich ausschließlich aus den bestehenden SharePoint-Berechtigungen des jeweiligen Benutzers.

```text
Delegated User Permissions
```

---

## Teams Admin Center

Falls die Anwendung in Teams nicht sichtbar ist, sollte geprüft werden, ob sie im Teams Admin Center verfügbar ist:

```text
Teams Admin Center
→ Teams Apps
→ Manage Apps
```

Dort sicherstellen, dass die Anwendung:

```text
Allowed
```

bzw.

```text
Published
```

ist.
``

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
