/********************* DATA MODEL **********************
 * SharePoint List Item Struktur
 ******************************************************/

/**
 * Schnittstelle für ein Medienelement in der Bibliothek
 * Repräsentiert ein SharePoint-Listenelement mit Metadaten
 * @interface IMediaItem
 * @property {number} id - Eindeutige ID des Elements in SharePoint
 * @property {string} name - Dateiname des Mediums
 * @property {string} fileRef - Relative URL zur Datei im SharePoint
 * @property {string} [category] - Kategorisierung des Mediums
 * @property {string} [notes] - Zusätzliche Notizen zum Element
 * @property {string} [created] - ISO-Datum der Erstellung
 * @property {string} [uniqueId] - Eindeutige GUID des Elements
 * @property {string[]} [tags] - Array von Tags für Suchfunktion
 * @property {string[]} [bucket] - Array von Bucket-/Ordner-Namen
 * @property {string} [driveId] - OneDrive Drive ID (VroomDriveID)
 * @property {string} [driveItemId] - OneDrive Item ID (VroomItemID)
 * @property {string} [format] - Dateityp (Bild, Video, Audio, Dokument)
 * @property {string} [createdBy] - Name des Erstellers
 * @property {string} [dienst] - Zugeordneter Dienst/Service
 */
export interface IMediaItem {
  id: number;
  uniqueId?: string;
  name: string;
  fileRef: string;
  category?: string;
  notes?: string;
  created?: string;
  modified?: string;
  tags?: string[];
  bucket?: string[];

  driveId?: string;
  driveItemId?: string;

  format?: string;
  createdBy?: string;
  dienst?: string;

  thumbnailUrl?: string;
}

/**
 * SharePoint REST-API Responseschnittstelle für Dateien
 * Repräsentiert die Struktur einer Datei aus der SharePoint REST API
 * @interface ISPFile
 * @property {string} Name - Dateiname
 * @property {string} ServerRelativeUrl - Relative URL der Datei im SharePoint
 * @property {string} TimeCreated - ISO-Zeitstempel der Erstellung
 * @property {Object} ListItemAllFields - Metadaten des Listenelements
 * @property {number} ListItemAllFields.Id - Element-ID
 * @property {string} [ListItemAllFields.Kategorie] - Kategoriefeld
 * @property {string} [ListItemAllFields.Dienste] - Dienstfeld
 * @property {string} [ListItemAllFields.Notizen] - Notizfeld
 * @property {string[]|string} [ListItemAllFields.Tags] - Tags als Array oder String
 * @property {string[]|string} [ListItemAllFields.Bucket] - Buckets als Array oder String
 * @property {string} [ListItemAllFields.Format] - Formattyp
 * @property {Object} [ListItemAllFields.Author] - Autor-Informationen
 * @property {string} [ListItemAllFields.Author.Title] - Autorenname
 * @property {Object} [ListItemAllFields.File] - Datei-Metadaten
 * @property {string} [ListItemAllFields.File.VroomDriveID] - OneDrive Drive ID
 * @property {string} [ListItemAllFields.File.VroomItemID] - OneDrive Item ID
 * @property {string} [ListItemAllFields.Ersteller] - Ersteller-Name
 */
export interface ISPFile {
  Name: string;
  ServerRelativeUrl: string;
  TimeCreated: string;
  TimeLastModified: string;

  UniqueId?: string;

  ListItemAllFields: {
    Id: number;
    Kategorie?: string;
    Dienste?: string;
    Notizen?: string;
    Tags?: string[] | string;
    Bucket?: string[] | string;
    Format?: string;

    Author?: {
      Title: string;
    };

    File?: {
      VroomDriveID?: string;
      VroomItemID?: string;
    };

    Ersteller?: string;
  };
}
