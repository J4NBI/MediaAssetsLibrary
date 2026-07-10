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
interface IMediaItem {
  id: number;
  uniqueId?: string;
  name: string;
  fileRef: string;
  category?: string;
  notes?: string;
  created?: string;
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
interface ISPFile {
  Name: string;
  ServerRelativeUrl: string;
  TimeCreated: string;

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

/********************* STATE ***************************
 * Enthält alle UI Zustände (Filter, Modals, Form Daten)
 ******************************************************/

/**
 * Zustandsschnittstelle für die MediaAssetsLib Komponente
 * Verwaltet alle UI-Zustände, Filter, Modal-Zustände und Formulardaten
 * @interface IMediaAssetsLibState
 * @property {IMediaItem[]} allItems - Alle geladenen Medienelemente
 * @property {IMediaItem[]} visibleItems - Nach Filtern gefilterte Elemente
 * @property {string} searchText - Aktueller Suchtext
 * @property {string} [filterCategory] - Ausgewählte Kategoriefilter
 * @property {number} [filterYear] - Ausgewähltes Filterjahr
 * @property {number} [filterMonth] - Ausgewählter Filtermonat (1-12)
 * @property {string} [filterFormat] - Ausgewählter Formatfilter (Bild, Video, Audio)
 * @property {boolean} isModalOpen - Preview-Modal sichtbar?
 * @property {IMediaItem} [selectedItem] - Aktuell ausgewähltes Element
 * @property {boolean} isEditOpen - Bearbeitungs-Modal sichtbar?
 * @property {string} editName - Name bei Bearbeitung
 * @property {string[]} editTags - Tags bei Bearbeitung
 * @property {string} editCategory - Kategorie bei Bearbeitung
 * @property {string} editFormat - Format bei Bearbeitung
 * @property {string[]} editBucket - Buckets bei Bearbeitung
 * @property {string[]} uploadBucket - Buckets für Upload
 * @property {boolean} isUploadOpen - Upload-Modal sichtbar?
 * @property {string} uploadName - Name für Upload
 * @property {string[]} uploadTags - Tags für Upload
 * @property {string} uploadCategory - Kategorie für Upload
 * @property {string} uploadDienst - Dienst für Upload
 * @property {File[]} [uploadFiles] - Ausgewählte Dateien zum Upload
 * @property {string} [uploadPreviewUrl] - Vorschau-URL für Upload
 * @property {string[]} bucketOptions - Verfügbare Bucket-Namen
 * @property {string} newBucketInputEdit - Neuer Bucket-Name bei Bearbeitung
 * @property {string} newBucketInputUpload - Neuer Bucket-Name bei Upload
 * @property {"buckets"|"items"} viewMode - Aktuelle Ansicht (Ordner oder Dateien)
 * @property {"folders"|"files"} resultMode - Ergebnis-Anzeigemodus
 * @property {string} [selectedBucket] - Aktuell ausgewählter Bucket
 * @property {boolean} isUploading - Upload läuft?
 * @property {number} [downloadingItemId] - ID der gerade heruntergeladenen Datei
 * @property {boolean} showScrollTop - "Nach oben"-Button sichtbar?
 * @property {number} bucketsToShow - Anzahl anzuzeigender Buckets (Pagination)
 * @property {number} visibleItemsCount - Anzahl anzuzeigender Elemente (Pagination)
 * @property {number} uploadProgress - Upload-Fortschritt (0-100%)
 * @property {number} uploadCurrentFile - Index der aktuellen Upload-Datei
 * @property {number} uploadTotalFiles - Gesamtzahl der Upload-Dateien
 * @property {string[]} categoryOptions - Verfügbare Kategorieoptionen
 * @property {string} [filterDienst] - Ausgewählter Dienst-Filter
 * @property {string} [filterCreator] - Ausgewählter Ersteller-Filter
 * @property {string[]} dienstOptions - Verfügbare Dienst-Optionen
 */
interface IMediaAssetsLibState {
  allItems: IMediaItem[];
  visibleItems: IMediaItem[];
  searchText: string;
  filterCategory?: string;
  filterYear?: number;
  filterMonth?: number;
  filterFormat?: string;

  isModalOpen: boolean;
  selectedItem?: IMediaItem;

  isEditOpen: boolean;
  editName: string;
  editTags: string[];
  editCategory: string;
  editFormat: string;
  editDienst: string;

  editBucket: string[];
  uploadBucket: string[];

  isUploadOpen: boolean; // ✅ NEU
  uploadName: string;
  uploadTags: string[];
  uploadCategory: string;
  uploadDienst: string;
  uploadFiles?: File[];
  uploadPreviewUrl?: string;

  bucketOptions: string[];

  newBucketInputEdit: string;
  newBucketInputUpload: string;

  viewMode: "buckets" | "items";
  resultMode: "folders" | "files";
  selectedBucket?: string;

  isUploading: boolean;
  downloadingItemId?: number;

  showScrollTop: boolean;

  bucketsToShow: number;

  visibleItemsCount: number;

  uploadProgress: number;
  uploadCurrentFile: number;
  uploadTotalFiles: number;

  categoryOptions: string[];
  filterDienst?: string;

  filterCreator?: string;
  dienstOptions: string[];
}

export { IMediaItem, ISPFile, IMediaAssetsLibState };
