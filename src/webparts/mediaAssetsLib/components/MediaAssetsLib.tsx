import * as React from "react";
import { SPHttpClient } from "@microsoft/sp-http";
import type { IMediaAssetsLibProps } from "./IMediaAssetsLibProps";
import styles from "./MediaAssetsLib.module.scss";

import BucketDropdown from "./BucketDropdown";

import UploadModal from "./UploadModal";

/*******************************************************
 * MEDIA ASSETS LIB V8
 * -----------------------------------------------------
 * SharePoint Medienverwaltung
 * - Upload
 * - Edit
 * - Delete
 * - Filter
 * - Bucket Dropdown
 *******************************************************/

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
interface IMediaItem {
  id: number;
  name: string;
  fileRef: string;
  category?: string;
  notes?: string;
  created?: string;
  uniqueId?: string;
  tags?: string[];
  bucket?: string[];

  driveId?: string;
  driveItemId?: string;

  format?: string;
  createdBy?: string;
  dienst?: string;
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
export interface IMediaAssetsLibState {
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

/********************* HAUPTKOMPNENTE ******************
 * Steuert:
 * - Datenladen
 * - Upload
 * - Update
 * - Rendering
 ******************************************************/

/**
 * MediaAssetsLib - SharePoint Media Assets Verwaltungskomponente
 *
 * Vollständige Medienbibliothek-Komponente für SharePoint mit folgenden Funktionen:
 * - Rekursives Laden aller Dateien und Ordner aus einer SharePoint-Dokumentbibliothek
 * - Organisierung von Dateien in Buckets (Ordner-basierte Kategorisierung)
 * - Upload von Dateien mit automatischer Metadaten-Erkennung
 * - Bearbeitung von Metadaten (Name, Tags, Kategorie, Format, Bucket)
 * - Löschen von Elementen
 * - Erweiterte Filterung (Text, Kategorie, Dienst, Format, Datum, Ersteller)
 * - Suchfunktion mit Tag-Unterstützung
 * - Vorschau von Bildern, Videos und Audio-Dateien
 * - Download-Funktionalität
 * - Pagination für Buckets und Elemente
 * - Intersection Observer für "Nach oben"-Button
 *
 * @class MediaAssetsLib
 * @extends {React.Component<IMediaAssetsLibProps, IMediaAssetsLibState>}
 * @example
 * <MediaAssetsLib
 *   siteUrl="https://tenant.sharepoint.com/sites/mysite"
 *   spHttpClient={spHttpClient}
 *   isDarkTheme={false}
 * />
 */
export default class MediaAssetsLib extends React.Component<
  IMediaAssetsLibProps,
  IMediaAssetsLibState
> {
  /********************* INITIAL STATE *******************/

  /**
   * Konstruktor der Komponente
   * Initialisiert alle State-Properties mit Standardwerten
   * @constructor
   * @param {IMediaAssetsLibProps} props - Komponenten-Props mit siteUrl und spHttpClient
   */
  constructor(props: IMediaAssetsLibProps) {
    super(props);

    this.state = {
      allItems: [],
      visibleItems: [],
      searchText: "",
      filterYear: undefined,
      filterMonth: undefined,

      isModalOpen: false,
      selectedItem: undefined,

      isEditOpen: false,
      editName: "",
      editTags: [],
      editCategory: "",
      editFormat: "",
      editBucket: [],

      isUploadOpen: false,
      uploadName: "",
      uploadTags: [],
      uploadCategory: "",
      uploadDienst: "",
      uploadBucket: [],
      uploadFiles: [],

      bucketOptions: [],

      newBucketInputEdit: "",
      newBucketInputUpload: "",

      viewMode: "buckets",
      resultMode: "folders",
      selectedBucket: undefined,

      isUploading: false,
      downloadingItemId: undefined,
      showScrollTop: false,
      bucketsToShow: 5,
      visibleItemsCount: 20,

      uploadProgress: 0,
      uploadCurrentFile: 0,
      uploadTotalFiles: 0,

      categoryOptions: [],

      filterDienst: undefined,
      dienstOptions: [],
    };
  }

  /* CHANGE SHAREPOINT URL */
  /**
   * Name der SharePoint-Dokumentbibliothek
   * Wird für alle REST-API-Aufrufe verwendet
   * @private
   * @type {string}
   */
  private readonly libraryName = "Medienbibliothek";

  /**
   * Lädt die verfügbaren Kategorieoptionen aus dem Kategorie-Feld der SharePoint-Liste
   * Speichert die Optionen im State für die Kategoriefilter-Dropdown
   * @private
   * @async
   * @returns {Promise<void>}
   * @throws {Error} Wenn der REST-API-Aufruf fehlschlägt
   */
  private async loadCategories(): Promise<void> {
    try {
      const response = await this.props.spHttpClient.get(
        `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.libraryName}')/fields/getbyinternalnameortitle('Kategorie')`,
        SPHttpClient.configurations.v1,
      );

      const data = await response.json();

      this.setState({
        categoryOptions: data.Choices || [],
      });

      console.log("SHAREPOINT CATEGORIES", data.Choices);
    } catch (error) {
      console.error("Fehler beim Laden der Kategorien", error);
    }
  }

  /**
   * Gibt den Bibliotheks-Pfad für REST-API-Aufrufe zurück
   * Kombiniert die Site-URL mit dem Bibliotheksnamen
   * @private
   * @returns {string} Relativer Pfad zur Bibliothek
   * @example
   * // Returns: "/sites/mysite/Medienbibliothek"
   * getLibraryPath();
   */
  private getLibraryPath(): string {
    return `${new URL(this.props.siteUrl).pathname}/${this.libraryName}`;
  }

  /**
   * Lädt die verfügbaren Dienstoptionen aus dem Dienste-Feld der SharePoint-Liste
   * Speichert die Optionen im State für die Dienste-Filter-Dropdown
   * @private
   * @async
   * @returns {Promise<void>}
   * @throws {Error} Wenn der REST-API-Aufruf fehlschlägt
   */
  private async loadDienste(): Promise<void> {
    try {
      const response = await this.props.spHttpClient.get(
        `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.libraryName}')/fields/getbyinternalnameortitle('Dienste')`,
        SPHttpClient.configurations.v1,
      );

      const data = await response.json();

      this.setState({
        dienstOptions: data.Choices || [],
      });
    } catch (error) {
      console.error("Fehler beim Laden der Dienste", error);
    }
  }

  /**
   * IntersectionObserver für das "Nach oben"-Button-Verhalten
   * Beobachtet die Sichtbarkeit des Header-Elements
   * @private
   * @type {IntersectionObserver|undefined}
   */
  private observer?: IntersectionObserver;

  /**
   * Erkennt den Dateityp basierend auf der Dateiendung
   * Kategorisiert Dateien in: Bild, Video, Audio oder Dokument
   * @private
   * @param {string} fileName - Der Dateiname mit Erweiterung
   * @returns {string} Der erkannte Dateityp (Bild, Video, Audio oder Dokument)
   * @example
   * detectFormat('photo.jpg')  // Returns: "Bild"
   * detectFormat('video.mp4')  // Returns: "Video"
   * detectFormat('song.mp3')   // Returns: "Audio"
   * detectFormat('doc.pdf')    // Returns: "Dokument"
   */
  private detectFormat(fileName: string): string {
    const ext = fileName.split(".").pop()?.toLowerCase();

    if (!ext) {
      return "Dokument";
    }

    const imageTypes = ["jpg", "jpeg", "png", "gif", "webp", "bmp", "svg"];

    const videoTypes = ["mp4", "mov", "avi", "webm", "mkv", "wmv", "m4v"];

    const audioTypes = [
      "mp3",
      "wav",
      "aiff",
      "aac",
      "flac",
      "ogg",
      "wma",
      "m4a",
    ];

    if (imageTypes.includes(ext)) {
      return "Bild";
    }

    if (videoTypes.includes(ext)) {
      return "Video";
    }

    if (audioTypes.includes(ext)) {
      return "Audio";
    }

    return "Dokument";
  }
  /**
   * Sammelt alle einzigartigen Bucket-Namen aus allen Elementen
   * @private
   * @returns {string[]} Array aller Bucket-Namen
   */
  private getAllBuckets(): string[] {
    const map: { [key: string]: number } = {};

    this.state.allItems.forEach((item) => {
      if (item.bucket) {
        item.bucket.forEach((b) => {
          map[b] = (map[b] || 0) + 1;
        });
      }
    });

    return Object.keys(map);
  }

  /**
   * Findet das erste Element eines bestimmten Buckets für die Vorschau
   * @private
   * @param {string} bucket - Der Bucket-Name
   * @returns {IMediaItem|undefined} Das erste Element des Buckets oder undefined
   */
  private getBucketPreview(bucket: string): IMediaItem | undefined {
    return this.state.allItems.find((item) => item.bucket?.includes(bucket));
  }

  /**
   * Filtert Buckets basierend auf aktiven Filtern und Suchtext
   * Ein Bucket wird angezeigt, wenn er mindestens ein gefordertes Element enthält
   * @private
   * @returns {string[]} Array der gefilterten Bucket-Namen
   */
  private getFilteredBuckets(): string[] {
    const {
      searchText,
      filterCategory,
      filterDienst,
      filterFormat,
      filterYear,
      filterMonth,
      filterCreator,
    } = this.state;

    const buckets = this.getAllBuckets();

    return buckets.filter((bucket) => {
      const items = this.state.allItems.filter((item) =>
        item.bucket?.includes(bucket),
      );

      const search = searchText.toLowerCase();

      // ✅ gleiche Filterlogik wie applyFilters
      const filteredItems = items.filter((item) => {
        // TEXT
        const matchesText =
          !search ||
          item.name.toLowerCase().includes(search) ||
          (item.tags || []).join(" ").toLowerCase().includes(search);

        // KATEGORIE
        const matchesCategory =
          !filterCategory || item.category === filterCategory;

        const matchesDienst = !filterDienst || item.dienst === filterDienst;

        // FORMAT
        const matchesFormat = !filterFormat || item.format === filterFormat;

        // DATUM
        const date = item.created ? new Date(item.created) : null;

        const matchesYear =
          !filterYear || (date && date.getFullYear() === filterYear);

        const matchesMonth =
          !filterMonth || (date && date.getMonth() + 1 === filterMonth);

        const matchesCreator =
          !filterCreator || item.createdBy === filterCreator;

        return (
          matchesText &&
          matchesCategory &&
          matchesDienst &&
          matchesCreator &&
          matchesFormat &&
          matchesYear &&
          matchesMonth
        );
      });

      const bucketMatchesSearch = search
        ? bucket.toLowerCase().includes(search)
        : false;

      // ✅ Wenn Suche aktiv → Name ODER Inhalt
      if (search) {
        return bucketMatchesSearch || filteredItems.length > 0;
      }

      // ✅ Wenn KEINE Suche → nur Inhalt (Filter!)
      return filteredItems.length > 0;
    });
  }

  /**
   * Zählt die Anzahl der Elemente in jedem Bucket
   * @private
   * @returns {{ [key: string]: number }} Objekt mit Bucket-Namen als Schlüssel und Elementanzahl als Wert
   */
  private getBucketCounts(): { [key: string]: number } {
    const counts: { [key: string]: number } = {};

    this.state.allItems.forEach((item) => {
      if (item.bucket) {
        item.bucket.forEach((b) => {
          counts[b] = (counts[b] || 0) + 1;
        });
      }
    });

    return counts;
  }

  /**
   * Sortiert Buckets nach dem neuesten Element in jedem Bucket (absteigend)
   * @private
   * @param {IMediaItem[]} items - Array von Medienelementen
   * @returns {string[]} Sortierte Bucket-Namen (neueste zuerst)
   */
  private getBucketsSortedByNewest(items: IMediaItem[]): string[] {
    const map: { [key: string]: number } = {};

    items.forEach((item) => {
      if (item.bucket && item.created) {
        const time = new Date(item.created).getTime();

        item.bucket.forEach((b) => {
          if (!map[b] || map[b] < time) {
            map[b] = time;
          }
        });
      }
    });

    return Object.keys(map).sort((a, b) => {
      return map[b] - map[a]; // neueste zuerst
    });
  }

  /**
   * Holt das FormDigestValue von SharePoint für POST/MERGE-Operationen
   * Erforderlich für Schreibzugriff auf SharePoint REST API
   * @private
   * @async
   * @returns {Promise<string>} Das FormDigestValue Token
   * @throws {Error} Wenn der Digest-Abruf fehlschlägt
   */
  private async getRequestDigest(): Promise<string> {
    const response = await this.props.spHttpClient.post(
      `${this.props.siteUrl}/_api/contextinfo`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=nometadata",
        },
      },
    );

    const data = await response.json();

    return data.FormDigestValue;
  }

  /**
   * Ruft Informationen des aktuellen angemeldeten Benutzers ab
   * Wird verwendet, um den Ersteller bei Upload zu speichern
   * @private
   * @async
   * @returns {Promise<{id: number, title: string}>} Benutzer-ID und Name
   * @throws {Error} Wenn der Benutzerabruf fehlschlägt
   */
  private async getCurrentUser(): Promise<{ id: number; title: string }> {
    const response = await this.props.spHttpClient.get(
      `${this.props.siteUrl}/_api/web/currentuser`,
      SPHttpClient.configurations.v1,
    );

    const user = await response.json();

    return {
      id: user.Id,
      title: user.Title,
    };
  }

  /********************* UPLOAD **************************
   * Lädt Dateien hoch und setzt Metadaten
   ******************************************************/

  /**
   * Lädt eine Datei zu SharePoint hoch mit Upload-Fortschrittsanzeige
   * Verwendet XMLHttpRequest für genaue Fortschrittsüberwachung
   * @private
   * @async
   * @param {string} url - Die SharePoint REST-API URL für den Upload
   * @param {ArrayBuffer} fileBuffer - Der Datei-Buffer
   * @returns {Promise<{ServerRelativeUrl: string}>} Die relative URL der hochgeladenen Datei
   * @throws {Error} Bei Upload-Fehler
   */
  private async uploadFileWithProgress(
    url: string,
    fileBuffer: ArrayBuffer,
  ): Promise<{ ServerRelativeUrl: string }> {
    const digest = await this.getRequestDigest();

    return new Promise((resolve, reject) => {
      const xhr = new XMLHttpRequest();

      // IMMER ZUERST OPEN
      xhr.open("POST", url, true);

      // DANN HEADER
      xhr.setRequestHeader("Accept", "application/json;odata=nometadata");

      xhr.setRequestHeader("X-RequestDigest", digest);

      xhr.upload.onprogress = (event) => {
        if (event.lengthComputable) {
          const percent = Math.round((event.loaded / event.total) * 100);

          this.setState({
            uploadProgress: percent,
          });
        }
      };

      xhr.onload = () => {
        if (xhr.status >= 200 && xhr.status < 300) {
          resolve(JSON.parse(xhr.responseText));
        } else {
          reject(xhr.responseText);
        }
      };

      xhr.onerror = () => reject(xhr.responseText);

      xhr.send(fileBuffer);
    });
  }

  /**
   * Orchestriert den Upload aller ausgewählten Dateien
   * - Validiert Input (Bucket, Kategorie, Dateien)
   * - Lädt jede Datei einzeln hoch
   * - Setzt Metadaten für jede Datei
   * - Zeigt Upload-Fortschritt an
   * - Lädt alle Daten neu nach Upload
   * @private
   * @async
   * @returns {Promise<void>}
   * @throws {Error} Wenn Upload fehlschlägt
   */
  private async uploadItem(): Promise<void> {
    const {
      uploadFiles,
      uploadTags,
      uploadCategory,
      uploadDienst,
      uploadBucket,
    } = this.state;

    this.setState({
      isUploading: true,
      uploadProgress: 0,
      uploadCurrentFile: 0,
      uploadTotalFiles: uploadFiles?.length || 0,
    });

    if (!uploadBucket || uploadBucket.length === 0) {
      alert("Bitte wählen Sie mindestens einen Ordner aus.");
      this.setState({ isUploading: false }); // ✅ FIX!!
      return;
    }

    if (!uploadCategory) {
      alert("Bitte wählen Sie eine Kategorie aus.");
      this.setState({ isUploading: false });
      return;
    }

    if (!uploadFiles || uploadFiles.length === 0) {
      console.log("❌ Keine Files gewählt");
      this.setState({ isUploading: false }); // ✅ FIX!!
      return;
    }

    for (let i = 0; i < uploadFiles.length; i++) {
      const uploadFile = uploadFiles[i];

      this.setState({
        uploadCurrentFile: i + 1,
      });

      try {
        const fileBuffer = await uploadFile.arrayBuffer();

        const uploadUrl =
          `${this.props.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${this.getLibraryPath()}')` +
          `/Files/add(overwrite=true,url='${uploadFile.name}')`;

        console.log("⬆️ Starte Upload:", uploadFile.name);

        const result = await this.uploadFileWithProgress(uploadUrl, fileBuffer);

        const fileUrl = result.ServerRelativeUrl;

        const itemResponse = await this.props.spHttpClient.get(
          `${this.props.siteUrl}/_api/web/GetFileByServerRelativeUrl('${fileUrl}')/ListItemAllFields`,
          SPHttpClient.configurations.v1,
        );

        const itemData = await itemResponse.json();
        const itemId = itemData.Id;

        const updateUrl = `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.libraryName}')/items(${itemId})`;

        const cleanTags = uploadTags
          .map((t) => t.toLowerCase().trim())
          .filter((t, i, arr) => t && arr.indexOf(t) === i);

        const cleanBuckets = uploadBucket
          .map((t) => t.trim())
          .filter((t, i, arr) => t && arr.indexOf(t) === i);

        const isSingleUpload = uploadFiles.length === 1;

        let finalName = uploadFile.name;

        if (isSingleUpload && this.state.uploadName) {
          const extension = uploadFile.name.includes(".")
            ? uploadFile.name.substring(uploadFile.name.lastIndexOf("."))
            : "";

          finalName = this.state.uploadName + extension;
        }

        const detectedFormat = this.detectFormat(uploadFile.name);
        const currentUser = await this.getCurrentUser();
        console.log("CURRENT USER", currentUser);
        const bodyData = {
          FileLeafRef: finalName,
          Kategorie: uploadCategory || null,
          Dienste: uploadDienst || null,
          Tags: cleanTags,
          Format: detectedFormat, // ✅ AUTO!
          Bucket: cleanBuckets,
          Ersteller: currentUser.title,
        };

        console.log("UPLOAD BUCKET RAW:", uploadBucket);
        console.log("CLEAN BUCKET:", cleanBuckets);

        await this.props.spHttpClient.post(
          updateUrl,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: "application/json;odata=nometadata",
              "Content-Type": "application/json;odata=nometadata",
              "IF-MATCH": "*",
              "X-HTTP-Method": "MERGE",
            },
            body: JSON.stringify(bodyData),
          },
        );
        console.log("FINAL BODY:", bodyData);
        console.log("✅ Fertig:", uploadFile.name);
        this.setState({
          uploadProgress: 100,
        });
      } catch (error) {
        console.error("❌ Fehler bei Datei:", uploadFile.name, error);
      }
    }

    await this.loadAllMedia();

    this.setState({
      isUploading: false,
      uploadProgress: 0,
      uploadCurrentFile: 0,
      uploadTotalFiles: 0,
    });

    this.setState(
      {
        isUploadOpen: false,
        uploadFiles: [],
        viewMode: "buckets", // ✅ zurück zur Bucket Ansicht
        selectedBucket: undefined, // ✅ Reset
        searchText: "", // ✅ wichtig für Ansicht
      },
      this.applyFilters, // ✅ neu berechnen
    );
  }

  /********************* DELETE **************************
   * Löscht ein Element aus SharePoint
   ******************************************************/

  /**
   * Löscht das ausgewählte Element aus SharePoint
   * - Fordert Bestätigung vom Benutzer an
   * - Löscht das Element über REST API
   * - Aktualisiert State (Rückkehr zur Bucket-Ansicht wenn Bucket leer)
   * @private
   * @async
   * @returns {Promise<void>}
   * @throws {Error} Wenn Delete fehlschlägt
   */
  private async deleteItem(): Promise<void> {
    const { selectedItem, selectedBucket } = this.state;

    if (!selectedItem) return;

    const confirmDelete = window.confirm(
      "Möchten Sie dieses Element wirklich löschen?",
    );

    if (!confirmDelete) return;

    try {
      const url = `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.libraryName}')/items(${selectedItem.id})`;
      const response = await this.props.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
            "IF-MATCH": "*",
            "X-HTTP-Method": "DELETE",
          },
        },
      );

      if (response.ok) {
        console.log("✅ Element gelöscht");

        const rootFolder = this.getLibraryPath();
        const items = await this.getFolderContent(rootFolder);

        // ✅ Buckets neu berechnen
        const sortedBuckets = this.getBucketsSortedByNewest(items);

        // ✅ Prüfen ob noch Dateien im aktuellen Ordner sind
        const remainingInBucket = items.filter((i) =>
          selectedBucket ? i.bucket?.includes(selectedBucket) : false,
        );

        this.setState(
          {
            allItems: items,
            bucketOptions: sortedBuckets,

            isEditOpen: false,
            selectedItem: undefined,

            // ✅ FALL 1: Ordner jetzt leer → zurück zur Übersicht
            ...(selectedBucket && remainingInBucket.length === 0
              ? {
                  viewMode: "buckets",
                  resultMode: "folders",
                  selectedBucket: undefined,
                  searchText: "",
                  filterCategory: undefined,
                  filterDienst: undefined,
                  filterFormat: undefined,
                  filterYear: undefined,
                  filterMonth: undefined,
                  filterCreator: undefined,
                }
              : // ✅ FALL 2: Ordner hat noch Dateien → drin bleiben
                {
                  viewMode: "items",
                  resultMode: "files",
                  selectedBucket: selectedBucket,
                  searchText: "",
                  filterCategory: undefined,
                  filterDienst: undefined,
                  filterFormat: undefined,
                  filterYear: undefined,
                  filterMonth: undefined,
                  filterCreator: undefined,
                }),
          },
          this.applyFilters,
        );
      } else {
        console.error("❌ Fehler beim Löschen", response);
      }
    } catch (error) {
      console.error("❌ Delete Fehler:", error);
    }
  }

  /********************* INIT ****************************
   * Lädt Daten beim Start
   ******************************************************/

  /**
   * React-Lifecycle-Methode: Wird aufgerufen nachdem die Komponente bereitgestellt wurde
   * - Lädt alle Mediendateien rekursiv
   * - Lädt verfügbare Kategorien und Dienste
   * - Setzt up IntersectionObserver für "Nach oben"-Button
   * @public
   * @async
   * @returns {Promise<void>}
   */
  public async componentDidMount(): Promise<void> {
    console.log("MOUNTED ✅");

    console.log("Aktuelle Site:");
    console.log(this.props.siteUrl);

    await this.loadAllMedia();
    await this.loadCategories();
    await this.loadDienste();

    const target = document.getElementById("top");

    if (target) {
      this.observer = new IntersectionObserver(
        ([entry]) => {
          console.log("HEADER VISIBLE:", entry.isIntersecting);

          this.setState({
            showScrollTop: !entry.isIntersecting,
          });
        },
        {
          root: null, // viewport
          threshold: 0,
        },
      );

      this.observer.observe(target);
    }
  }

  /********************* REKURSIVER LOAD *****************
   * Lädt alle Dateien + Unterordner
   ******************************************************/

  /**
   * Lädt rekursiv alle Dateien und Unterordner eines SharePoint-Ordners
   * - Verwendet SharePoint REST API mit $expand für Dateien und Ordner
   * - Konvertiert SharePoint-Felder in IMediaItem-Objekte
   * - Behandelt mehrzeilige Text- und Choice-Felder (Tags, Buckets)
   * - Ignoriert den "Forms"-Ordner (System-Ordner)
   * @private
   * @async
   * @param {string} folderUrl - Relative URL des Ordners
   * @returns {Promise<IMediaItem[]>} Array aller Medienelemente im Ordner und Unterordnern
   * @throws {Error} Wenn REST-API-Aufruf fehlschlägt
   */
  private async getFolderContent(folderUrl: string): Promise<IMediaItem[]> {
    const url = `${this.props.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')
?$expand=Folders,Files,Files/ListItemAllFields
&$select=
Name,
ServerRelativeUrl,
TimeCreated,
Files/Name,
Files/ServerRelativeUrl,
Files/TimeCreated,
Files/ListItemAllFields/Id,
Files/ListItemAllFields/Kategorie,
Files/ListItemAllFields/Dienste,
Files/ListItemAllFields/Notizen,
Files/ListItemAllFields/UniqueId,
Files/ListItemAllFields/Tags,
Files/ListItemAllFields/Bucket,
Files/ListItemAllFields/Format,
Files/ListItemAllFields/Ersteller`;
    const response = await this.props.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
    );

    const data = await response.json();
    console.log("DATA", data);

    let results: IMediaItem[] = [];

    data.Files.forEach((file: unknown): void => {
      const f = file as ISPFile;

      results.push({
        id: f.ListItemAllFields.Id,
        name: f.Name,
        fileRef: f.ServerRelativeUrl,
        category: f.ListItemAllFields?.Kategorie,
        dienst: f.ListItemAllFields?.Dienste,
        notes: f.ListItemAllFields?.Notizen,
        created: f.TimeCreated,
        tags: Array.isArray(f.ListItemAllFields?.Tags)
          ? f.ListItemAllFields.Tags
          : f.ListItemAllFields?.Tags
            ? String(f.ListItemAllFields.Tags)
                .split(/[;,#]+/)
                .map((t: string) => t.trim())
                .filter((t: string) => t)
            : [],
        bucket: Array.isArray(f.ListItemAllFields?.Bucket)
          ? f.ListItemAllFields.Bucket
          : f.ListItemAllFields?.Bucket
            ? String(f.ListItemAllFields.Bucket)
                .split(/[;,#]+/)
                .map((t: string) => t.trim())
                .filter((t: string) => t)
            : [],

        driveId: f.ListItemAllFields?.File?.VroomDriveID,
        driveItemId: f.ListItemAllFields?.File?.VroomItemID,
        format: f.ListItemAllFields?.Format,
        createdBy: f.ListItemAllFields?.Ersteller,
      });
      console.log(f.ListItemAllFields);
    });

    for (const folder of data.Folders) {
      if (folder.Name !== "Forms") {
        const subItems = await this.getFolderContent(folder.ServerRelativeUrl);
        results = results.concat(subItems);
      }
    }

    return results;
  }

  /********************* LOAD MEDIA **********************/

  /**
   * Hauptmethode zum Laden aller Mediendateien aus der Bibliothek
   * - Ruft die Root-Bibliothek auf
   * - Sortiert Buckets nach neuesten Elementen
   * - Aktualisiert State und wendet Filter an
   * @private
   * @async
   * @returns {Promise<void>}
   * @throws {Error} Wenn Laden fehlschlägt
   */
  private async loadAllMedia(): Promise<void> {
    try {
      const rootFolder = this.getLibraryPath();
      const items = await this.getFolderContent(rootFolder);

      const uniqueBuckets = this.getBucketsSortedByNewest(items);

      this.setState(
        {
          allItems: items,
          bucketOptions: uniqueBuckets,
        },
        this.applyFilters,
      );
    } catch (error) {
      console.error(error);
    }
  }

  /********************* UPDATE **************************
   * Aktualisiert Metadaten eines Elements
   ******************************************************/

  /**
   * Aktualisiert die Metadaten des ausgewählten Elements
   * - Ändert Name, Tags, Kategorie, Format und Buckets
   * - Sendet MERGE-Request an SharePoint REST API
   * - Lädt alle Daten neu nach Update
   * @private
   * @async
   * @returns {Promise<void>}
   * @throws {Error} Wenn Update fehlschlägt
   */
  private async updateItem(): Promise<void> {
    const {
      selectedItem,
      editName,
      editTags,
      editCategory,
      editFormat,
      editBucket,
    } = this.state;

    if (!selectedItem) return;

    try {
      /* BIBLIOTHEK ÄNDERN */
      const url = `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.libraryName}')/items(${selectedItem.id})`;
      const tagsArray = editTags;

      const bucketsArray = editBucket;

      // DEBBUG ___________________________________________________
      console.log(
        "🧪 UPDATE BODY:",
        JSON.stringify(
          {
            FileLeafRef: editName,
            Kategorie: editCategory,
            body: JSON.stringify({
              FileLeafRef: editName,

              Kategorie: editCategory,

              Tags: tagsArray,

              Format: editFormat,

              Bucket: bucketsArray,
            }),
          },
          null,
          2,
        ),
      );

      console.log("🧪 EDIT BUCKET:", editBucket);

      const response = await this.props.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE",
          },
          /* _______________________________________________________________________ */
          body: JSON.stringify({
            FileLeafRef: editName,
            Kategorie: editCategory,
            Tags: tagsArray,
            Format: editFormat,
            Bucket: bucketsArray,
          }),
        },
      );

      const errorText = await response.text();
      console.log("🧪 UPDATE RESPONSE:", errorText);

      console.log("Update erfolgreich");

      // neu laden → damit UI sofort aktualisiert wird
      await this.loadAllMedia();

      this.setState({ isEditOpen: false });
    } catch (error) {
      console.error("Update Fehler:", error);
    }
  }
  /********************* FILTER LOGIK ********************
   * Suche + Kategorie + Datum + Format
   ******************************************************/

  /**
   * Wendet alle aktiven Filter auf die Elemente an und aktualisiert visibleItems
   * Filtert nach:
   * - Bucket (wenn ausgewählt)
   * - Suchtext (Name, Notes, Tags mit AND-Logik)
   * - Kategorie, Dienst, Format, Ersteller
   * - Jahr und Monat
   * Sortiert Ergebnisse nach Erstellungsdatum (neueste zuerst)
   * @private
   * @returns {void}
   */
  private applyFilters = (): void => {
    const {
      allItems,
      searchText,
      filterCategory,
      filterDienst,
      filterFormat,
      filterCreator,
      selectedBucket,
    } = this.state;

    let filtered = [...allItems];

    // ✅ IMMER ZUERST BUCKET einschränken
    if (selectedBucket) {
      filtered = filtered.filter((item) =>
        item.bucket?.includes(selectedBucket),
      );
    }
    console.log("SELECTED BUCKET:", selectedBucket);
    console.log("ITEMS NACH BUCKET FILTER:", filtered.length);

    const search = searchText.toLowerCase();

    if (search) {
      const terms = search.split(" ").filter((t) => t);

      filtered = filtered.filter((item) => {
        const name = item.name?.toLowerCase() || "";
        const notes = item.notes?.toLowerCase() || "";
        const tags = (item.tags || []).join(" ").toLowerCase();

        return terms.every(
          (term) =>
            name.includes(term) || notes.includes(term) || tags.includes(term),
        );
      });
    }

    if (filterCategory) {
      filtered = filtered.filter((item) => item.category === filterCategory);
    }

    if (filterDienst) {
      filtered = filtered.filter((item) => item.dienst === filterDienst);
    }

    if (filterFormat) {
      filtered = filtered.filter((item) => item.format === filterFormat);
    }

    if (filterCreator) {
      filtered = filtered.filter((item) => item.createdBy === filterCreator);
    }

    if (this.state.filterYear) {
      filtered = filtered.filter((item) => {
        if (!item.created) return false;

        const date = new Date(item.created);
        return date.getFullYear() === this.state.filterYear;
      });
    }

    /*
    if (this.state.filterYear) {
      filtered = filtered.filter((item) => {
        console.log("CREATED RAW:", item.created);

        if (!item.created) return false;

        const date = new Date(item.created);

        console.log("PARSED DATE:", date);
        console.log("YEAR:", date.getFullYear());

        const itemYear = date.getFullYear();
        const selectedYear = this.state.filterYear;

        console.log("COMPARE:", itemYear, selectedYear);

        return itemYear == selectedYear;
      });
    }
      */

    if (this.state.filterMonth) {
      filtered = filtered.filter((item) => {
        if (!item.created) return false;

        const date = new Date(item.created);

        // Achtung: JS Monate = 0–11
        const month = date.getMonth() + 1;

        return month === this.state.filterMonth;
      });
    }

    // SORTIERUNG (neueste zuerst)
    filtered.sort((a, b) => {
      const dateA = a.created ? new Date(a.created).getTime() : 0;
      const dateB = b.created ? new Date(b.created).getTime() : 0;

      return dateB - dateA; // neueste zuerst
    });

    this.setState({ visibleItems: filtered, visibleItemsCount: 20 });
  };

  /* +++++++++  SUCHE ++++++++ */

  /**
   * Event-Handler für Suchtext-Eingabe
   * - Aktualisiert Suchtext im State
   * - Wechselt zur Artikel-Ansicht (viewMode = "items")
   * - Wendet Filter an
   * @private
   * @param {React.ChangeEvent<HTMLInputElement>} e - Input-Change-Event
   * @returns {void}
   */
  private onSearchChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const value = e.target.value;

    this.setState(
      {
        searchText: value,
        viewMode: "items",
        // ✅ WICHTIG: selectedBucket NICHT löschen!
      },
      this.applyFilters,
    );
  };

  /**
   * Extrahiert alle einzigartigen Kategorien aus den geladenen Elementen
   * @private
   * @returns {string[]} Array eindeutiger Kategoriewerte
   */
  private getUniqueCategories(): string[] {
    const values = this.state.allItems
      .map((item) => item.category)
      .filter((v) => v);
    console.log("ALL ITEMS", this.state.allItems);
    console.log("CATEGORIES", values);
    return Array.from(new Set(values)) as string[];
  }

  /**
   * Extrahiert alle einzigartigen Dienste aus den geladenen Elementen
   * @private
   * @returns {string[]} Array eindeutiger Dienst-Werte
   */
  private getUniqueDienste(): string[] {
    const values = this.state.allItems
      .map((item) => item.dienst)
      .filter((v) => v);

    return Array.from(new Set(values)) as string[];
  }

  /**
   * Extrahiert alle einzigartigen Dateiformate aus den geladenen Elementen
   * @private
   * @returns {string[]} Array eindeutiger Format-Werte
   */
  private getUniqueFormats(): string[] {
    const values = this.state.allItems
      .map((item) => item.format)
      .filter((v): v is string => !!v);

    return Array.from(new Set(values));
  }

  /**
   * Extrahiert alle einzigartigen Jahre aus den geladenen Elementen
   * Sortiert sie in absteigender Reihenfolge (neueste zuerst)
   * @private
   * @returns {number[]} Array einzigartiger Jahre
   */
  private getUniqueYears(): number[] {
    const years = this.state.allItems
      .map((item) => {
        if (!item.created) return null;

        const date = new Date(item.created);
        return date.getFullYear();
      })
      .filter((y) => y !== null) as number[];

    // Duplikate entfernen + sortieren (neueste zuerst)
    return Array.from(new Set(years)).sort((a, b) => b - a);
  }

  /**
   * Extrahiert alle einzigartigen Ersteller aus den geladenen Elementen
   * Sortiert sie alphabetisch
   * @private
   * @returns {string[]} Array eindeutiger Ersteller-Namen
   */
  private getUniqueCreators(): string[] {
    const values = this.state.allItems
      .map((item) => item.createdBy)
      .filter((v): v is string => !!v);

    return Array.from(new Set(values)).sort();
  }
  /********************* BUCKET LADEN ********************
   * Lädt Choice Werte aus SharePoint
   ******************************************************/

  /**
   * React-Render-Methode: Erzeugt die komplette Benutzeroberfläche
   * - Header mit Upload-Button und View-Toggles
   * - Suchleiste und Filter-Selects
   * - Grid-Ansicht für Buckets oder Elemente
   * - Modals für Preview, Bearbeitung und Upload
   * - "Nach oben"-Button bei Bedarf
   * @public
   * @returns {React.ReactElement<IMediaAssetsLibProps>} Die JSX-Komponente
   */
  public render(): React.ReactElement<IMediaAssetsLibProps> {
    const categoryOptions = this.getUniqueCategories();

    const dienstOptions = this.getUniqueDienste();

    const yearOptions = this.getUniqueYears();

    const formatOptions = this.getUniqueFormats();

    const creatorOptions = this.getUniqueCreators();

    const bucketCounts = this.getBucketCounts();

    /********************* RENDER **************************/

    return (
      <div className={styles.container}>
        {/* **************** HEADER **************** */}
        <h2 id="top">Media Library</h2>

        <div className={styles.header}>
          {!this.state.selectedBucket && (
            <div className={styles.toggleGroup}>
              <button
                onClick={() => this.setState({ resultMode: "folders" })}
                className={`${styles.toggleBtn} ${
                  this.state.resultMode === "folders"
                    ? styles.active
                    : styles.inactive
                }`}
              >
                Ordner
              </button>

              <button
                onClick={() => this.setState({ resultMode: "files" })}
                className={`${styles.toggleBtn} ${
                  this.state.resultMode === "files"
                    ? styles.active
                    : styles.inactive
                }`}
              >
                Dateien
              </button>
            </div>
          )}

          {/* ✅ HIER FEHLT ER – EINFÜGEN */}
          <button
            className={styles.uploadMainBtn}
            onClick={() => this.setState({ isUploadOpen: true })}
          >
            ＋
          </button>
        </div>

        {/* **************** SEARCH **************** */}
        <div>
          <input
            type="text"
            placeholder={
              this.state.selectedBucket
                ? `Suche in "${this.state.selectedBucket}"`
                : "Suche..."
            }
            value={this.state.searchText}
            onChange={this.onSearchChange}
            className={styles.searchInput}
          />
        </div>
        {this.state.viewMode === "items" && (
          <button
            onClick={() =>
              this.setState(
                {
                  viewMode: "buckets",
                  resultMode: "folders", // ✅ zurück zu Ordnern
                  selectedBucket: undefined,
                  searchText: "",
                  filterCategory: undefined,
                  filterDienst: undefined,
                  filterFormat: undefined,
                  filterYear: undefined,
                  filterMonth: undefined,
                  filterCreator: undefined,
                },
                this.applyFilters,
              )
            }
            className={styles.backBtn}
          >
            ← Zurück zu Ordnern
          </button>
        )}
        <div className={styles.filterRow}>
          <div className={styles.filterGroup}>
            {/* **************** FILTER **************** */}
            <select
              value={this.state.filterCategory || ""}
              onChange={(e) =>
                this.setState(
                  {
                    filterCategory: e.target.value || undefined,
                  },
                  this.applyFilters,
                )
              }
            >
              <option value="">Kategorie</option>
              {categoryOptions.map((cat) => (
                <option key={cat}>{cat}</option>
              ))}
            </select>

            <select
              value={this.state.filterDienst || ""}
              onChange={(e) =>
                this.setState(
                  {
                    filterDienst: e.target.value || undefined,
                  },
                  this.applyFilters,
                )
              }
            >
              <option value="">Dienst</option>

              {dienstOptions.map((dienst) => (
                <option key={dienst} value={dienst}>
                  {dienst}
                </option>
              ))}
            </select>
            <select
              value={this.state.filterCreator || ""}
              onChange={(e) =>
                this.setState(
                  {
                    filterCreator: e.target.value || undefined,
                  },
                  this.applyFilters,
                )
              }
            >
              <option value="">Ersteller</option>

              {creatorOptions.map((creator) => (
                <option key={creator} value={creator}>
                  {creator}
                </option>
              ))}
            </select>
            <select
              onChange={(e) =>
                this.setState(
                  { filterFormat: e.target.value || undefined },
                  this.applyFilters,
                )
              }
              value={this.state.filterFormat || ""}
            >
              <option value="">Format</option>

              {formatOptions.map((format) => (
                <option key={format} value={format}>
                  {format}
                </option>
              ))}
            </select>
          </div>
          <div className={styles.filterGroup}>
            {/* Jahr Filter */}
            <select
              value={this.state.filterMonth || ""}
              onChange={(e: React.ChangeEvent<HTMLSelectElement>) =>
                this.setState(
                  {
                    filterYear: e.target.value
                      ? Number(e.target.value)
                      : undefined,
                  },
                  this.applyFilters,
                )
              }
            >
              <option value="">Jahr</option>
              {yearOptions.map((year) => (
                <option key={year} value={year}>
                  {year}
                </option>
              ))}
            </select>

            {/* Monat Filter */}
            <select
              value={this.state.filterYear || ""}
              onChange={(e: React.ChangeEvent<HTMLSelectElement>) =>
                this.setState(
                  {
                    filterMonth: e.target.value
                      ? Number(e.target.value)
                      : undefined,
                  },
                  this.applyFilters,
                )
              }
            >
              <option value="">Monat</option>

              {[
                { v: 1, n: "Jan" },
                { v: 2, n: "Feb" },
                { v: 3, n: "Mär" },
                { v: 4, n: "Apr" },
                { v: 5, n: "Mai" },
                { v: 6, n: "Jun" },
                { v: 7, n: "Jul" },
                { v: 8, n: "Aug" },
                { v: 9, n: "Sep" },
                { v: 10, n: "Okt" },
                { v: 11, n: "Nov" },
                { v: 12, n: "Dez" },
              ].map((m) => (
                <option key={m.v} value={m.v}>
                  {m.n}
                </option>
              ))}
            </select>
          </div>
        </div>

        {this.state.viewMode === "items" && (
          <p>
            Ergebnisse: {this.state.visibleItems.length}
            {this.state.selectedBucket && (
              <>
                {" | "}
                <strong>{this.state.selectedBucket}</strong>
              </>
            )}
          </p>
        )}
        {this.state.resultMode === "folders" ? (
          <div className={styles.grid}>
            {this.getFilteredBuckets()
              .slice(0, this.state.bucketsToShow)
              .map((bucket) => {
                const preview = this.getBucketPreview(bucket);

                const fileType = preview?.name?.split(".").pop()?.toLowerCase();
                const isVideo = fileType === "mp4" || fileType === "mov";
                const isAudio = [
                  "mp3",
                  "wav",
                  "aiff",
                  "aac",
                  "flac",
                  "ogg",
                  "m4a",
                ].includes(fileType || "");

                const imageUrl = preview
                  ? `${window.location.origin}/_layouts/15/getpreview.ashx?path=${encodeURIComponent(preview.fileRef)}`
                  : "";

                return (
                  <div
                    key={bucket}
                    className={styles.bucketCard}
                    onClick={() =>
                      this.setState(
                        {
                          viewMode: "items",
                          resultMode: "files",
                          selectedBucket: bucket,
                          visibleItemsCount: 20,
                        },
                        this.applyFilters,
                      )
                    }
                  >
                    {/* +++++++++++ UPLOAD BUTTON BOTTOM RIGHT +++++++++++ */}
                    <button
                      className={styles.bucketUploadBtn}
                      onClick={(e) => {
                        e.stopPropagation();
                        e.preventDefault(); // ✅ WICHTIG!

                        this.setState({
                          isUploadOpen: true,
                          uploadBucket: [bucket],
                        });
                      }}
                    >
                      <span className={styles.plusIcon}>+</span>
                    </button>

                    {preview && !isVideo && !isAudio && (
                      <img
                        src={imageUrl}
                        className={styles.previewImg}
                        onError={(e) => {
                          e.currentTarget.src =
                            "data:image/svg+xml;charset=UTF-8," +
                            encodeURIComponent(`
          <svg width="250" height="150" xmlns="http://www.w3.org/2000/svg">
            <rect width="100%" height="100%" fill="#f3f2f1"/>
            <text x="50%" y="50%" dominant-baseline="middle" text-anchor="middle" fill="#605e5c">
              Kein Preview
            </text>
          </svg>
        `);
                        }}
                      />
                    )}

                    {preview && isVideo && (
                      <div
                        className={styles.itemImg}
                        style={{
                          display: "flex",
                          alignItems: "center",
                          justifyContent: "center",
                          fontSize: "50px",
                          background: "#f3f2f1",
                          cursor: "pointer",
                        }}
                      >
                        🎬
                      </div>
                    )}
                    {preview && isAudio && (
                      <div
                        className={styles.itemImg}
                        style={{
                          display: "flex",
                          alignItems: "center",
                          justifyContent: "center",
                          fontSize: "50px",
                          background: "#f3f2f1",
                          cursor: "pointer",
                        }}
                      >
                        🔊
                      </div>
                    )}

                    <div className={styles.bucketContent}>
                      <h3>{bucket}</h3>
                      <p className={styles.bucketCount}>
                        {bucketCounts[bucket] || 0} Dateien
                      </p>
                    </div>
                  </div>
                );
              })}
          </div>
        ) : (
          <div className={styles.grid}>
            {/* **************** MEDIA GRID **************** */}

            {this.state.visibleItems
              .slice(0, this.state.visibleItemsCount)
              .map((item) => {
                const fileUrl = `${window.location.origin}${item.fileRef}`;
                const downloadUrl = `${window.location.origin}${item.fileRef}`;

                const fileType = item.name?.split(".").pop()?.toLowerCase();

                const isVideo = fileType === "mp4" || fileType === "mov";
                const isAudio = [
                  "mp3",
                  "wav",
                  "aiff",
                  "aac",
                  "flac",
                  "ogg",
                  "m4a",
                ].includes(fileType || "");

                if (isAudio) {
                  console.log("Audio URL:", fileUrl);
                }
                return (
                  <div key={item.id} className={styles.itemCard}>
                    {isVideo && (
                      <div
                        className={styles.itemImg}
                        onClick={() =>
                          this.setState({
                            isModalOpen: true,
                            selectedItem: item,
                          })
                        }
                        style={{
                          display: "flex",
                          alignItems: "center",
                          justifyContent: "center",
                          fontSize: "50px",
                          background: "#f3f2f1",
                          cursor: "pointer",
                        }}
                      >
                        🎬
                      </div>
                    )}

                    {isAudio && (
                      <>
                        <div
                          className={styles.itemImg}
                          onClick={() =>
                            this.setState({
                              isModalOpen: true,
                              selectedItem: item,
                            })
                          }
                          style={{
                            display: "flex",
                            alignItems: "center",
                            justifyContent: "center",
                            fontSize: "50px",
                            background: "#f3f2f1",
                            cursor: "pointer",
                          }}
                        >
                          🔊
                        </div>
                      </>
                    )}
                    {/* IMAGE */}
                    {!isVideo && !isAudio && (
                      <img
                        src={fileUrl}
                        className={styles.modalMedia}
                        onClick={() =>
                          this.setState({
                            isModalOpen: true,
                            selectedItem: item,
                          })
                        }
                      />
                    )}

                    <div className={styles.itemContent}>
                      <h3>{item.name}</h3>
                      <p>Ersteller: {item.createdBy || "-"}</p>

                      <p>Kategorie: {item.category || "-"}</p>

                      <p>Dienst: {item.dienst || "-"}</p>

                      {/* ✅ TAG CHIPS HIER */}
                      <div className={styles.tagList}>
                        {(item.tags || []).map((tag, i) => (
                          <span
                            key={i}
                            onClick={() =>
                              this.setState(
                                { searchText: tag.toLowerCase() },
                                this.applyFilters,
                              )
                            }
                            className={styles.tag}
                          >
                            {tag}
                          </span>
                        ))}
                      </div>
                      <p className={styles.itemDate}>
                        Erstellt am:{" "}
                        {item.created
                          ? new Date(item.created).toLocaleDateString()
                          : "-"}
                      </p>
                      <div className={styles.itemActions}>
                        {/* DOWNLOAD */}
                        <button
                          onClick={() => {
                            this.setState({ downloadingItemId: item.id });

                            const link = document.createElement("a");
                            link.href = downloadUrl;
                            link.target = "_blank";
                            link.download = item.name;

                            document.body.appendChild(link);
                            link.click();
                            document.body.removeChild(link);

                            setTimeout(() => {
                              this.setState({ downloadingItemId: undefined });
                            }, 1500);
                          }}
                          className={styles.downloadBtn}
                        >
                          {this.state.downloadingItemId === item.id
                            ? "⏳ Lädt..."
                            : "Download"}
                        </button>

                        {/* EDIT BUTTON (neu) */}
                        <button
                          onClick={() => {
                            this.setState({
                              isEditOpen: true,
                              selectedItem: item,
                              editName: item.name || "",
                              editTags: item.tags || [],
                              editCategory: item.category || "",
                              editFormat: item.format || "",
                              editBucket: item.bucket || [],
                            });
                          }}
                          className={styles.editBtn}
                        >
                          Editieren
                        </button>
                      </div>
                    </div>
                  </div>
                );
              })}
          </div>
        )}

        {/* MEHR LADEN ORDNER */}
        {this.state.resultMode === "folders" &&
          this.state.bucketsToShow < this.getFilteredBuckets().length && (
            <button
              onClick={() =>
                this.setState({
                  bucketsToShow: this.state.bucketsToShow + 5,
                })
              }
              className={styles.loadMoreBtn}
            >
              Mehr laden
            </button>
          )}

        {/* MEHR LADEN DATEIEN */}
        {this.state.resultMode === "files" &&
          this.state.visibleItemsCount < this.state.visibleItems.length && (
            <button
              onClick={() => {
                this.setState({
                  visibleItemsCount: this.state.visibleItemsCount + 20,
                });
              }}
              className={styles.loadMoreBtn}
            >
              Mehr laden
            </button>
          )}

        {/* **************** PREVIEW MODAL **************** */}
        {this.state.isModalOpen && this.state.selectedItem && (
          <div
            className={`${styles.modalOverlay} ${styles.modalOverlayPreview}`}
          >
            {/* CLOSE BUTTON */}
            <button
              onClick={() =>
                this.setState({ isModalOpen: false, selectedItem: undefined })
              }
              className={styles.modalClose}
            >
              ✕
            </button>
            {/* CONTENT MODAL */}
            <div className={styles.modalContent}>
              {(() => {
                const item = this.state.selectedItem;
                const fileType = item?.name.split(".").pop()?.toLowerCase();
                const isVideo = fileType === "mp4" || fileType === "mov";
                const isAudio = [
                  "mp3",
                  "wav",
                  "aiff",
                  "aac",
                  "flac",
                  "ogg",
                  "m4a",
                  "wma",
                ].includes(fileType || "");

                const fileUrl = `${window.location.origin}${item?.fileRef}`;
                const downloadUrl = `${window.location.origin}${item?.fileRef}`;

                return (
                  <>
                    {/* IMAGE */}
                    {!isVideo && !isAudio && (
                      <img src={fileUrl} className={styles.modalMedia} />
                    )}

                    {/* VIDEO */}
                    {isVideo && (
                      <video
                        controls
                        style={{
                          maxWidth: "100%",
                          maxHeight: "80vh",
                        }}
                      >
                        <source src={fileUrl} />
                      </video>
                    )}

                    {isAudio && (
                      <audio
                        controls
                        style={{
                          minWidth: "200px",
                          marginTop: "20px",
                        }}
                      >
                        <source src={fileUrl} />
                      </audio>
                    )}

                    {/* DOWNLOAD BUTTON */}
                    <button
                      onClick={() => {
                        const link = document.createElement("a");

                        link.href = downloadUrl;
                        link.target = "_blank";
                        link.download = item?.name || "file";

                        document.body.appendChild(link);
                        link.click();
                        document.body.removeChild(link);
                      }}
                      className={`${styles.downloadBtn} ${styles.modalDownload}`}
                    >
                      Download
                    </button>
                  </>
                );
              })()}
            </div>
          </div>
        )}
        {/* **************** EDIT MODAL **************** */}
        {this.state.isEditOpen && this.state.selectedItem && (
          <div className={`${styles.modalOverlay} ${styles.modalOverlayEdit}`}>
            <div className={styles.modalBox}>
              <h3>Element bearbeiten</h3>
              {/* PREVIEW */}
              {(() => {
                const item = this.state.selectedItem!;
                const fileUrl = `${window.location.origin}${item.fileRef}`;

                const fileType = item.fileRef?.split(".").pop()?.toLowerCase();

                const isImage = ["jpg", "jpeg", "png", "gif", "webp"].includes(
                  fileType || "",
                );
                const isVideo = ["mp4", "mov", "webm"].includes(fileType || "");

                if (isImage) {
                  return <img src={fileUrl} className={styles.editPreview} />;
                }

                if (isVideo) {
                  return (
                    <video
                      src={fileUrl}
                      controls
                      className={styles.editPreview}
                    />
                  );
                }

                return (
                  <div className={styles.editFallback}>📄 {item.name}</div>
                );
              })()}
              <input
                type="text"
                value={this.state.editName}
                onChange={(e) => this.setState({ editName: e.target.value })}
                placeholder="Name"
              />
              {/* TAGS */}
              <div>
                <div className={styles.tagList}>
                  {this.state.editTags.map((tag, index) => (
                    <span
                      key={index}
                      className={`${styles.tag} ${styles.editTag}`}
                    >
                      {tag}
                      <span
                        onClick={() => {
                          const newTags = [...this.state.editTags];
                          newTags.splice(index, 1);
                          this.setState({ editTags: newTags });
                        }}
                        className={styles.tagRemove}
                      >
                        ✕
                      </span>
                    </span>
                  ))}
                </div>

                <input
                  type="text"
                  className={styles.tagInput}
                  placeholder="Tag hinzufügen + Enter"
                  onKeyDown={(e) => {
                    if (e.key === "Enter") {
                      e.preventDefault();
                      const value = (e.target as HTMLInputElement).value.trim();
                      if (!value) return;

                      this.setState({
                        editTags: [...this.state.editTags, value],
                      });

                      (e.target as HTMLInputElement).value = "";
                    }
                  }}
                />
              </div>
              <select
                value={this.state.editCategory}
                onChange={(e) =>
                  this.setState({ editCategory: e.target.value })
                }
              >
                <option value="">Kategorie wählen</option>

                {this.state.categoryOptions.map((cat) => (
                  <option key={cat} value={cat}>
                    {cat}
                  </option>
                ))}
              </select>
              <select
                value={this.state.editFormat}
                onChange={(e) => this.setState({ editFormat: e.target.value })}
              >
                <option value="">Format wählen</option>
                <option value="Bild">Bild</option>
                <option value="Video">Video</option>
                <option value="Audio">Audio</option>
              </select>
              <BucketDropdown
                options={this.state.bucketOptions}
                selected={this.state.editBucket}
                onChange={(values) =>
                  this.setState({
                    editBucket: values,
                  })
                }
              />

              <div className={styles.modalActions}>
                <button
                  onClick={() => this.setState({ isEditOpen: false })}
                  style={{ flex: 1 }}
                >
                  Abbrechen
                </button>

                <button
                  onClick={() => this.updateItem()}
                  className={styles.btnFlex}
                >
                  Speichern
                </button>

                <button
                  onClick={() => this.deleteItem()}
                  className={styles.btnDelete}
                >
                  Löschen
                </button>
              </div>
            </div>
          </div>
        )}
        {/* **************** UPLOAD MODAL **************** */}

        <UploadModal
          isOpen={this.state.isUploadOpen}
          state={this.state}
          setState={(newState: Partial<IMediaAssetsLibState>) =>
            this.setState(
              newState as Pick<
                IMediaAssetsLibState,
                keyof IMediaAssetsLibState
              >,
            )
          }
          onClose={() => this.setState({ isUploadOpen: false })}
          onUpload={() => this.uploadItem()}
        />

        {this.state.showScrollTop && (
          <a href="#top">
            <div
              className={`${styles.scrollTop} ${
                this.state.showScrollTop ? styles.visible : styles.hidden
              }`}
            >
              ↑
            </div>
          </a>
        )}
      </div>
    );
  }
}
