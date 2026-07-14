import * as React from "react";
import { SPHttpClient } from "@microsoft/sp-http";
import type { IMediaAssetsLibProps } from "./IMediaAssetsLibProps";
import styles from "./MediaAssetsLib.module.scss";

import UploadModal from "./UploadModal";
import FilterBar from "./FilterBar";
import FileCard from "./FileCard";
import PreviewModal from "./PreviewModal";
import EditModal from "./EditModal";

import { detectFormat } from "../utils/mediaUtils";
import { IMediaItem, ISPFile } from "../models/types";
import {
  getUniqueFormats,
  getUniqueYears,
  getUniqueCreators,
} from "../utils/mediaHelpers";
import {
  getRequestDigest,
  getCurrentUser,
  getChoiceFieldOptions,
} from "../services/spService";
import {
  getBucketCounts,
  getBucketsSortedByNewest,
  getBucketPreview,
  getFilteredBuckets,
} from "../utils/bucketHelpers";

import { getLibraryPath } from "../utils/sharepointHelpers";

/*******************************************************
 * MEDIA ASSETS LIB V11.1
 * -----------------------------------------------------
 * SharePoint Medienverwaltung
 * - Upload
 * - Edit
 * - Delete
 * - Filter
 * - Bucket Dropdown
 *******************************************************/

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

  editingBucket?: string;
  newBucketName: string;

  isRenameBucketOpen: boolean;
  bucketToRename?: string;
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
      editBucket: [],
      editDienst: "",

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

      editingBucket: undefined,
      newBucketName: "",
      isRenameBucketOpen: false,
      bucketToRename: undefined,
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
      const categories = await getChoiceFieldOptions(
        this.props.siteUrl,
        this.libraryName,
        "Kategorie",
        this.props.spHttpClient,
      );

      this.setState({
        categoryOptions: categories,
      });
    } catch (error) {
      console.error("Fehler beim Laden der Kategorien", error);
    }
  }

  /**git
   * Lädt die verfügbaren Dienstoptionen aus dem Dienste-Feld der SharePoint-Liste
   * Speichert die Optionen im State für die Dienste-Filter-Dropdown
   * @private
   * @async
   * @returns {Promise<void>}
   * @throws {Error} Wenn der REST-API-Aufruf fehlschlägt
   */
  private async loadDienste(): Promise<void> {
    try {
      const dienste = await getChoiceFieldOptions(
        this.props.siteUrl,
        this.libraryName,
        "Dienste",
        this.props.spHttpClient,
      );

      this.setState({
        dienstOptions: dienste,
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
    const digest = await getRequestDigest(
      this.props.siteUrl,
      this.props.spHttpClient,
    );
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
          `${this.props.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${getLibraryPath(
            this.props.siteUrl,
            this.libraryName,
          )}')` + `/Files/add(overwrite=true,url='${uploadFile.name}')`;

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

        const detectedFormat = detectFormat(uploadFile.name);
        const currentUser = await getCurrentUser(
          this.props.siteUrl,
          this.props.spHttpClient,
        );
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
    await this.loadVideoThumbnails();

    this.setState({
      isUploading: false,
      uploadProgress: 0,
      uploadCurrentFile: 0,
      uploadTotalFiles: 0,
    });

    const targetBucket =
      uploadBucket && uploadBucket.length > 0 ? uploadBucket[0] : undefined;

    this.setState(
      {
        isUploadOpen: false,

        uploadName: "",
        uploadTags: [],
        uploadCategory: "",
        uploadDienst: "",
        uploadBucket: [],
        uploadFiles: [],
        uploadPreviewUrl: undefined,

        viewMode: targetBucket ? "items" : "buckets",
        resultMode: targetBucket ? "files" : "folders",
        selectedBucket: targetBucket,

        searchText: "",
      },
      this.applyFilters,
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
    const { selectedItem } = this.state;

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

        await this.loadAllMedia();
        await this.loadVideoThumbnails();

        this.setState({
          isEditOpen: false,
          selectedItem: undefined,
        });
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
    await this.loadAllMedia();
    await this.loadVideoThumbnails();
    await this.loadCategories();
    await this.loadDienste();

    const target = document.getElementById("top");

    if (target) {
      this.observer = new IntersectionObserver(
        ([entry]) => {
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
Files/ListItemAllFields/Ersteller,
Files/UniqueId`;
    const response = await this.props.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
    );

    const data = await response.json();

    let results: IMediaItem[] = [];

    data.Files.forEach((file: unknown): void => {
      const f = file as ISPFile;

      results.push({
        id: f.ListItemAllFields.Id,
        uniqueId: f.UniqueId,
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
      const rootFolder = getLibraryPath(this.props.siteUrl, this.libraryName);
      const items = await this.getFolderContent(rootFolder);

      const uniqueBuckets = getBucketsSortedByNewest(items);
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
  private async loadVideoThumbnails(): Promise<void> {
    console.log("THUMB ITEMS", this.state.allItems.length);
    const drivesResponse = await this.props.spHttpClient.get(
      `${this.props.siteUrl}/_api/v2.1/drives`,
      SPHttpClient.configurations.v1,
    );

    const drivesData = await drivesResponse.json();

    const mediaDrive = drivesData.value.find(
      (d: { name: string; id: string }) => d.name === "Medienbibliothek",
    );

    if (!mediaDrive) {
      return;
    }

    const childrenResponse = await this.props.spHttpClient.get(
      `${this.props.siteUrl}/_api/v2.1/drives/${mediaDrive.id}/root/children`,
      SPHttpClient.configurations.v1,
    );

    const childrenData = await childrenResponse.json();

    const driveItemMap: { [key: string]: string } = {};

    childrenData.value.forEach((item: { id: string; name: string }) => {
      driveItemMap[item.name] = item.id;
    });

    const updatedItems = await Promise.all(
      this.state.allItems.map(async (item) => {
        const driveItemId = driveItemMap[item.name];

        if (!item.name.toLowerCase().endsWith(".mp4") || !driveItemId) {
          return item;
        }

        try {
          const thumbnailResponse = await this.props.spHttpClient.get(
            `${this.props.siteUrl}/_api/v2.1/drives/${mediaDrive.id}/items/${driveItemId}/thumbnails`,
            SPHttpClient.configurations.v1,
          );

          const thumbnailData = await thumbnailResponse.json();

          const thumbUrl = thumbnailData.value?.[0]?.medium?.url;

          return {
            ...item,
            driveId: mediaDrive.id,
            driveItemId,
            thumbnailUrl: thumbUrl,
          };
        } catch {
          return item;
        }
      }),
    );

    this.setState(
      {
        allItems: updatedItems,
      },
      this.applyFilters,
    );

    const videoCount = updatedItems.filter((i) => i.thumbnailUrl).length;

    console.log("ITEMS MIT THUMBNAIL", videoCount);
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
      editDienst,
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
            Dienste: editDienst,
            body: JSON.stringify({
              FileLeafRef: editName,

              Kategorie: editCategory,
              Dienste: editDienst,

              Tags: tagsArray,

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
            Dienste: editDienst,
            Tags: tagsArray,
            Bucket: bucketsArray,
          }),
        },
      );

      const errorText = await response.text();
      console.log("🧪 UPDATE RESPONSE:", errorText);

      console.log("Update erfolgreich");

      // neu laden → damit UI sofort aktualisiert wird
      await this.loadAllMedia();
      await this.loadVideoThumbnails();

      this.setState({ isEditOpen: false });
    } catch (error) {
      console.error("Update Fehler:", error);
    }
  }

  private async renameBucket(): Promise<void> {
    const { bucketToRename, newBucketName, allItems } = this.state;

    if (!bucketToRename) {
      return;
    }

    if (!newBucketName.trim()) {
      return;
    }

    const alreadyExists = this.state.bucketOptions.some(
      (bucket) =>
        bucket.toLowerCase().trim() === newBucketName.toLowerCase().trim(),
    );

    if (
      alreadyExists &&
      bucketToRename.toLowerCase().trim() !== newBucketName.toLowerCase().trim()
    ) {
      alert("Ein Ordner mit diesem Namen existiert bereits.");
      return;
    }

    try {
      const affectedItems = allItems.filter((item) =>
        item.bucket?.includes(bucketToRename),
      );

      for (const item of affectedItems) {
        const updatedBuckets =
          item.bucket?.map((bucket) =>
            bucket === bucketToRename ? newBucketName.trim() : bucket,
          ) || [];

        const url = `${this.props.siteUrl}/_api/web/lists/getbytitle('${this.libraryName}')/items(${item.id})`;

        await this.props.spHttpClient.post(
          url,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: "application/json;odata=nometadata",
              "Content-Type": "application/json;odata=nometadata",
              "IF-MATCH": "*",
              "X-HTTP-Method": "MERGE",
            },
            body: JSON.stringify({
              Bucket: updatedBuckets,
            }),
          },
        );
      }

      await this.loadAllMedia();

      this.setState({
        isRenameBucketOpen: false,
        bucketToRename: undefined,
        newBucketName: "",
      });
    } catch (error) {
      console.error("Bucket umbenennen fehlgeschlagen", error);
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
    const categoryOptions = this.state.categoryOptions;

    const dienstOptions = this.state.dienstOptions;

    const yearOptions = getUniqueYears(this.state.allItems);

    const formatOptions = getUniqueFormats(this.state.allItems);

    const creatorOptions = getUniqueCreators(this.state.allItems);

    const bucketCounts = getBucketCounts(this.state.allItems);

    const bucketNameExists = this.state.bucketOptions.some(
      (bucket) =>
        bucket.toLowerCase().trim() ===
          this.state.newBucketName.toLowerCase().trim() &&
        bucket !== this.state.bucketToRename,
    );

    const renameDisabled = !this.state.newBucketName.trim() || bucketNameExists;

    const filteredBuckets = getFilteredBuckets(
      this.state.bucketOptions,
      this.state.allItems,
      {
        searchText: this.state.searchText,
        filterCategory: this.state.filterCategory,
        filterDienst: this.state.filterDienst,
        filterFormat: this.state.filterFormat,
        filterYear: this.state.filterYear,
        filterMonth: this.state.filterMonth,
        filterCreator: this.state.filterCreator,
      },
    );
    /********************* RENDER **************************/

    return (
      <div className={styles.container}>
        {/* **************** HEADER **************** */}
        <h2 id="top">Caritas Media Library</h2>

        <div className={styles.header}>
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
            <FilterBar
              filterCategory={this.state.filterCategory}
              filterDienst={this.state.filterDienst}
              filterCreator={this.state.filterCreator}
              filterFormat={this.state.filterFormat}
              filterYear={this.state.filterYear}
              filterMonth={this.state.filterMonth}
              categoryOptions={categoryOptions}
              dienstOptions={dienstOptions}
              creatorOptions={creatorOptions}
              formatOptions={formatOptions}
              yearOptions={yearOptions}
              onChange={(values) => this.setState(values, this.applyFilters)}
            />
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
            {filteredBuckets
              .slice(0, this.state.bucketsToShow)
              .map((bucket) => {
                const preview = getBucketPreview(this.state.allItems, bucket);
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
                      <img
                        src={preview.thumbnailUrl}
                        className={styles.videoImg}
                        style={{ cursor: "pointer" }}
                      />
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
                      <div className={styles.bucketHeader}>
                        <h3>{bucket}</h3>

                        <button
                          className={styles.bucketEditBtn}
                          onClick={(e) => {
                            e.stopPropagation();

                            this.setState({
                              isRenameBucketOpen: true,
                              bucketToRename: bucket,
                              newBucketName: bucket,
                            });
                          }}
                        >
                          ✏️
                        </button>
                      </div>

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
              .map((item) => (
                <FileCard
                  key={item.id}
                  item={item}
                  downloadingItemId={this.state.downloadingItemId}
                  onPreview={() =>
                    this.setState({
                      isModalOpen: true,
                      selectedItem: item,
                    })
                  }
                  onEdit={() =>
                    this.setState({
                      isEditOpen: true,
                      selectedItem: item,
                      editName: item.name || "",
                      editTags: item.tags || [],
                      editCategory: item.category || "",
                      editBucket: item.bucket || [],
                      editDienst: item.dienst || "",
                    })
                  }
                  onDownload={() => {
                    this.setState({
                      downloadingItemId: item.id,
                    });

                    const downloadUrl = `${window.location.origin}${item.fileRef}`;

                    const link = document.createElement("a");

                    link.href = downloadUrl;
                    link.download = item.name;
                    link.target = "_blank";

                    document.body.appendChild(link);

                    link.click();

                    document.body.removeChild(link);

                    setTimeout(() => {
                      this.setState({
                        downloadingItemId: undefined,
                      });
                    }, 1500);
                  }}
                />
              ))}
          </div>
        )}

        {/* MEHR LADEN ORDNER */}
        {this.state.resultMode === "folders" &&
          this.state.bucketsToShow < filteredBuckets.length && (
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
        <PreviewModal
          isOpen={this.state.isModalOpen}
          item={this.state.selectedItem}
          onClose={() =>
            this.setState({
              isModalOpen: false,
              selectedItem: undefined,
            })
          }
        />
        {/* **************** EDIT MODAL **************** */}
        <EditModal
          isOpen={this.state.isEditOpen}
          item={this.state.selectedItem}
          editName={this.state.editName}
          editTags={this.state.editTags}
          editCategory={this.state.editCategory}
          editDienst={this.state.editDienst}
          editBucket={this.state.editBucket}
          categoryOptions={this.state.categoryOptions}
          dienstOptions={this.state.dienstOptions}
          bucketOptions={this.state.bucketOptions}
          onClose={() =>
            this.setState({
              isEditOpen: false,
            })
          }
          onSave={() => this.updateItem()}
          onDelete={() => this.deleteItem()}
          onNameChange={(value) =>
            this.setState({
              editName: value,
            })
          }
          onCategoryChange={(value) =>
            this.setState({
              editCategory: value,
            })
          }
          onDienstChange={(value) =>
            this.setState({
              editDienst: value,
            })
          }
          onBucketChange={(values) =>
            this.setState({
              editBucket: values,
            })
          }
          onTagsChange={(values) =>
            this.setState({
              editTags: values,
            })
          }
        />
        {/* **************** UPLOAD MODAL **************** */}
        {this.state.isRenameBucketOpen && (
          <div className={styles.modalOverlay}>
            <div className={styles.renameModal}>
              <h3>Ordner umbenennen</h3>

              <input
                type="text"
                value={this.state.newBucketName}
                onChange={(e) =>
                  this.setState({
                    newBucketName: e.target.value,
                  })
                }
              />
              {bucketNameExists && (
                <p className={styles.errorText}>
                  Ein Ordner mit diesem Namen existiert bereits.
                </p>
              )}

              <div className={styles.modalActions}>
                <button
                  onClick={() =>
                    this.setState({
                      isRenameBucketOpen: false,
                      bucketToRename: undefined,
                      newBucketName: "",
                    })
                  }
                >
                  Abbrechen
                </button>

                <button
                  disabled={renameDisabled}
                  onClick={() => this.renameBucket()}
                >
                  Speichern
                </button>
              </div>
            </div>
          </div>
        )}

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
          onClose={() =>
            this.setState({
              isUploadOpen: false,

              uploadName: "",
              uploadTags: [],
              uploadCategory: "",
              uploadDienst: "",
              uploadBucket: [],
              uploadFiles: [],
              uploadPreviewUrl: undefined,
            })
          }
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
