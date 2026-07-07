import * as React from "react";
import { SPHttpClient } from "@microsoft/sp-http";
import type { IMediaAssetsLibProps } from "./IMediaAssetsLibProps";
import styles from "./MediaAssetsLib.module.scss";

import BucketDropdown from "./BucketDropdown";

/*******************************************************
 * MEDIA ASSETS LIB V7
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
}

interface ISPFile {
  Name: string;
  ServerRelativeUrl: string;
  TimeCreated: string;

  ListItemAllFields: {
    Id: number;
    Kategorie?: string;
    Notizen?: string;
    Tags?: string[] | string;
    Bucket?: string[] | string;
    Format?: string;

    File?: {
      VroomDriveID?: string;
      VroomItemID?: string;
    };
  };
}
/********************* STATE ***************************
 * Enthält alle UI Zustände (Filter, Modals, Form Daten)
 ******************************************************/
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

  editBucket: string[];
  uploadBucket: string[];

  isUploadOpen: boolean; // ✅ NEU
  uploadName: string;
  uploadTags: string[];
  uploadCategory: string;
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
}

/********************* HAUPTKOMPNENTE ******************
 * Steuert:
 * - Datenladen
 * - Upload
 * - Update
 * - Rendering
 ******************************************************/

export default class MediaAssetsLib extends React.Component<
  IMediaAssetsLibProps,
  IMediaAssetsLibState
> {
  /********************* INITIAL STATE *******************/
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
    };
  }

  /* CHANGE SHAREPOINT URL */
  private readonly libraryName = "Medienbibliothek";

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

  private getLibraryPath(): string {
    return `${new URL(this.props.siteUrl).pathname}/${this.libraryName}`;
  }

  private observer?: IntersectionObserver;

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

  private getBucketPreview(bucket: string): IMediaItem | undefined {
    return this.state.allItems.find((item) => item.bucket?.includes(bucket));
  }

  private getFilteredBuckets(): string[] {
    const {
      searchText,
      filterCategory,
      filterFormat,
      filterYear,
      filterMonth,
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

        // FORMAT
        const matchesFormat = !filterFormat || item.format === filterFormat;

        // DATUM
        const date = item.created ? new Date(item.created) : null;

        const matchesYear =
          !filterYear || (date && date.getFullYear() === filterYear);

        const matchesMonth =
          !filterMonth || (date && date.getMonth() + 1 === filterMonth);

        return (
          matchesText &&
          matchesCategory &&
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

  /********************* UPLOAD **************************
   * Lädt Dateien hoch und setzt Metadaten
   ******************************************************/

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

  private async uploadItem(): Promise<void> {
    const { uploadFiles, uploadTags, uploadCategory, uploadBucket } =
      this.state;

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

        const bodyData = {
          FileLeafRef: finalName,
          Kategorie: uploadCategory || null,
          Tags: cleanTags,
          Format: detectedFormat, // ✅ AUTO!
          Bucket: cleanBuckets,
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
                  filterFormat: undefined,
                  filterYear: undefined,
                  filterMonth: undefined,
                }
              : // ✅ FALL 2: Ordner hat noch Dateien → drin bleiben
                {
                  viewMode: "items",
                  resultMode: "files",
                  selectedBucket: selectedBucket,
                  searchText: "",
                  filterCategory: undefined,
                  filterFormat: undefined,
                  filterYear: undefined,
                  filterMonth: undefined,
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

  public async componentDidMount(): Promise<void> {
    console.log("MOUNTED ✅");

    console.log("Aktuelle Site:");
    console.log(this.props.siteUrl);

    await this.loadAllMedia();
    await this.loadCategories();

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

  private async getFolderContent(folderUrl: string): Promise<IMediaItem[]> {
    const url = `${this.props.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')?$expand=Folders,Files,Files/ListItemAllFields&$select=Name,ServerRelativeUrl,TimeCreated,Files/Name,Files/ServerRelativeUrl,Files/TimeCreated,Files/ListItemAllFields/Id,Files/ListItemAllFields/Kategorie,Files/ListItemAllFields/Notizen,Files/ListItemAllFields/UniqueId,Files/ListItemAllFields/Tags,Files/ListItemAllFields/Format,Files/ListItemAllFields/Bucket`;

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
        name: f.Name,
        fileRef: f.ServerRelativeUrl,
        category: f.ListItemAllFields?.Kategorie,
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

  private applyFilters = (): void => {
    const {
      allItems,
      searchText,
      filterCategory,
      filterFormat,
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

    if (filterFormat) {
      filtered = filtered.filter((item) => item.format === filterFormat);
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

  private getUniqueCategories(): string[] {
    const values = this.state.allItems
      .map((item) => item.category)
      .filter((v) => v);
    console.log("ALL ITEMS", this.state.allItems);
    console.log("CATEGORIES", values);
    return Array.from(new Set(values)) as string[];
  }

  private getUniqueFormats(): string[] {
    const values = this.state.allItems
      .map((item) => item.format)
      .filter((v): v is string => !!v);

    return Array.from(new Set(values));
  }

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
  /********************* BUCKET LADEN ********************
   * Lädt Choice Werte aus SharePoint
   ******************************************************/

  public render(): React.ReactElement<IMediaAssetsLibProps> {
    const categoryOptions = this.getUniqueCategories();

    const yearOptions = this.getUniqueYears();

    const formatOptions = this.getUniqueFormats();

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
                  filterFormat: undefined,
                  filterYear: undefined,
                  filterMonth: undefined,
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
              onChange={(e) =>
                this.setState(
                  { filterCategory: e.target.value },
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
          <p>Ergebnisse: {this.state.visibleItems.length}</p>
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
                          searchText: "",
                          filterCategory: undefined,
                          filterFormat: undefined,
                          filterYear: undefined,
                          filterMonth: undefined,
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
        {this.state.isUploadOpen && (
          <div
            className={`${styles.modalOverlay} $${styles.modalOverlayUpload}`}
          >
            <div className={`${styles.modalBox} ${styles.uploadBox}`}>
              <h3>Upload</h3>
              {/* MODAL PREVIEW */}
              {this.state.uploadFiles?.[0] && this.state.uploadPreviewUrl && (
                <div className={styles.uploadPreview}>
                  {(() => {
                    const file = this.state.uploadFiles![0];
                    const fileType = file.name.split(".").pop()?.toLowerCase();

                    const isImage = [
                      "jpg",
                      "jpeg",
                      "png",
                      "gif",
                      "webp",
                    ].includes(fileType || "");
                    const isVideo = ["mp4", "mov", "webm"].includes(
                      fileType || "",
                    );

                    if (isImage) {
                      return (
                        <img
                          src={this.state.uploadPreviewUrl}
                          className={styles.uploadPreviewMedia}
                        />
                      );
                    }

                    if (isVideo) {
                      return (
                        <video
                          src={this.state.uploadPreviewUrl}
                          controls
                          className={styles.uploadPreviewMedia}
                        />
                      );
                    }

                    return (
                      <div className={styles.uploadFallback}>
                        📄 {file.name}
                      </div>
                    );
                  })()}
                </div>
              )}
              <input
                type="file"
                multiple
                onChange={(e) => {
                  const files = e.target.files;
                  if (!files || files.length === 0) return;

                  const fileArray = Array.from(files);

                  const firstFile = fileArray[0];

                  const fileName = firstFile.name;
                  const baseName = fileName.includes(".")
                    ? fileName.substring(0, fileName.lastIndexOf("."))
                    : fileName;

                  const previewUrl = URL.createObjectURL(firstFile);

                  this.setState({
                    uploadFiles: fileArray,
                    uploadName: baseName,
                    uploadPreviewUrl: previewUrl,
                  });
                }}
              />
              <div>{this.state.uploadFiles?.length} Dateien gewählt</div>
              {(!this.state.uploadFiles ||
                this.state.uploadFiles.length <= 1) && (
                <input
                  type="text"
                  placeholder="Name"
                  value={this.state.uploadName}
                  onChange={(e) =>
                    this.setState({ uploadName: e.target.value })
                  }
                />
              )}

              <BucketDropdown
                options={this.state.bucketOptions}
                selected={this.state.uploadBucket}
                onChange={(values) =>
                  this.setState({
                    uploadBucket: values,
                  })
                }
              />

              <select
                value={this.state.uploadCategory}
                onChange={(e) =>
                  this.setState({ uploadCategory: e.target.value })
                }
              >
                <option value="">Kategorie wählen</option>

                {this.state.categoryOptions.map((cat) => (
                  <option key={cat} value={cat}>
                    {cat}
                  </option>
                ))}
              </select>

              <div>
                <div className="styles.tagList">
                  {this.state.uploadTags.map((tag, index) => (
                    <span
                      key={index}
                      className={`${styles.tag} ${styles.editTag}`}
                    >
                      {tag}
                      <span
                        onClick={() => {
                          const newTags = [...this.state.uploadTags];
                          newTags.splice(index, 1);
                          this.setState({ uploadTags: newTags });
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
                  placeholder="Tag hinzufügen + Enter"
                  className={styles.tagInput}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") {
                      e.preventDefault();

                      const value = (e.target as HTMLInputElement).value.trim();
                      if (!value) return;

                      this.setState({
                        uploadTags: [...this.state.uploadTags, value],
                      });

                      (e.target as HTMLInputElement).value = "";
                    }
                  }}
                />
              </div>

              <button onClick={() => this.setState({ isUploadOpen: false })}>
                Schließen
              </button>
              {this.state.isUploading && (
                <div style={{ marginTop: 20 }}>
                  <div>
                    Datei {this.state.uploadCurrentFile} von{" "}
                    {this.state.uploadTotalFiles}
                  </div>

                  <div
                    style={{
                      width: "100%",
                      height: "12px",
                      background: "#ddd",
                      marginTop: "10px",
                    }}
                  >
                    <div
                      style={{
                        width: `${this.state.uploadProgress}%`,
                        height: "100%",
                        background: "#ecdd04",
                        transition: "width 0.2s ease",
                      }}
                    />
                  </div>

                  <div style={{ marginTop: 6 }}>
                    {this.state.uploadProgress}%
                  </div>
                </div>
              )}
              <button
                onClick={() => this.uploadItem()}
                disabled={
                  this.state.isUploading ||
                  !this.state.uploadBucket ||
                  this.state.uploadBucket.length === 0 ||
                  !this.state.uploadCategory
                }
                className={`${styles.uploadBtn} ${
                  this.state.isUploading ||
                  !this.state.uploadBucket ||
                  this.state.uploadBucket.length === 0 ||
                  !this.state.uploadCategory
                    ? styles.disabled
                    : ""
                }`}
              >
                {this.state.isUploading
                  ? "⏳ Wird hochgeladen..."
                  : "Hochladen"}
              </button>
            </div>
          </div>
        )}
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
