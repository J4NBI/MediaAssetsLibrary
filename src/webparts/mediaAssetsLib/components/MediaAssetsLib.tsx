import * as React from "react";
import { SPHttpClient } from "@microsoft/sp-http";
import type { IMediaAssetsLibProps } from "./IMediaAssetsLibProps";
/*******************************************************
 * MEDIA ASSETS LIB V4
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
}

/********************* BUCKET DROPDOWN *****************
 * Custom Multi-Select Dropdown:
 * - Suche
 * - Mehrfachauswahl
 * - neue Buckets hinzufügen
 ******************************************************/
const BucketDropdown = ({
  options,
  selected,
  onChange,
}: {
  options: string[];
  selected: string[];
  onChange: (values: string[]) => void;
}) => {
  /**************** LOCAL STATE ****************/
  const [input, setInput] = React.useState("");
  const [open, setOpen] = React.useState(false);
  const ref = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) {
        setOpen(false);
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  /**************** FILTER LOGIK ****************/
  const filtered = options.filter((o) =>
    o.toLowerCase().includes(input.toLowerCase()),
  );

  /**************** ADD BUCKET ****************/
  const addValue = (value: string): void => {
    const clean = value.trim();
    if (!clean) return;

    if (!selected.some((s) => s.toLowerCase() === clean.toLowerCase())) {
      onChange([...selected, clean]);
    }

    setInput("");
    setOpen(false);
  };

  /**************** RENDER ****************/

  return (
    <div ref={ref} style={{ position: "relative", width: "100%" }}>
      {/* INPUT */}
      <label>Ordner</label>
      <input
        type="text"
        style={{ width: "100%" }}
        value={input}
        placeholder="Ordner wählen oder hinzufügen..."
        onFocus={() => setOpen(true)}
        onChange={(e) => {
          setInput(e.target.value);
          setOpen(true);
        }}
        onKeyDown={(e) => {
          if (e.key === "Enter") {
            e.preventDefault();
            addValue(input);
          }
        }}
      />

      {/* DROPDOWN */}
      {open && (
        <div
          style={{
            position: "absolute",
            top: "100%",
            left: 0,
            right: 0,
            background: "white",
            border: "1px solid #ccc",
            zIndex: 1000,
            maxHeight: "200px",
            overflow: "auto",
          }}
        >
          {/* NEU */}
          {input.trim() &&
            !options.some(
              (o) => o.toLowerCase() === input.trim().toLowerCase(),
            ) && (
              <div
                style={{
                  padding: "8px",
                  fontWeight: "bold",
                  cursor: "pointer",
                }}
                onClick={() => addValue(input)}
              >
                + "{input}" hinzufügen
              </div>
            )}

          {/* LISTE */}
          {filtered.map((o) => (
            <div
              key={o}
              style={{ padding: "8px", cursor: "pointer" }}
              onClick={() => addValue(o)}
            >
              {o}
            </div>
          ))}
        </div>
      )}

      {/* AUSWAHL */}
      <div style={{ marginTop: 6, display: "flex", flexWrap: "wrap", gap: 6 }}>
        {selected.map((b) => (
          <span
            key={b}
            style={{
              background: "#ecdd04",
              color: "black",
              padding: "4px 8px",
              borderRadius: "0px",
              cursor: "pointer",
            }}
            onClick={() => onChange(selected.filter((x) => x !== b))}
          >
            {b} ✕
          </span>
        ))}
      </div>
    </div>
  );
};

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
    };
  }

  private detectFormat(fileName: string): string {
    const ext = fileName.split(".").pop()?.toLowerCase();

    if (!ext) return "Dokument";

    const imageTypes = ["jpg", "jpeg", "png", "gif", "webp", "bmp", "svg"];
    const videoTypes = ["mp4", "mov", "avi", "webm", "mkv"];

    if (imageTypes.includes(ext)) return "Bild";
    if (videoTypes.includes(ext)) return "Video";

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

      // ✅ Bucket nur anzeigen wenn Treffer vorhanden
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

  /********************* UPLOAD **************************
   * Lädt Dateien hoch und setzt Metadaten
   ******************************************************/

  private async uploadItem(): Promise<void> {
    this.setState({ isUploading: true });
    const { uploadFiles, uploadTags, uploadCategory, uploadBucket } =
      this.state;

    if (!uploadBucket || uploadBucket.length === 0) {
      alert("Bitte wählen Sie mindestens einen Ordner aus.");
      this.setState({ isUploading: false }); // ✅ FIX!!
      return;
    }

    if (!uploadFiles || uploadFiles.length === 0) {
      console.log("❌ Keine Files gewählt");
      this.setState({ isUploading: false }); // ✅ FIX!!
      return;
    }

    for (const uploadFile of uploadFiles) {
      try {
        const fileBuffer = await uploadFile.arrayBuffer();

        const uploadUrl =
          this.props.siteUrl +
          "/_api/web/GetFolderByServerRelativeUrl('/sites/IntranetSpielwiese/Medienbibliothek')" +
          "/Files/add(overwrite=true,url='" +
          uploadFile.name +
          "')";

        console.log("⬆️ Starte Upload:", uploadFile.name);

        const response = await this.props.spHttpClient.post(
          uploadUrl,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: "application/json;odata=nometadata",
            },
            body: fileBuffer,
          },
        );

        const result = await response.json();
        const fileUrl = result.ServerRelativeUrl;

        const itemResponse = await this.props.spHttpClient.get(
          `${this.props.siteUrl}/_api/web/GetFileByServerRelativeUrl('${fileUrl}')/ListItemAllFields`,
          SPHttpClient.configurations.v1,
        );

        const itemData = await itemResponse.json();
        const itemId = itemData.Id;

        const updateUrl =
          this.props.siteUrl +
          "/_api/web/lists/getbytitle('Medienbibliothek')/items(" +
          itemId +
          ")";

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
      } catch (error) {
        console.error("❌ Fehler bei Datei:", uploadFile.name, error);
      }
    }

    await this.loadAllMedia();

    this.setState({ isUploading: false });

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
      const url = `${this.props.siteUrl}/_api/web/lists/getbytitle('Medienbibliothek')/items(${selectedItem.id})`;

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

        const rootFolder = "/sites/IntranetSpielwiese/Medienbibliothek";
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
    await this.loadAllMedia();
  }

  /********************* REKURSIVER LOAD *****************
   * Lädt alle Dateien + Unterordner
   ******************************************************/

  private async getFolderContent(folderUrl: string): Promise<IMediaItem[]> {
    /* const url = `${this.props.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')?$expand=Folders,Files/ListItemAllFields&$select=Name,ServerRelativeUrl,TimeCreated,ListItemAllFields/Id,ListItemAllFields/Kategorie,ListItemAllFields/Notizen,ListItemAllFields/UniqueId,ListItemAllFields/Tags`;*/

    const url = `${this.props.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')?$expand=Folders,Files,Files/ListItemAllFields&$select=Name,ServerRelativeUrl,TimeCreated,Files/Name,Files/ServerRelativeUrl,Files/TimeCreated,Files/ListItemAllFields/Id,Files/ListItemAllFields/Kategorie,Files/ListItemAllFields/Notizen,Files/ListItemAllFields/UniqueId,Files/ListItemAllFields/Tags,Files/ListItemAllFields/Format,Files/ListItemAllFields/Bucket`;

    const response = await this.props.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
    );

    const data = await response.json();

    let results: IMediaItem[] = [];

    data.Files.forEach((file: unknown) => {
      const f = file as any;

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
      const rootFolder = "/sites/IntranetSpielwiese/Medienbibliothek";

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
      const url = `${this.props.siteUrl}/_api/web/lists/getbytitle('Medienbibliothek')/items(${selectedItem.id})`;

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

    this.setState({ visibleItems: filtered });
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
      <div style={{ padding: "20px" }}>
        {/* **************** HEADER **************** */}
        <h2>Media Library</h2>
        <div
          style={{
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
          }}
        >
          {!this.state.selectedBucket && (
            <div style={{ marginBottom: "10px", display: "flex", gap: "10px" }}>
              <button
                onClick={() => this.setState({ resultMode: "folders" })}
                style={{
                  padding: "6px 10px",
                  background:
                    this.state.resultMode === "folders" ? "#0078d4" : "#ddd",
                  color:
                    this.state.resultMode === "folders" ? "white" : "black",
                  border: "none",
                  borderRadius: "6px",
                  cursor: "pointer",
                }}
              >
                Ordner
              </button>

              <button
                onClick={() => this.setState({ resultMode: "files" })}
                style={{
                  padding: "6px 10px",
                  background:
                    this.state.resultMode === "files" ? "#0078d4" : "#ddd",
                  color: this.state.resultMode === "files" ? "white" : "black",
                  border: "none",
                  borderRadius: "6px",
                  cursor: "pointer",
                }}
              >
                Dateien
              </button>
            </div>
          )}

          <button
            onClick={() => this.setState({ isUploadOpen: true })}
            style={{
              padding: "10px 16px",
              fontSize: "20px",
              background: "#0078d4",
              color: "white",
              border: "none",
              borderRadius: "8px",
              cursor: "pointer",
              marginBottom: "10px",
              alignSelf: "end",
            }}
          >
            ＋
          </button>
        </div>

        {/* **************** SEARCH **************** */}
        <div style={{ display: "flex", flexDirection: "column" }}>
          <input
            type="text"
            placeholder={
              this.state.selectedBucket
                ? `Suche in "${this.state.selectedBucket}"`
                : "Suche..."
            }
            value={this.state.searchText}
            onChange={this.onSearchChange}
            style={{ padding: "8px", width: "100%", maxWidth: "300px" }}
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
            style={{
              marginTop: "10px",
              padding: "6px 10px",
              cursor: "pointer",
            }}
          >
            ← Zurück zu Ordnern
          </button>
        )}

        <div style={{ display: "flex", gap: "5px", marginTop: "10px" }}>
          <div style={{ marginTop: "10px", display: "flex", gap: "8px" }}>
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
          <div style={{ marginTop: "10px", display: "flex", gap: "8px" }}>
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
          <div
            style={{
              display: "flex",
              flexWrap: "wrap",
              gap: "20px",
              marginTop: "20px",
            }}
          >
            {this.getFilteredBuckets().map((bucket) => {
              const preview = this.getBucketPreview(bucket);

              const fileType = preview?.name?.split(".").pop()?.toLowerCase();
              const isVideo = fileType === "mp4" || fileType === "mov";

              const imageUrl = preview
                ? `${window.location.origin}/_layouts/15/getpreview.ashx?path=${encodeURIComponent(preview.fileRef)}`
                : "";

              const videoUrl = preview
                ? `${window.location.origin}${preview.fileRef}`
                : "";

              return (
                <div
                  key={bucket}
                  style={{
                    position: "relative", // ✅ nötig für absolutes + Icon
                    width: "250px",
                    border: "1px solid #ddd",
                    borderRadius: "12px",
                    overflow: "hidden",
                    cursor: "pointer",
                    boxShadow: "0 4px 10px rgba(0,0,0,0.1)",
                  }}
                  onClick={() =>
                    this.setState(
                      {
                        viewMode: "items",
                        resultMode: "files",
                        selectedBucket: bucket,
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
                  <button
                    style={{
                      position: "absolute",
                      bottom: "8px",
                      right: "8px",
                      width: "32px",
                      height: "32px",
                      borderRadius: "50%",
                      background: "#0078d4",
                      color: "white",
                      border: "none",
                      fontSize: "18px",
                      fontWeight: "bold",
                      cursor: "pointer",
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      boxShadow: "0 2px 6px rgba(0,0,0,0.3)",
                    }}
                    onClick={(e) => {
                      e.stopPropagation();
                      e.preventDefault(); // ✅ WICHTIG!

                      this.setState({
                        isUploadOpen: true,
                        uploadBucket: [bucket],
                      });
                    }}
                  >
                    <span style={{ transform: "translateY(-2px)" }}>+</span>
                  </button>

                  {preview && (
                    <img
                      src={!isVideo ? imageUrl : ""}
                      ref={(el) => {
                        if (!el || !preview) return;

                        if (isVideo) {
                          const video = document.createElement("video");

                          video.src = videoUrl; // ✅ WICHTIG FIX
                          video.crossOrigin = "anonymous";
                          video.currentTime = 2; // ✅ besser als 100

                          video.onloadeddata = () => {
                            const canvas = document.createElement("canvas");

                            canvas.width = 250;
                            canvas.height = 150;

                            const ctx = canvas.getContext("2d");

                            if (ctx) {
                              const videoRatio =
                                video.videoWidth / video.videoHeight;
                              const canvasRatio = canvas.width / canvas.height;

                              let drawWidth, drawHeight, offsetX, offsetY;

                              if (videoRatio > canvasRatio) {
                                drawHeight = canvas.height;
                                drawWidth = canvas.height * videoRatio;
                                offsetX = (canvas.width - drawWidth) / 2;
                                offsetY = 0;
                              } else {
                                drawWidth = canvas.width;
                                drawHeight = canvas.width / videoRatio;
                                offsetX = 0;
                                offsetY = (canvas.height - drawHeight) / 2;
                              }

                              ctx.drawImage(
                                video,
                                offsetX,
                                offsetY,
                                drawWidth,
                                drawHeight,
                              );

                              el.src = canvas.toDataURL();
                            }
                          };
                        }
                      }}
                      style={{
                        width: "100%",
                        height: "150px",
                        objectFit: "cover",
                      }}
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

                  <div style={{ padding: "10px" }}>
                    <h3>{bucket}</h3>
                    <p style={{ fontSize: "12px", color: "#666" }}>
                      {bucketCounts[bucket] || 0} Dateien
                    </p>
                  </div>
                </div>
              );
            })}
          </div>
        ) : (
          <div
            style={{
              display: "flex",
              flexWrap: "wrap",
              gap: "20px",
              marginTop: "20px",
            }}
          >
            {/* **************** MEDIA GRID **************** */}

            {this.state.visibleItems.map((item) => {
              const downloadUrl = `${window.location.origin}${item.fileRef}`;

              let thumbnailUrl = `${window.location.origin}/_layouts/15/getpreview.ashx?path=${encodeURIComponent(item.fileRef)}&resolution=1`;

              const fileType = item.name?.split(".").pop()?.toLowerCase();
              const isVideo = fileType === "mp4" || fileType === "mov";

              // VIDEO FALL → später via canvas erzeugen
              if (isVideo) {
                thumbnailUrl = ""; // erstmal leer
              }
              // VIDEO FALL → später via canvas erzeugen
              if (isVideo) {
                thumbnailUrl = ""; // erstmal leer
              }
              return (
                <div
                  key={item.id}
                  style={{
                    width: "300px",
                    border: "1px solid #ddd",
                    borderRadius: "12px",
                    overflow: "hidden",
                    boxShadow: "0 4px 10px rgba(0,0,0,0.1)",
                  }}
                >
                  {isVideo ? (
                    <div
                      onClick={() =>
                        this.setState({
                          isModalOpen: true,
                          selectedItem: item,
                        })
                      }
                      style={{
                        position: "relative",
                        cursor: "pointer",
                      }}
                    >
                      <img
                        src={thumbnailUrl}
                        ref={(el) => {
                          if (el && isVideo && !thumbnailUrl) {
                            const video = document.createElement("video");
                            video.src = `${window.location.origin}${item.fileRef}`;
                            video.crossOrigin = "anonymous";
                            video.currentTime = 100;

                            video.onloadeddata = () => {
                              const canvas = document.createElement("canvas");

                              // Zielgröße (wie deine Card)
                              canvas.width = 300;
                              canvas.height = 200;

                              const ctx = canvas.getContext("2d");

                              if (ctx) {
                                const videoRatio =
                                  video.videoWidth / video.videoHeight;
                                const canvasRatio =
                                  canvas.width / canvas.height;

                                let drawWidth, drawHeight, offsetX, offsetY;

                                if (videoRatio > canvasRatio) {
                                  // Video breiter → links/rechts abschneiden
                                  drawHeight = canvas.height;
                                  drawWidth = canvas.height * videoRatio;
                                  offsetX = (canvas.width - drawWidth) / 2;
                                  offsetY = 0;
                                } else {
                                  // Video höher → oben/unten abschneiden (wichtig für Hochformat!)
                                  drawWidth = canvas.width;
                                  drawHeight = canvas.width / videoRatio;
                                  offsetX = 0;
                                  offsetY = (canvas.height - drawHeight) / 2;
                                }

                                ctx.drawImage(
                                  video,
                                  offsetX,
                                  offsetY,
                                  drawWidth,
                                  drawHeight,
                                );
                                el.src = canvas.toDataURL();
                              }
                            };
                          }
                        }}
                        onError={(e) => {
                          e.currentTarget.src =
                            "data:image/svg+xml;charset=UTF-8," +
                            encodeURIComponent(`
        <svg width="300" height="200" xmlns="http://www.w3.org/2000/svg">
          <rect width="100%" height="100%" fill="#f3f2f1"/>
          <text x="50%" y="50%" dominant-baseline="middle" text-anchor="middle" fill="#605e5c" font-size="16">
            Kein Preview
          </text>
        </svg>
      `);
                        }}
                        style={{
                          width: "100%",
                          height: "200px",
                          objectFit: "cover",
                        }}
                      />

                      <div
                        style={{
                          position: "absolute",
                          inset: 0,
                          display: "flex",
                          justifyContent: "center",
                          alignItems: "center",
                          fontSize: "40px",
                          color: "white",
                          background: "rgba(0,0,0,0.4)",
                        }}
                      >
                        ▶
                      </div>
                    </div>
                  ) : (
                    <img
                      src={thumbnailUrl}
                      style={{ width: "100%", cursor: "pointer" }}
                      onClick={() =>
                        this.setState({
                          isModalOpen: true,
                          selectedItem: item,
                        })
                      }
                      onError={(e) => {
                        e.currentTarget.src =
                          "data:image/svg+xml;charset=UTF-8," +
                          encodeURIComponent(`
        <svg width="300" height="200" xmlns="http://www.w3.org/2000/svg">
          <rect width="100%" height="100%" fill="#f3f2f1"/>
          <text x="50%" y="50%" dominant-baseline="middle" text-anchor="middle" fill="#605e5c" font-size="16">
            Kein Preview
          </text>
        </svg>
      `);
                      }}
                    />
                  )}

                  <div style={{ padding: "12px" }}>
                    <h3>{item.name}</h3>
                    {/* ✅ TAG CHIPS HIER */}
                    <div
                      style={{
                        display: "flex",
                        flexWrap: "wrap",
                        gap: "6px",
                        marginBottom: "8px",
                      }}
                    >
                      {(item.tags || []).map((tag, i) => (
                        <span
                          key={i}
                          onClick={() =>
                            this.setState(
                              { searchText: tag.toLowerCase() },
                              this.applyFilters,
                            )
                          }
                          style={{
                            background: "#0078d4",
                            color: "white",
                            padding: "4px 8px",
                            borderRadius: "12px",
                            fontSize: "11px",
                            cursor: "pointer",
                          }}
                        >
                          {tag}
                        </span>
                      ))}
                    </div>
                    <p style={{ fontSize: "12px", color: "#666" }}>
                      Erstellt am:{" "}
                      {item.created
                        ? new Date(item.created).toLocaleDateString()
                        : "-"}
                    </p>
                    <div
                      style={{ display: "flex", gap: "8px", marginTop: "10px" }}
                    >
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
                        style={{
                          flex: 1,
                          padding: "10px",
                          background: "#e42828", // ✅ ROT bleibt!
                          color: "white",
                          borderRadius: "6px",
                          border: "none",
                          cursor: "pointer",
                          fontWeight: "bold",
                        }}
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
                        style={{
                          flex: 1,
                          padding: "10px",
                          background: "#e1dfdd", // leicht grau
                          color: "#323130",
                          borderRadius: "6px",
                          border: "1px solid #c8c6c4",
                          cursor: "pointer",
                          fontWeight: "bold",
                        }}
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

        {/* **************** PREVIEW MODAL **************** */}

        {this.state.isModalOpen && this.state.selectedItem && (
          <div
            style={{
              position: "fixed",
              top: 0,
              left: 0,
              width: "100vw",
              height: "100vh",
              background: "rgba(0,0,0,0.9)",
              display: "flex",
              justifyContent: "center",
              alignItems: "center",
              zIndex: 9999,
            }}
          >
            {/* CLOSE BUTTON */}
            <button
              onClick={() =>
                this.setState({ isModalOpen: false, selectedItem: undefined })
              }
              style={{
                position: "absolute",
                top: "20px",
                right: "20px",
                fontSize: "24px",
                background: "transparent",
                color: "white",
                border: "none",
                cursor: "pointer",
              }}
            >
              ✕
            </button>
            {/* CONTENT MODAL */}
            <div
              style={{
                maxWidth: "90%",
                maxHeight: "90%",
                display: "flex",
                flexDirection: "column",
                alignItems: "center",
                gap: "10px",
              }}
            >
              {(() => {
                const item = this.state.selectedItem;
                const fileType = item?.name.split(".").pop()?.toLowerCase();
                const isVideo = fileType === "mp4" || fileType === "mov";

                const fileUrl = `${window.location.origin}${item?.fileRef}`;
                const downloadUrl = `${window.location.origin}/_layouts/15/download.aspx?SourceUrl=${encodeURIComponent(fileUrl)}`;

                return (
                  <>
                    {/* IMAGE */}
                    {!isVideo && (
                      <img
                        src={fileUrl}
                        style={{
                          maxWidth: "100%",
                          maxHeight: "80vh",
                          objectFit: "contain",
                        }}
                      />
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

                    {/* DOWNLOAD BUTTON */}
                    <button
                      onClick={() => {
                        const link = document.createElement("a");
                        link.href = downloadUrl;
                        link.setAttribute("download", item?.name || "file");

                        document.body.appendChild(link);
                        link.click();
                        document.body.removeChild(link);
                      }}
                      style={{
                        padding: "12px 20px",
                        background: "#e42828",
                        color: "white",
                        border: "none",
                        borderRadius: "6px",
                        cursor: "pointer",
                        fontWeight: "bold",
                      }}
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
          <div
            style={{
              position: "fixed",
              top: 0,
              left: 0,
              width: "100vw",
              height: "100vh",
              background: "rgba(0,0,0,0.7)",
              display: "flex",
              justifyContent: "center",
              alignItems: "center",
              zIndex: 9999,
            }}
          >
            <div
              style={{
                background: "white",
                padding: "20px",
                borderRadius: "12px",
                width: "400px",
                display: "flex",
                flexDirection: "column",
                gap: "10px",
              }}
            >
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
                  return (
                    <img
                      src={fileUrl}
                      style={{
                        width: "100%",
                        maxHeight: "200px",
                        objectFit: "cover",
                        borderRadius: "8px",
                      }}
                    />
                  );
                }

                if (isVideo) {
                  return (
                    <video
                      src={fileUrl}
                      controls
                      style={{
                        width: "100%",
                        maxHeight: "200px",
                        borderRadius: "8px",
                      }}
                    />
                  );
                }

                return (
                  <div
                    style={{
                      padding: "20px",
                      background: "#f3f2f1",
                      borderRadius: "8px",
                      textAlign: "center",
                    }}
                  >
                    📄 {item.name}
                  </div>
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
                <div style={{ display: "flex", flexWrap: "wrap", gap: "6px" }}>
                  {this.state.editTags.map((tag, index) => (
                    <span
                      key={index}
                      style={{
                        background: "#0078d4",
                        color: "white",
                        padding: "4px 8px",
                        borderRadius: "12px",
                        fontSize: "12px",
                        display: "flex",
                        alignItems: "center",
                        gap: "6px",
                      }}
                    >
                      {tag}
                      <span
                        onClick={() => {
                          const newTags = [...this.state.editTags];
                          newTags.splice(index, 1);
                          this.setState({ editTags: newTags });
                        }}
                        style={{
                          cursor: "pointer",
                          fontWeight: "bold",
                        }}
                      >
                        ✕
                      </span>
                    </span>
                  ))}
                </div>

                <input
                  type="text"
                  style={{
                    width: "100%",
                    marginTop: "10px",
                  }}
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
                <option value="Säugetier">Säugetier</option>
                <option value="Vogel">Vogel</option>
              </select>
              <select
                value={this.state.editFormat}
                onChange={(e) => this.setState({ editFormat: e.target.value })}
              >
                <option value="">Format wählen</option>
                <option value="Bild">Bild</option>
                <option value="Video">Video</option>
                <option value="Dokument">Dokument</option>
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

              <div style={{ display: "flex", gap: "10px", marginTop: "10px" }}>
                <button
                  onClick={() => this.setState({ isEditOpen: false })}
                  style={{ flex: 1 }}
                >
                  Abbrechen
                </button>

                <button onClick={() => this.updateItem()} style={{ flex: 1 }}>
                  Speichern
                </button>

                <button
                  onClick={() => this.deleteItem()}
                  style={{
                    flex: 1,
                    background: "#d13438",
                    color: "white",
                    border: "none",
                    borderRadius: "6px",
                    cursor: "pointer",
                    fontWeight: "bold",
                  }}
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
            style={{
              position: "fixed",
              top: 0,
              left: 0,
              width: "100vw",
              height: "100vh",
              background: "rgba(0,0,0,0.7)",
              display: "flex",
              justifyContent: "center",
              alignItems: "center",
              zIndex: 9999,
              overflowY: "auto",
              padding: "20px",
            }}
          >
            <div
              style={{
                background: "white",
                padding: "20px",
                borderRadius: "12px",
                width: "400px",
                display: "flex",
                flexDirection: "column",
                gap: "10px",
                maxHeight: "90vh",
                overflowY: "auto",
              }}
            >
              <h3>Upload</h3>
              {/* MODAL PREVIEW */}
              {this.state.uploadFiles?.[0] && this.state.uploadPreviewUrl && (
                <div style={{ marginTop: "10px" }}>
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
                          style={{
                            width: "100%",
                            maxHeight: "200px",
                            objectFit: "cover",
                            borderRadius: "8px",
                          }}
                        />
                      );
                    }

                    if (isVideo) {
                      return (
                        <video
                          src={this.state.uploadPreviewUrl}
                          controls
                          style={{
                            width: "100%",
                            maxHeight: "200px",
                            borderRadius: "8px",
                          }}
                        />
                      );
                    }

                    return (
                      <div
                        style={{
                          padding: "20px",
                          background: "#f3f2f1",
                          borderRadius: "8px",
                          textAlign: "center",
                        }}
                      >
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
              <BucketDropdown
                options={this.state.bucketOptions}
                selected={this.state.uploadBucket}
                onChange={(values) =>
                  this.setState({
                    uploadBucket: values,
                  })
                }
              />
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

              <select
                value={this.state.uploadCategory}
                onChange={(e) =>
                  this.setState({ uploadCategory: e.target.value })
                }
              >
                <option value="">Kategorie wählen</option>

                {this.getUniqueCategories().map((cat) => (
                  <option key={cat} value={cat}>
                    {cat}
                  </option>
                ))}
              </select>

              <div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: "6px" }}>
                  {this.state.uploadTags.map((tag, index) => (
                    <span
                      key={index}
                      style={{
                        background: "#0078d4",
                        color: "white",
                        padding: "4px 8px",
                        borderRadius: "12px",
                        fontSize: "12px",
                        display: "flex",
                        alignItems: "center",
                        gap: "6px",
                      }}
                    >
                      {tag}
                      <span
                        onClick={() => {
                          const newTags = [...this.state.uploadTags];
                          newTags.splice(index, 1);
                          this.setState({ uploadTags: newTags });
                        }}
                        style={{ cursor: "pointer", fontWeight: "bold" }}
                      >
                        ✕
                      </span>
                    </span>
                  ))}
                </div>

                <input
                  type="text"
                  placeholder="Tag hinzufügen + Enter"
                  style={{ width: "100%", marginTop: "10px" }}
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
              <button
                onClick={() => this.uploadItem()}
                disabled={
                  this.state.isUploading ||
                  !this.state.uploadBucket ||
                  this.state.uploadBucket.length === 0
                }
                style={{
                  opacity:
                    this.state.isUploading ||
                    !this.state.uploadBucket ||
                    this.state.uploadBucket.length === 0
                      ? 0.5
                      : 1,
                  cursor:
                    this.state.isUploading ||
                    !this.state.uploadBucket ||
                    this.state.uploadBucket.length === 0
                      ? "not-allowed"
                      : "pointer",
                }}
              >
                {this.state.isUploading
                  ? "⏳ Wird hochgeladen..."
                  : "Hochladen"}
              </button>
            </div>
          </div>
        )}
      </div>
    );
  }
}
