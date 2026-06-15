import * as React from "react";
import { SPHttpClient } from "@microsoft/sp-http";
import type { IMediaAssetsLibProps } from "./IMediaAssetsLibProps";

interface IMediaItem {
  id: number;
  name: string;
  fileRef: string;
  category?: string;
  notes?: string;
  created?: string;
  uniqueId?: string;
  tags?: string[];

  driveId?: string;
  driveItemId?: string;

  format?: string;
}

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

  isUploadOpen: boolean; // ✅ NEU
  uploadName: string;
  uploadTags: string[];
  uploadCategory: string;
  uploadFormat: string;
  uploadBucket: string;
  uploadFile?: File;
  uploadPreviewUrl?: string;

  bucketOptions: string[];

  uploadProgress: number;
  uploading: boolean;
}

export default class MediaAssetsLib extends React.Component<
  IMediaAssetsLibProps,
  IMediaAssetsLibState
> {
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

      isUploadOpen: false,
      uploadName: "",
      uploadTags: [],
      uploadCategory: "",
      uploadFormat: "",
      uploadBucket: "",
      uploadFile: undefined,

      bucketOptions: [],

      uploadProgress: 0,
      uploading: false,
    };
  }

  public async componentDidMount(): Promise<void> {
    await this.loadAllMedia();
    await this.loadBuckets(); // ✅ NEU
  }

  private async getFolderContent(folderUrl: string): Promise<IMediaItem[]> {
    /* const url = `${this.props.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')?$expand=Folders,Files/ListItemAllFields&$select=Name,ServerRelativeUrl,TimeCreated,ListItemAllFields/Id,ListItemAllFields/Kategorie,ListItemAllFields/Notizen,ListItemAllFields/UniqueId,ListItemAllFields/Tags`;*/

    const url = `${this.props.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')?$expand=Folders,Files,Files/ListItemAllFields&$select=Name,ServerRelativeUrl,TimeCreated,Files/Name,Files/ServerRelativeUrl,Files/TimeCreated,Files/ListItemAllFields/Id,Files/ListItemAllFields/Kategorie,Files/ListItemAllFields/Notizen,Files/ListItemAllFields/UniqueId,Files/ListItemAllFields/Tags,Files/ListItemAllFields/Format,Files/ListItemAllFields/Bucket`;
    const response = await this.props.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
    );

    const data = await response.json();

    let results: IMediaItem[] = [];

    data.Files.forEach((file: any) => {
      /*console.log("FIELDS:", file.ListItemAllFields);*/
      console.log("FIELDS:", file.ListItemAllFields);

      results.push({
        id: file.ListItemAllFields.Id,
        name: file.Name,
        fileRef: file.ServerRelativeUrl,
        category: file.ListItemAllFields?.Kategorie,
        notes: file.ListItemAllFields?.Notizen,
        created: file.TimeCreated,
        tags: Array.isArray(file.ListItemAllFields?.Tags)
          ? file.ListItemAllFields.Tags
          : file.ListItemAllFields?.Tags
            ? String(file.ListItemAllFields.Tags)
                .split(/[;,#]+/)
                .map((t: string) => t.trim())
                .filter((t: string) => t)
            : [],

        driveId: file.ListItemAllFields?.File?.VroomDriveID,
        driveItemId: file.ListItemAllFields?.File?.VroomItemID,
        format: file.ListItemAllFields?.Format,
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

  private async loadAllMedia(): Promise<void> {
    try {
      const rootFolder = "/sites/IntranetSpielwiese/Medienbibliothek";

      const items = await this.getFolderContent(rootFolder);

      this.setState(
        {
          allItems: items,
        },
        this.applyFilters,
      );
    } catch (error) {
      console.error(error);
    }
  }

  private async updateItem(): Promise<void> {
    const { selectedItem, editName, editTags, editCategory, editFormat } =
      this.state;

    if (!selectedItem) return;

    try {
      /* BIBLIOTHEK ÄNDERN */
      const url = `${this.props.siteUrl}/_api/web/lists/getbytitle('Medienbibliothek')/items(${selectedItem.id})`;

      const tagsArray = editTags;

      await this.props.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata",
          "IF-MATCH": "*",
          "X-HTTP-Method": "MERGE",
        },
        body: JSON.stringify({
          FileLeafRef: editName, // Dateiname-Feld (wichtig!)
          Kategorie: editCategory, // deine SP-Spalte
          Tags: {
            results: tagsArray || [],
          },
          Format: editFormat,
        }),
      });

      console.log("Update erfolgreich");

      // neu laden → damit UI sofort aktualisiert wird
      await this.loadAllMedia();

      this.setState({ isEditOpen: false });
    } catch (error) {
      console.error("Update Fehler:", error);
    }
  }

  private async uploadFileToSP(): Promise<void> {
    this.setState({
      uploading: true,
      uploadProgress: 20,
    });

    const {
      uploadFile,
      uploadName,
      uploadCategory,
      uploadFormat,
      uploadTags,
      uploadBucket,
    } = this.state;

    if (!uploadFile) {
      alert("Bitte Datei auswählen");

      this.setState({
        uploading: false,
        uploadProgress: 0,
      });

      return;
    }

    try {
      const folderUrl = "/sites/IntranetSpielwiese/Medienbibliothek";
      const extension = uploadFile.name.split(".").pop();
      const fileName = `${uploadName}.${extension}`;

      // ✅ 1. Datei hochladen
      const uploadUrl = `${this.props.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')/Files/add(overwrite=true,url='${fileName}')?$expand=ListItemAllFields`;
      const uploadResponse = await this.props.spHttpClient.post(
        uploadUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
          },
          body: uploadFile,
        },
      );

      const fileData = await uploadResponse.json();

      this.setState({ uploadProgress: 60 });

      const itemId = fileData?.ListItemAllFields?.Id;

      if (!itemId) {
        console.error("Kein ListItem gefunden:", fileData);
        throw new Error("Upload fehlgeschlagen: Kein ItemId");
      }
      await this.ensureBucketExists(uploadBucket);
      await this.loadBuckets(); // ✅ WICHTIG

      // ✅ 2. Metadaten setzen
      const updateUrl = `${this.props.siteUrl}/_api/web/lists/getbytitle('Medienbibliothek')/items(${itemId})`;
      // ✅ DEBUG – zeigt dir genau was an SharePoint geht
      console.log("METADATA:", {
        FileLeafRef: fileName,
        Kategorie: uploadCategory,
        Format: uploadFormat,
        Tags: { results: uploadTags || [] },
        Bucket: uploadBucket,
      });
      const res = await this.props.spHttpClient.post(
        updateUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE",
          },
          body: JSON.stringify({
            FileLeafRef: fileName, // optional später: uploadName einsetzen
            Kategorie: uploadCategory,
            Format: uploadFormat,

            Tags: {
              results: uploadTags || [],
            },

            Bucket: uploadBucket || null,
          }),
        },
      );
      console.log("STATUS:", res.status);
      this.setState({ uploadProgress: 100 });
      const checkUrl = `${this.props.siteUrl}/_api/web/lists/getbytitle('Medienbibliothek')/items(${itemId})?$select=Tags`;

      const checkRes = await this.props.spHttpClient.get(
        checkUrl,
        SPHttpClient.configurations.v1,
      );

      const checkData = await checkRes.json();

      console.log("GESPEICHERTE TAGS:", checkData);
      console.log("Upload Response:", fileData);

      console.log("Upload erfolgreich ✅");

      // ✅ UI reset
      this.setState({
        isUploadOpen: false,
        uploadFile: undefined,
        uploadName: "",
        uploadTags: [],
        uploadCategory: "",
        uploadFormat: "",
        uploadBucket: "",
        uploadPreviewUrl: undefined,
        uploading: false, // ✅ FEHLT
        uploadProgress: 0, // ✅ FEHLT
      });

      // ✅ neu laden
      await this.loadAllMedia();
    } catch (error) {
      console.error("UPLOAD ERROR:", error);

      this.setState({
        uploading: false,
        uploadProgress: 0,
      });
    }
  }

  private applyFilters = (): void => {
    /*console.log("FILTER TRIGGERED");
    console.log("Year:", this.state.filterYear);
    console.log("Month:", this.state.filterMonth);*/

    const { allItems, searchText, filterCategory, filterFormat } = this.state;

    let filtered = [...allItems];
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

  private onSearchChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    this.setState({ searchText: e.target.value }, this.applyFilters);
  };

  private getUniqueCategories(): string[] {
    const values = this.state.allItems
      .map((item) => item.category)
      .filter((v) => v);

    return Array.from(new Set(values)) as string[];
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

  private async ensureBucketExists(value: string): Promise<void> {
    if (!value) return;

    if (this.state.bucketOptions.includes(value)) return;

    try {
      const newChoices = [...this.state.bucketOptions, value];

      const url = `${this.props.siteUrl}/_api/web/lists/getbytitle('Medienbibliothek')/fields/getbyinternalnameortitle('Bucket')`;

      await this.props.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: {
          Accept: "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=nometadata",
          "IF-MATCH": "*",
          "X-HTTP-Method": "MERGE",
        },
        body: JSON.stringify({
          Choices: newChoices,
        }),
      });

      console.log("Neuer Bucket hinzugefügt ✅");
    } catch (error) {
      console.error("Bucket erstellen Fehler:", error);
    }
  }

  private async loadBuckets(): Promise<void> {
    try {
      const url = `${this.props.siteUrl}/_api/web/lists/getbytitle('Medienbibliothek')/fields/getbyinternalnameortitle('Bucket')`;

      const response = await this.props.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
      );

      const data = await response.json();

      const choices = data.Choices || [];

      // ✅ Neueste zuerst (einfach reverse – SharePoint sortiert alt→neu)
      const sorted = choices.reverse();

      this.setState({
        bucketOptions: sorted, // wir brauchen gleich ein neues State Feld
      });
    } catch (error) {
      console.error("Bucket laden Fehler:", error);
    }
  }

  public render(): React.ReactElement<IMediaAssetsLibProps> {
    const categoryOptions = this.getUniqueCategories();

    const yearOptions = this.getUniqueYears();

    /* HEADER */
    return (
      <div style={{ padding: "20px" }}>
        <div
          style={{
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
          }}
        >
          <h2>Media Library</h2>

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
            }}
          >
            ＋
          </button>
        </div>
        <input
          type="text"
          placeholder="Suche..."
          value={this.state.searchText}
          onChange={this.onSearchChange}
          style={{ padding: "8px", width: "300px" }}
        />
        <div style={{ marginTop: "10px" }}>
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
        </div>
        <div style={{ marginTop: "10px", display: "flex", gap: "10px" }}>
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
        <p>Ergebnisse: {this.state.visibleItems.length}</p>
        <div
          style={{
            display: "flex",
            flexWrap: "wrap",
            gap: "20px",
            marginTop: "20px",
          }}
        >
          {this.state.visibleItems.map((item) => {
            /*const videoUrl = `${window.location.origin}${item.fileRef}?web=1`;*/

            const downloadUrl = `${window.location.origin}/_layouts/15/download.aspx?SourceUrl=${encodeURIComponent(
              `${window.location.origin}${item.fileRef}`,
            )}`;
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
                          video.currentTime = 1;

                          video.onloadeddata = () => {
                            const canvas = document.createElement("canvas");

                            // Zielgröße (wie deine Card)
                            canvas.width = 300;
                            canvas.height = 200;

                            const ctx = canvas.getContext("2d");

                            if (ctx) {
                              const videoRatio =
                                video.videoWidth / video.videoHeight;
                              const canvasRatio = canvas.width / canvas.height;

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
                        const link = document.createElement("a");
                        link.href = downloadUrl;
                        link.setAttribute("download", item.name);

                        document.body.appendChild(link);
                        link.click();
                        document.body.removeChild(link);
                      }}
                      style={{
                        flex: 1,
                        padding: "10px",
                        background: "#e42828",
                        color: "white",
                        borderRadius: "6px",
                        border: "none",
                        cursor: "pointer",
                        fontWeight: "bold",
                      }}
                    >
                      Download
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
        // EDIT MODAL //////////
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
                    <span key={index}>
                      {tag}
                      <span
                        onClick={() => {
                          const newTags = [...this.state.editTags];
                          newTags.splice(index, 1);
                          this.setState({ editTags: newTags });
                        }}
                      >
                        ✕
                      </span>
                    </span>
                  ))}
                </div>

                <input
                  type="text"
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

              <div style={{ display: "flex", gap: "10px", marginTop: "10px" }}>
                <button onClick={() => this.setState({ isEditOpen: false })}>
                  Abbrechen
                </button>

                <button onClick={() => this.updateItem()}>Speichern</button>
              </div>
            </div>
          </div>
        )}
        {/* UPLOAD MODAL */}
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
              <h3>Upload</h3>
              {/* UPLOAD BALKEN */}
              {this.state.uploading && (
                <div style={{ marginBottom: "10px" }}>
                  <div
                    style={{
                      height: "10px",
                      background: "#eee",
                      borderRadius: "5px",
                      overflow: "hidden",
                    }}
                  >
                    <div
                      style={{
                        width: `${this.state.uploadProgress}%`,
                        background: "#0078d4",
                        height: "100%",
                        transition: "width 0.3s",
                      }}
                    />
                  </div>

                  <p style={{ fontSize: "12px" }}>
                    Upload: {this.state.uploadProgress}%
                  </p>
                </div>
              )}

              {/* MODAL PREVIEW */}
              {this.state.uploadFile && this.state.uploadPreviewUrl && (
                <div style={{ marginTop: "10px" }}>
                  {(() => {
                    const file = this.state.uploadFile!;
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
                onChange={(e) => {
                  const file = e.target.files?.[0];
                  if (!file) return;

                  const fileName = file.name;
                  const baseName = fileName.includes(".")
                    ? fileName.substring(0, fileName.lastIndexOf("."))
                    : fileName;

                  const previewUrl = URL.createObjectURL(file); // ✅ NEU

                  this.setState({
                    uploadFile: file,
                    uploadName: baseName,
                    uploadPreviewUrl: previewUrl, // ✅ WICHTIG
                  });
                }}
              />
              <input
                type="text"
                placeholder="Name"
                value={this.state.uploadName}
                onChange={(e) => this.setState({ uploadName: e.target.value })}
              />
              {/* TAGS UPLOAD*/}
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
                  style={{ marginTop: "8px", width: "100%" }}
                />
              </div>
              <select
                value={this.state.uploadCategory}
                onChange={(e) =>
                  this.setState({ uploadCategory: e.target.value })
                }
              >
                <option value="">Kategorie wählen</option>
                <option value="Säugetier">Säugetier</option>
                <option value="Vogel">Vogel</option>
              </select>
              <select
                value={this.state.uploadFormat}
                onChange={(e) =>
                  this.setState({ uploadFormat: e.target.value })
                }
              >
                <option value="">Format wählen</option>
                <option value="Bild">Bild</option>
                <option value="Video">Video</option>
                <option value="Dokument">Dokument</option>
              </select>
              {/* BUCKET */}
              <input
                list="bucket-list"
                placeholder="Bucket wählen oder neu eingeben"
                value={this.state.uploadBucket}
                onChange={(e) =>
                  this.setState({ uploadBucket: e.target.value })
                }
              />
              <datalist id="bucket-list">
                {this.state.bucketOptions.map((b, i) => (
                  <option key={i} value={b} />
                ))}
              </datalist>
              <div style={{ display: "flex", gap: "10px", marginTop: "10px" }}>
                <button
                  onClick={() => this.setState({ isUploadOpen: false })}
                  style={{ flex: 1 }}
                >
                  Abbrechen
                </button>

                <button
                  onClick={() => this.uploadFileToSP()}
                  disabled={this.state.uploading}
                  style={{
                    flex: 1,
                    background: "#0078d4",
                    color: "white",
                    border: "none",
                    borderRadius: "6px",
                    padding: "10px",
                  }}
                >
                  Hochladen
                </button>
              </div>
            </div>
          </div>
        )}
      </div>
    );
  }
}
