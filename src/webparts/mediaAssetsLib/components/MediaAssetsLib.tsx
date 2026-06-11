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
}

interface IMediaAssetsLibState {
  allItems: IMediaItem[];
  visibleItems: IMediaItem[];
  searchText: string;
  filterCategory?: string;
  filterYear?: number;
  filterMonth?: number;
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
    };
  }

  public async componentDidMount(): Promise<void> {
    await this.loadAllMedia();
  }

  private async getFolderContent(folderUrl: string): Promise<IMediaItem[]> {
    const url = `${this.props.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')?$expand=Folders,Files/ListItemAllFields&$select=Name,ServerRelativeUrl,TimeCreated,ListItemAllFields/Id,ListItemAllFields/Kategorie,ListItemAllFields/Notizen,ListItemAllFields/UniqueId,ListItemAllFields/Tags`;

    const response = await this.props.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
    );

    const data = await response.json();

    let results: IMediaItem[] = [];

    data.Files.forEach((file: any) => {
      /*console.log("FIELDS:", file.ListItemAllFields);*/
      results.push({
        id: file.ListItemAllFields.Id,
        name: file.Name,
        fileRef: file.ServerRelativeUrl,
        category: file.ListItemAllFields?.Kategorie,
        notes: file.ListItemAllFields?.Notizen,
        created: file.TimeCreated,
        tags: file.ListItemAllFields?.Tags || [],
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

      this.setState({
        allItems: items,
        visibleItems: items,
      });
    } catch (error) {
      console.error(error);
    }
  }

  private applyFilters = (): void => {
    console.log("FILTER TRIGGERED");
    console.log("Year:", this.state.filterYear);
    console.log("Month:", this.state.filterMonth);

    const { allItems, searchText, filterCategory } = this.state;

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

  public render(): React.ReactElement<IMediaAssetsLibProps> {
    const categoryOptions = this.getUniqueCategories();

    const yearOptions = this.getUniqueYears();

    return (
      <div style={{ padding: "20px" }}>
        <h2>Media Library</h2>

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
            const videoUrl = `${window.location.origin}${item.fileRef}?web=1`;
            const downloadUrl = `${window.location.origin}/_layouts/15/download.aspx?SourceUrl=${encodeURIComponent(
              `${window.location.origin}${item.fileRef}`,
            )}`;
            const fileType = item.name.split(".").pop()?.toLowerCase();

            const isVideo = fileType === "mp4" || fileType === "mov";

            const thumbnailUrl = isVideo
              ? undefined
              : `${window.location.origin}/_layouts/15/getpreview.ashx?path=${encodeURIComponent(item.fileRef)}&resolution=1`;

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
                    onClick={() => window.open(videoUrl, "_blank")}
                    style={{
                      position: "relative",
                      cursor: "pointer",
                    }}
                  >
                    <img
                      src={thumbnailUrl}
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
                      /* IMAGE GRÖSSE CARD*/
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
                    style={{ width: "100%" }}
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
                      display: "block",
                      marginTop: "10px",
                      padding: "10px",
                      background: "#e42828",
                      color: "white",
                      textAlign: "center",
                      borderRadius: "6px",
                      border: "none",
                      cursor: "pointer",
                      fontWeight: "bold",
                    }}
                  >
                    Download
                  </button>
                </div>
              </div>
            );
          })}
        </div>
      </div>
    );
  }
}
