import * as React from "react";
import { SPHttpClient } from "@microsoft/sp-http";
import type { IMediaAssetsLibProps } from "./IMediaAssetsLibProps";

interface IMediaItem {
  id: number;
  name: string;
  fileRef: string;
  tags?: string;
  category?: string;
  notes?: string;
  format?: string;
  published?: string;
  created?: string;
}

interface IMediaAssetsLibState {
  allItems: IMediaItem[];
  visibleItems: IMediaItem[];
  searchText: string;
  filterTag?: string;
  filterCategory?: string;
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
    };
  }

  public async componentDidMount(): Promise<void> {
    await this.loadAllMedia();
  }

  private async getFolderContent(folderUrl: string): Promise<IMediaItem[]> {
    const url = `${this.props.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')?$expand=Folders,Files/ListItemAllFields`;

    const response = await this.props.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
    );

    const data = await response.json();
    let results: IMediaItem[] = [];

    data.Files.forEach((file: any) => {
      results.push({
        id: file.ListItemAllFields.Id,
        name: file.Name,
        fileRef: file.ServerRelativeUrl,
        tags: file.ListItemAllFields?.Tags,
        category: file.ListItemAllFields?.Kategorie,
        notes: file.ListItemAllFields?.Notizen,
        format: file.ListItemAllFields?.Format,
        published: file.ListItemAllFields?.Veröffentlichungsdatum,
        created: file.TimeCreated,
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

      let items = await this.getFolderContent(rootFolder);

      items = items.sort((a, b) => {
        const d1 = new Date(a.published || a.created || 0).getTime();
        const d2 = new Date(b.published || b.created || 0).getTime();
        return d2 - d1;
      });

      this.setState({
        allItems: items,
        visibleItems: items,
      });
    } catch (error) {
      console.error(error);
    }
  }

  private applyFilters = (): void => {
    const { allItems, searchText, filterTag, filterCategory } = this.state;

    let filtered = [...allItems];
    const search = searchText.toLowerCase();

    if (search) {
      filtered = filtered.filter(
        (item) =>
          item.name.toLowerCase().includes(search) ||
          (item.tags && item.tags.toLowerCase().includes(search)) ||
          (item.notes && item.notes.toLowerCase().includes(search)),
      );
    }

    if (filterTag) {
      filtered = filtered.filter(
        (item) => item.tags && item.tags.includes(filterTag),
      );
    }

    if (filterCategory) {
      filtered = filtered.filter((item) => item.category === filterCategory);
    }

    this.setState({ visibleItems: filtered });
  };

  private onSearchChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    this.setState({ searchText: e.target.value }, this.applyFilters);
  };

  private getUniqueValues(field: "tags" | "category"): string[] {
    const values = this.state.allItems
      .map((item) => item[field])
      .filter((v) => v);

    return Array.from(new Set(values)) as string[];
  }

  public render(): React.ReactElement<IMediaAssetsLibProps> {
    const tagOptions = this.getUniqueValues("tags");
    const categoryOptions = this.getUniqueValues("category");

    return (
      <div style={{ padding: "20px" }}>
        <h2>Media Library</h2>

        {/* 🔍 Suche */}
        <input
          type="text"
          placeholder="Suche..."
          value={this.state.searchText}
          onChange={this.onSearchChange}
          style={{ padding: "8px", width: "300px" }}
        />

        {/* 🎯 Filter */}
        <div style={{ marginTop: "10px", display: "flex", gap: "10px" }}>
          <select
            onChange={(e) =>
              this.setState({ filterTag: e.target.value }, this.applyFilters)
            }
          >
            <option value="">Tag</option>
            {tagOptions.map((tag) => (
              <option key={tag}>{tag}</option>
            ))}
          </select>

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

        {/* 📦 CARDS */}
        <div
          style={{
            display: "flex",
            flexWrap: "wrap",
            gap: "20px",
            marginTop: "20px",
          }}
        >
          {this.state.visibleItems.map((item) => {
            const previewUrl = `${this.props.siteUrl}/_layouts/15/getpreview.ashx?path=${encodeURIComponent(
              item.fileRef,
            )}&resolution=1`;

            const downloadUrl = `${this.props.siteUrl}${item.fileRef}`;

            const fileType = item.name.split(".").pop()?.toLowerCase();
            const isVideo = fileType === "mp4" || fileType === "mov";

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
                {/* ✅ Thumbnail */}
                <div style={{ position: "relative" }}>
                  <img
                    src={previewUrl}
                    style={{
                      width: "100%",
                      height: "180px",
                      objectFit: "cover",
                    }}
                  />

                  {/* ▶ Video Overlay */}
                  {isVideo && (
                    <div
                      style={{
                        position: "absolute",
                        top: 0,
                        left: 0,
                        width: "100%",
                        height: "100%",
                        display: "flex",
                        justifyContent: "center",
                        alignItems: "center",
                        fontSize: "40px",
                        color: "white",
                        background: "rgba(0,0,0,0.3)",
                      }}
                    >
                      ▶
                    </div>
                  )}
                </div>

                {/* 📄 Inhalt */}
                <div style={{ padding: "12px" }}>
                  <h3 style={{ marginBottom: "8px" }}>{item.name}</h3>

                  <p style={{ fontSize: "12px", color: "#666" }}>
                    Erstellt am:{" "}
                    {item.created
                      ? new Date(item.created).toLocaleDateString()
                      : "-"}
                  </p>

                  <a
                    href={downloadUrl}
                    target="_blank"
                    style={{
                      display: "block",
                      marginTop: "10px",
                      background: "#0078d4",
                      color: "white",
                      textAlign: "center",
                      padding: "10px",
                      borderRadius: "6px",
                      textDecoration: "none",
                      fontWeight: "bold",
                    }}
                  >
                    Download
                  </a>
                </div>
              </div>
            );
          })}
        </div>
      </div>
    );
  }
}
``;
