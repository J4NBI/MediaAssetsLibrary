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
}

interface IMediaAssetsLibState {
  allItems: IMediaItem[];
  visibleItems: IMediaItem[];
  searchText: string;
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
        category: file.ListItemAllFields?.Kategorie,
        notes: file.ListItemAllFields?.Notizen,
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
    const { allItems, searchText, filterCategory } = this.state;

    let filtered = [...allItems];
    const search = searchText.toLowerCase();

    if (search) {
      filtered = filtered.filter((item) => {
        const name = item.name?.toLowerCase() || "";
        const notes = item.notes?.toLowerCase() || "";

        return name.includes(search) || notes.includes(search);
      });
    }

    if (filterCategory) {
      filtered = filtered.filter((item) => item.category === filterCategory);
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

  public render(): React.ReactElement<IMediaAssetsLibProps> {
    const categoryOptions = this.getUniqueCategories();

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

        {/* 🎯 Kategorie */}
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
            // ✅ richtige URLs
            const videoUrl = `${window.location.origin}${item.fileRef}?web=1`;
            const imageUrl = `${window.location.origin}/_layouts/15/getpreview.ashx?path=${encodeURIComponent(
              item.fileRef,
            )}&resolution=1`;

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
                {/* ✅ MEDIA */}
                {isVideo ? (
                  <div
                    onClick={() => window.open(videoUrl, "_blank")}
                    style={{ position: "relative", cursor: "pointer" }}
                  >
                    <img
                      src={imageUrl}
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
                    src={imageUrl}
                    style={{
                      width: "100%",
                      height: "200px",
                      objectFit: "cover",
                    }}
                  />
                )}

                {/* ✅ INFO */}
                <div style={{ padding: "12px" }}>
                  <h3>{item.name}</h3>

                  <p style={{ fontSize: "12px", color: "#666" }}>
                    Erstellt am:{" "}
                    {item.created
                      ? new Date(item.created).toLocaleDateString()
                      : "-"}
                  </p>

                  <a
                    href={videoUrl}
                    target="_blank"
                    rel="noopener noreferrer"
                    style={{
                      display: "block",
                      marginTop: "10px",
                      padding: "10px",
                      background: "#0078d4",
                      color: "white",
                      textAlign: "center",
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
