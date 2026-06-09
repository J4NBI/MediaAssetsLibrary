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

  // ✅ rekursiv
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

      // ✅ nach Veröffentlichungsdatum sortieren (fallback: erstellt)
      items = items.sort((a, b) => {
        const dateA = new Date(a.published || a.created || 0).getTime();
        const dateB = new Date(b.published || b.created || 0).getTime();
        return dateB - dateA;
      });

      this.setState({
        allItems: items,
        visibleItems: items.slice(0, 3), // ✅ nur 3 neueste
      });
    } catch (error) {
      console.error(error);
    }
  }

  private applyFilters = (): void => {
    const { allItems, searchText, filterTag, filterCategory } = this.state;

    let filtered = [...allItems];

    // 🔍 Suche
    const search = searchText.toLowerCase();

    if (search) {
      filtered = filtered.filter(
        (item) =>
          item.name.toLowerCase().includes(search) ||
          (item.tags && item.tags.toLowerCase().includes(search)) ||
          (item.notes && item.notes.toLowerCase().includes(search)),
      );
    }

    // 🎯 Tags
    if (filterTag) {
      filtered = filtered.filter(
        (item) => item.tags && item.tags.includes(filterTag),
      );
    }

    // 🎯 Kategorie
    if (filterCategory) {
      filtered = filtered.filter((item) => item.category === filterCategory);
    }

    this.setState({
      visibleItems: filtered,
    });
  };

  private onSearchChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    this.setState({ searchText: e.target.value }, this.applyFilters);
  };

  // ✅ Dynamische Werte für Dropdown
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
              <option key={tag} value={tag}>
                {tag}
              </option>
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
              <option key={cat} value={cat}>
                {cat}
              </option>
            ))}
          </select>
        </div>

        {/* 📦 Anzeige */}
        <div
          style={{
            display: "flex",
            flexWrap: "wrap",
            gap: "20px",
            marginTop: "20px",
          }}
        >
          {this.state.visibleItems.map((item) => (
            <div
              key={item.id}
              style={{
                width: "300px",
                border: "1px solid #ddd",
                padding: "10px",
                borderRadius: "8px",
              }}
            >
              {`${window.location.origin}${item.fileRef}`}

              <h3>{item.name}</h3>

              <p>
                <b>Kategorie:</b> {item.category}
              </p>
              <p>
                <b>Tags:</b> {item.tags}
              </p>
              <p>
                <b>Format:</b> {item.format}
              </p>
            </div>
          ))}
        </div>
      </div>
    );
  }
}
