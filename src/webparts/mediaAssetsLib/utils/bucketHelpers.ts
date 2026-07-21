import { IMediaItem } from "../models/types";
/**
 * Zählt die Anzahl der Elemente in jedem Bucket.
 *
 * @param {IMediaItem[]} items - Liste aller Medienelemente.
 * @returns {{ [key: string]: number }} Objekt mit Bucket-Namen als Schlüssel und Elementanzahl als Wert.
 */

export const getBucketCounts = (
  items: IMediaItem[],
): { [key: string]: number } => {
  const counts: { [key: string]: number } = {};

  items.forEach((item) => {
    item.bucket?.forEach((bucket) => {
      counts[bucket] = (counts[bucket] || 0) + 1;
    });
  });

  return counts;
};

/**
 * Sortiert Buckets nach dem neuesten Element in jedem Bucket (absteigend).
 *
 * @param {IMediaItem[]} items - Array von Medienelementen.
 * @returns {string[]} Sortierte Bucket-Namen (neueste zuerst).
 */

export const getBucketsSortedByNewest = (items: IMediaItem[]): string[] => {
  const map: { [key: string]: number } = {};

  items.forEach((item) => {
    if (item.bucket && item.modified) {
      const time = new Date(item.modified).getTime();
      item.bucket.forEach((bucket) => {
        if (!map[bucket] || map[bucket] < time) {
          map[bucket] = time;
        }
      });
    }
  });

  return Object.keys(map).sort((a, b) => map[b] - map[a]);
};

/**
 * Findet das erste Element eines bestimmten Buckets für die Vorschau.
 *
 * @param {IMediaItem[]} items - Liste aller Medienelemente.
 * @param {string} bucket - Der Bucket-Name.
 * @returns {IMediaItem|undefined} Das erste Element des Buckets oder undefined.
 */
export const getBucketPreview = (
  items: IMediaItem[],
  bucket: string,
): IMediaItem | undefined => {
  const bucketItems = items.filter((item) => item.bucket?.includes(bucket));

  bucketItems.sort((a, b) => {
    const dateA = a.modified ? new Date(a.modified).getTime() : 0;

    const dateB = b.modified ? new Date(b.modified).getTime() : 0;

    return dateB - dateA;
  });

  return bucketItems[0];
};

/**
 * Filtert Buckets basierend auf aktiven Filtern und Suchtext.
 * Ein Bucket wird angezeigt, wenn er mindestens ein passendes Element enthält.
 *
 * @param {string[]} buckets - Verfügbare Bucket-Namen.
 * @param {IMediaItem[]} items - Liste aller Medienelemente.
 * @param {object} filters - Aktive Filterkriterien für Suche und Metadaten.
 * @returns {string[]} Array der gefilterten Bucket-Namen.
 */

export const getFilteredBuckets = (
  buckets: string[],
  items: IMediaItem[],
  filters: {
    searchText: string;
    filterCategory?: string;
    filterDienst?: string;
    filterFormat?: string;
    filterYear?: number;
    filterMonth?: number;
    filterCreator?: string;
  },
): string[] => {
  const {
    searchText,
    filterCategory,
    filterDienst,
    filterFormat,
    filterYear,
    filterMonth,
    filterCreator,
  } = filters;

  return buckets.filter((bucket) => {
    const bucketItems = items.filter((item) => item.bucket?.includes(bucket));

    const search = searchText.toLowerCase();

    const filteredItems = bucketItems.filter((item) => {
      const matchesText =
        !search ||
        item.name.toLowerCase().includes(search) ||
        (item.tags || []).join(" ").toLowerCase().includes(search);

      const matchesCategory =
        !filterCategory || item.category === filterCategory;

      const matchesDienst = !filterDienst || item.dienst === filterDienst;

      const matchesFormat = !filterFormat || item.format === filterFormat;

      const date = item.created ? new Date(item.created) : null;

      const matchesYear =
        !filterYear || (date && date.getFullYear() === filterYear);

      const matchesMonth =
        !filterMonth || (date && date.getMonth() + 1 === filterMonth);

      const matchesCreator = !filterCreator || item.createdBy === filterCreator;

      return (
        matchesText &&
        matchesCategory &&
        matchesDienst &&
        matchesFormat &&
        matchesYear &&
        matchesMonth &&
        matchesCreator
      );
    });

    const bucketMatchesSearch = search
      ? bucket.toLowerCase().includes(search)
      : false;

    if (search) {
      return bucketMatchesSearch || filteredItems.length > 0;
    }

    return filteredItems.length > 0;
  });
};
