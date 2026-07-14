import { IMediaItem } from "../models/types";
/**
 * Zählt die Anzahl der Elemente in jedem Bucket
 * @private
 * @returns {{ [key: string]: number }} Objekt mit Bucket-Namen als Schlüssel und Elementanzahl als Wert
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
 * Sortiert Buckets nach dem neuesten Element in jedem Bucket (absteigend)
 * @private
 * @param {IMediaItem[]} items - Array von Medienelementen
 * @returns {string[]} Sortierte Bucket-Namen (neueste zuerst)
 */

export const getBucketsSortedByNewest = (items: IMediaItem[]): string[] => {
  const map: { [key: string]: number } = {};

  items.forEach((item) => {
    if (item.bucket && item.created) {
      const time = new Date(item.created).getTime();

      item.bucket.forEach((bucket) => {
        if (!map[bucket] || map[bucket] < time) {
          map[bucket] = time;
        }
      });
    }
  });

  return Object.keys(map).sort((a, b) => map[b] - map[a]);
};
