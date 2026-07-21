import { IMediaItem } from "../models/types";

/**
 * Struktur für eine gruppierten Monatssicht in der Bucket-Übersicht.
 */
export interface IMonthGroup {
  /** Eindeutiger Schlüssel im Format JJJJ-MM. */
  key: string;
  /** Kalenderjahr des Monats. */
  year: number;
  /** Kalendermonat von 1 bis 12. */
  month: number;
  /** Benutzerfreundliche Monatsbezeichnung. */
  label: string;
}

const MONTH_NAMES_DE = [
  "Januar",
  "Februar",
  "März",
  "April",
  "Mai",
  "Juni",
  "Juli",
  "August",
  "September",
  "Oktober",
  "November",
  "Dezember",
];

function getMonthKey(date: Date): string {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function getMonthLabel(year: number, month: number): string {
  return `${MONTH_NAMES_DE[month - 1]} ${year}`;
}

/**
 * Findet das Datum des neuesten Elements innerhalb eines Buckets.
 * Gibt undefined zurück, wenn der Bucket keine Elemente mit Datum hat.
 *
 * @param {IMediaItem[]} items - Liste aller Medienelemente.
 * @param {string} bucket - Name des Buckets.
 * @returns {Date|undefined} Das neueste Änderungsdatum des Buckets oder undefined.
 */
export function getBucketLatestDate(
  items: IMediaItem[],
  bucket: string,
): Date | undefined {
  const dates = items
    .filter((item) => item.bucket?.includes(bucket))
    .map((item) => (item.modified ? new Date(item.modified) : undefined))
    .filter((d): d is Date => !!d);

  if (dates.length === 0) return undefined;

  return new Date(Math.max(...dates.map((d) => d.getTime())));
}

/**
 * Gruppiert alle Buckets nach Jahr/Monat ihres jeweils neuesten Elements.
 * Buckets ohne verwertbares Datum werden ausgelassen.
 * Sortiert absteigend (neuester Monat zuerst).
 *
 * @param {IMediaItem[]} items - Liste aller Medienelemente.
 * @param {string[]} buckets - Verfügbare Bucket-Namen.
 * @returns {IMonthGroup[]} Array gruppierter Monate mit Bucket-Zuordnung.
 */
export function getMonthGroups(
  items: IMediaItem[],
  buckets: string[],
): IMonthGroup[] {
  const groupMap = new Map<string, IMonthGroup>();

  buckets.forEach((bucket) => {
    const latest = getBucketLatestDate(items, bucket);
    if (!latest) return;

    const key = getMonthKey(latest);

    if (!groupMap.has(key)) {
      groupMap.set(key, {
        key,
        year: latest.getFullYear(),
        month: latest.getMonth() + 1,
        label: getMonthLabel(latest.getFullYear(), latest.getMonth() + 1),
      });
    }
  });

  return Array.from(groupMap.values()).sort((a, b) =>
    b.key.localeCompare(a.key),
  );
}

/**
 * Liefert alle Bucket-Namen, deren neuestes Element in den angegebenen Monat fällt.
 *
 * @param {IMediaItem[]} items - Liste aller Medienelemente.
 * @param {string[]} buckets - Verfügbare Bucket-Namen.
 * @param {number} year - Zieljahr.
 * @param {number} month - Zielmonat (1-12).
 * @returns {string[]} Buckets, deren neuestes Element in den angegebenen Monat fällt.
 */
export function getBucketsForMonth(
  items: IMediaItem[],
  buckets: string[],
  year: number,
  month: number,
): string[] {
  return buckets.filter((bucket) => {
    const latest = getBucketLatestDate(items, bucket);
    return (
      !!latest &&
      latest.getFullYear() === year &&
      latest.getMonth() + 1 === month
    );
  });
}

/**
 * Liefert das neueste Element innerhalb eines Jahr/Monats für die Vorschau-Kachel,
 * unabhängig davon, in welchem Bucket es liegt.
 *
 * @param {IMediaItem[]} items - Liste aller Medienelemente.
 * @param {number} year - Zieljahr.
 * @param {number} month - Zielmonat (1-12).
 * @returns {IMediaItem|undefined} Das neueste Element des Monats oder undefined.
 */
export function getMonthPreview(
  items: IMediaItem[],
  year: number,
  month: number,
): IMediaItem | undefined {
  const monthItems = items.filter((item) => {
    if (!item.modified) return false;

    const date = new Date(item.modified);

    return date.getFullYear() === year && date.getMonth() + 1 === month;
  });

  if (monthItems.length === 0) return undefined;

  return monthItems.reduce((latest, item) => {
    const latestTime = latest.modified
      ? new Date(latest.modified).getTime()
      : 0;

    const itemTime = item.modified ? new Date(item.modified).getTime() : 0;
    return itemTime > latestTime ? item : latest;
  });
}
