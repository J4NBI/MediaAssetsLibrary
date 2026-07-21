import { IMediaItem } from "../models/types";

/**
 * Extrahiert alle einzigartigen Dateiformate aus den geladenen Elementen.
 *
 * @param {IMediaItem[]} items - Liste aller geladenen Medienelemente.
 * @returns {string[]} Array eindeutiger Format-Werte.
 */

export const getUniqueFormats = (items: IMediaItem[]): string[] => {
  const values = items
    .map((item) => item.format)
    .filter((v): v is string => !!v);

  return Array.from(new Set(values));
};

/**
 * Extrahiert alle einzigartigen Jahre aus den geladenen Elementen.
 * Sortiert sie in absteigender Reihenfolge (neueste zuerst).
 *
 * @param {IMediaItem[]} items - Liste aller geladenen Medienelemente.
 * @returns {number[]} Array einzigartiger Jahre.
 */
export const getUniqueYears = (items: IMediaItem[]): number[] => {
  const years = items
    .map((item) => {
      if (!item.created) {
        return null;
      }

      const date = new Date(item.created);

      return date.getFullYear();
    })
    .filter((y): y is number => y !== null);

  return Array.from(new Set(years)).sort((a, b) => b - a);
};
/**
 * Extrahiert alle einzigartigen Ersteller aus den geladenen Elementen.
 * Sortiert sie alphabetisch.
 *
 * @param {IMediaItem[]} items - Liste aller geladenen Medienelemente.
 * @returns {string[]} Array eindeutiger Ersteller-Namen.
 */

export const getUniqueCreators = (items: IMediaItem[]): string[] => {
  const values = items
    .map((item) => item.createdBy)
    .filter((v): v is string => !!v);

  return Array.from(new Set(values)).sort();
};
