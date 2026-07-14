import { IMediaItem } from "../models/types";

/**
 * Extrahiert alle einzigartigen Dateiformate aus den geladenen Elementen
 * @private
 * @returns {string[]} Array eindeutiger Format-Werte
 */

export const getUniqueFormats = (items: IMediaItem[]): string[] => {
  const values = items
    .map((item) => item.format)
    .filter((v): v is string => !!v);

  return Array.from(new Set(values));
};
