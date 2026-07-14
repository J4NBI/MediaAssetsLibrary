/**
 * Gibt den Bibliotheks-Pfad für REST-API-Aufrufe zurück
 * Kombiniert die Site-URL mit dem Bibliotheksnamen
 * @private
 * @returns {string} Relativer Pfad zur Bibliothek
 * @example
 * // Returns: "/sites/mysite/Medienbibliothek"
 * getLibraryPath();
 */

export const getLibraryPath = (
  siteUrl: string,
  libraryName: string,
): string => {
  return `${new URL(siteUrl).pathname}/${libraryName}`;
};
