/**
 * Ermittelt den SharePoint-Serverpfad zu einer Bibliothek.
 *
 * @param {string} siteUrl - Absolute SharePoint-Site-URL.
 * @param {string} libraryName - Name der Zielbibliothek.
 * @returns {string} Serverrelativer Pfad zur Bibliothek.
 */
export function getLibraryPath(siteUrl: string, libraryName: string): string {
  return `${new URL(siteUrl).pathname}/${libraryName}`;
}
