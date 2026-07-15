export function getLibraryPath(siteUrl: string, libraryName: string): string {
  return `${new URL(siteUrl).pathname}/${libraryName}`;
}
