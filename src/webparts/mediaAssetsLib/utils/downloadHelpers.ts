import { SPHttpClient } from "@microsoft/sp-http";

/**
 * Lädt eine Datei authentifiziert als Blob herunter und speichert sie
 * plattformgerecht auf dem Gerät:
 * - iOS: über die Web Share API (natives "Sichern"-Menü)
 * - Android/Desktop: über einen Blob-URL-Download (umgeht den
 *   Android DownloadManager, der keine WebView-Session-Cookies hat)
 */
export async function downloadFileToDevice(
  fileUrl: string,
  fileName: string,
  spHttpClient: SPHttpClient,
): Promise<void> {
  const response = await spHttpClient.get(
    fileUrl,
    SPHttpClient.configurations.v1,
  );

  if (!response.ok) {
    throw new Error(`Download fehlgeschlagen: ${response.status}`);
  }

  const blob = await response.blob();
  const file = new File([blob], fileName, { type: blob.type });

  // iOS: Web Share API nutzen, falls verfügbar (WKWebView unterstützt
  // das download-Attribut nicht, aber navigator.share funktioniert)
  const nav = navigator as Navigator & {
    canShare?: (data: { files: File[] }) => boolean;
    share?: (data: { files: File[] }) => Promise<void>;
  };

  if (nav.canShare && nav.canShare({ files: [file] }) && nav.share) {
    await nav.share({ files: [file] });
    return;
  }

  // Android/Desktop-Fallback: Blob-URL statt Server-URL verwenden
  const blobUrl = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = blobUrl;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);

  // Blob-URL nach kurzer Zeit wieder freigeben
  setTimeout(() => URL.revokeObjectURL(blobUrl), 1000);
}
