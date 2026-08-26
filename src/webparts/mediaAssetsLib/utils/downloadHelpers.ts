import { SPHttpClient } from "@microsoft/sp-http";

/**
 * Lädt eine Datei authentifiziert als Blob herunter und speichert sie
 * plattformgerecht auf dem Gerät:
 * - iOS: über die Web Share API (natives "Sichern"-Menü)
 * - Android/Desktop: über einen Blob-URL-Download (umgeht den
 *   Android DownloadManager, der keine WebView-Session-Cookies hat)
 */
function getMimeTypeFromFileName(fileName: string): string {
  const ext = fileName.split(".").pop()?.toLowerCase() || "";
  const map: { [key: string]: string } = {
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    png: "image/png",
    gif: "image/gif",
    webp: "image/webp",
    mp4: "video/mp4",
    mov: "video/quicktime",
    webm: "video/webm",
    avi: "video/x-msvideo",
    m4v: "video/x-m4v",
  };
  return map[ext] || "application/octet-stream";
}

export async function downloadFileToDevice(
  fileUrl: string,
  fileName: string,
  spHttpClient: SPHttpClient,
): Promise<void> {
  const response = await spHttpClient.get(
    fileUrl,
    SPHttpClient.configurations.v1,
  );
  alert("Download geklickt");

  if (!response.ok) {
    throw new Error(`Download fehlgeschlagen: ${response.status}`);
  }

  const rawBlob = await response.blob();
  const mimeType =
    rawBlob.type && rawBlob.type !== "application/octet-stream"
      ? rawBlob.type
      : getMimeTypeFromFileName(fileName);

  // Neuer Blob mit korrektem, erzwungenem MIME-Typ
  const blob = new Blob([rawBlob], { type: mimeType });
  console.log("Downloaded blob type:", blob.type, "size:", blob.size);
  const file = new File([blob], fileName, { type: mimeType });

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
