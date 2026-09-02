import { SPHttpClient } from "@microsoft/sp-http";
import { openInExternalBrowser } from "./teamsFileHelpers";

/**
 * Download-Hilfsfunktionen für die Media Library.
 *
 * Zielverhalten je Plattform:
 * - Desktop: klassischer <a download>-Link
 * - iPhone / Android-Chrome (normaler Browser): Blob-Download bzw.
 *   Web-Share-API ("Sichern"-Menü)
 * - Android in der Teams-App (eingebettete WebView): Datei-URL im
 *   externen System-Browser öffnen, da die Teams-Android-WebView
 *   Datei-Downloads laut Microsoft offiziell nicht unterstützt
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

function detectPlatform(): "androidTeamsWebView" | "mobile" | "desktop" {
  const ua = navigator.userAgent || "";
  const isIOS = /iPhone|iPad|iPod/i.test(ua);
  const isAndroid = /Android/i.test(ua);
  const isAndroidTeamsWebView =
    isAndroid && (/Teams/i.test(ua) || /; wv\)/i.test(ua));

  if (isAndroidTeamsWebView) return "androidTeamsWebView";
  if (isIOS || isAndroid) return "mobile";
  return "desktop";
}

async function downloadAsBlob(
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

  const rawBlob = await response.blob();
  const mimeType =
    rawBlob.type && rawBlob.type !== "application/octet-stream"
      ? rawBlob.type
      : getMimeTypeFromFileName(fileName);

  const blob = new Blob([rawBlob], { type: mimeType });
  const file = new File([blob], fileName, { type: mimeType });

  const nav = navigator as Navigator & {
    canShare?: (data: { files: File[] }) => boolean;
    share?: (data: { files: File[] }) => Promise<void>;
  };

  if (nav.canShare && nav.canShare({ files: [file] }) && nav.share) {
    await nav.share({ files: [file] });
    return;
  }

  const blobUrl = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = blobUrl;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);

  setTimeout(() => URL.revokeObjectURL(blobUrl), 1000);
}

function downloadViaAnchor(fileUrl: string, fileName: string): void {
  const link = document.createElement("a");
  link.href = fileUrl;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

/**
 * Zentrale Download-Funktion für die gesamte Media Library.
 * Entscheidet automatisch anhand der Plattform, welcher Weg genutzt wird.
 */
export async function downloadItem(
  fileUrl: string,
  fileName: string,
  spHttpClient: SPHttpClient,
): Promise<void> {
  const platform = detectPlatform();

  switch (platform) {
    case "androidTeamsWebView":
      await openInExternalBrowser(fileUrl);
      break;
    case "mobile":
      await downloadAsBlob(fileUrl, fileName, spHttpClient);
      break;
    case "desktop":
      downloadViaAnchor(fileUrl, fileName);
      break;
  }
}
