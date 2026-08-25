import * as microsoftTeams from "@microsoft/teams-js";

let teamsInitialized = false;
let cachedIsInTeams: boolean | null = null;

/**
 * Prüft, ob die App aktuell im Teams-Mobile-Context läuft.
 * Ergebnis wird gecacht, da sich das während einer Session nicht ändert.
 */
export async function isRunningInTeamsMobile(): Promise<boolean> {
  if (cachedIsInTeams !== null) return cachedIsInTeams;

  try {
    if (!teamsInitialized) {
      await microsoftTeams.app.initialize();
      teamsInitialized = true;
    }
    const context = await microsoftTeams.app.getContext();
    const hostClientType = context.app.host.clientType;
    // android / ios = Teams Mobile App, sonst Desktop/Web/Browser
    cachedIsInTeams =
      hostClientType === microsoftTeams.HostClientType.android ||
      hostClientType === microsoftTeams.HostClientType.ios;
  } catch {
    // Kein Teams-Context verfügbar -> normaler Browser
    cachedIsInTeams = false;
  }

  return cachedIsInTeams;
}

/**
 * Wandelt ein Teams-Media-Objekt in ein echtes File-Objekt um,
 * damit der restliche Upload-Code (der File[] erwartet) unverändert bleibt.
 */
function getExtensionFromMimeType(mimeType: string): string {
  const map: { [key: string]: string } = {
    "image/jpeg": "jpg",
    "image/jpg": "jpg",
    "image/png": "png",
    "image/gif": "gif",
    "image/webp": "webp",
    "video/mp4": "mp4",
    "video/quicktime": "mov",
    "video/webm": "webm",
    "video/x-msvideo": "avi",
    "video/3gpp": "3gp",
  };

  if (map[mimeType]) return map[mimeType];

  // Kein exakter Treffer -> anhand des Präfixes sinnvoll raten,
  // statt fälschlich immer "jpg" zu wählen
  if (mimeType.startsWith("video/")) return "mp4";
  if (mimeType.startsWith("image/")) return "jpg";

  return "jpg"; // letzter Fallback, sollte praktisch nie greifen
}

async function teamsMediaToFile(
  media: microsoftTeams.media.Media,
): Promise<File> {
  return new Promise((resolve, reject) => {
    media.getMedia((error, blob) => {
      if (error || !blob) {
        reject(error ?? new Error("Kein Blob erhalten"));
        return;
      }

      const mimeType = media.mimeType || blob.type || "image/jpeg";

      let fileName = media.name;
      if (!fileName || !fileName.includes(".")) {
        const extension = getExtensionFromMimeType(mimeType);
        fileName = `datei_${Date.now()}.${extension}`;
      }

      const file = new File([blob], fileName, { type: mimeType });
      resolve(file);
    });
  });
}

/**
 * Öffnet den nativen Teams-Mobile-Picker für Bilder/Videos.
 * Unterstützt echte Mehrfachauswahl (bis zu 10 Dateien).
 */
export function selectMediaFilesViaTeams(): Promise<File[]> {
  return new Promise((resolve, reject) => {
    const mediaInputs: microsoftTeams.media.MediaInputs = {
      mediaType: microsoftTeams.media.MediaType.VideoAndImage,
      maxMediaCount: 10,
    };

    microsoftTeams.media.selectMedia(mediaInputs, async (error, mediaArr) => {
      if (error) {
        reject(error);
        return;
      }
      if (!mediaArr || mediaArr.length === 0) {
        resolve([]);
        return;
      }
      try {
        const files = await Promise.all(mediaArr.map(teamsMediaToFile));
        resolve(files);
      } catch (e) {
        reject(e);
      }
    });
  });
}
