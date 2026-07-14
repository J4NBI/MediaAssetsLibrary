/**
 * Erkennt den Dateityp basierend auf der Dateiendung
 * Kategorisiert Dateien in: Bild, Video, Audio oder Dokument
 * @private
 * @param {string} fileName - Der Dateiname mit Erweiterung
 * @returns {string} Der erkannte Dateityp (Bild, Video, Audio oder Dokument)
 * @example
 * detectFormat('photo.jpg')  // Returns: "Bild"
 * detectFormat('video.mp4')  // Returns: "Video"
 * detectFormat('song.mp3')   // Returns: "Audio"
 * detectFormat('doc.pdf')    // Returns: "Dokument"
 */

export const detectFormat = (fileName: string): string => {
  const ext = fileName.split(".").pop()?.toLowerCase();

  if (!ext) {
    return "Dokument";
  }

  const imageTypes = ["jpg", "jpeg", "png", "gif", "webp", "bmp", "svg"];
  const videoTypes = ["mp4", "mov", "avi", "webm", "mkv", "wmv", "m4v"];
  const audioTypes = ["mp3", "wav", "aiff", "aac", "flac", "ogg", "wma", "m4a"];

  if (imageTypes.includes(ext)) {
    return "Bild";
  }

  if (videoTypes.includes(ext)) {
    return "Video";
  }

  if (audioTypes.includes(ext)) {
    return "Audio";
  }

  return "Dokument";
};
