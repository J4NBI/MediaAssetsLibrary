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
