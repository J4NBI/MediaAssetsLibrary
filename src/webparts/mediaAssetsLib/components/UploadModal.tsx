import * as React from "react";
import styles from "./MediaAssetsLib.module.scss";
import BucketDropdown from "./BucketDropdown";
import type { IMediaAssetsLibState } from "./MediaAssetsLib";
import {
  isRunningInTeamsMobile,
  selectMediaFilesViaTeams,
} from "../utils/teamsFileHelpers";

interface IUploadModalProps {
  isOpen: boolean;
  setState: (state: Partial<IMediaAssetsLibState>) => void;
  state: IMediaAssetsLibState;
  onClose: () => void;
  onUpload: () => void;
}

const UploadModal: React.FC<IUploadModalProps> = ({
  isOpen,
  state,
  setState,
  onClose,
  onUpload,
}) => {
  const [useFallbackPicker, setUseFallbackPicker] = React.useState(false);
  if (!isOpen) return null;
  const isEmbeddedAndroidWebView = (): boolean => {
    const ua = navigator.userAgent || "";
    const isAndroid = /Android/i.test(ua);
    if (!isAndroid) return false;
    // Android System WebView (Basis von Teams-App & SharePoint-App) enthält "; wv)".
    // Echtes Android-Chrome enthält das nicht.
    return /; wv\)/i.test(ua) || /Teams/i.test(ua) || /SharePoint/i.test(ua);
  };
  const handleFilesSelected = (
    newFiles: File[],
    append: boolean = false,
  ): void => {
    if (newFiles.length === 0) return;

    const existingFiles = append ? state.uploadFiles || [] : [];

    // Duplikate vermeiden (gleicher Name + gleiche Größe = vermutlich gleiche Datei)
    const combined = [...existingFiles];
    newFiles.forEach((f) => {
      const isDuplicate = combined.some(
        (existing) => existing.name === f.name && existing.size === f.size,
      );
      if (!isDuplicate) combined.push(f);
    });

    const firstFile = combined[0];
    const fileName = firstFile.name;
    const baseName = fileName.includes(".")
      ? fileName.substring(0, fileName.lastIndexOf("."))
      : fileName;

    // Preview nur beim ersten Mal setzen, nicht bei jedem weiteren Hinzufügen überschreiben
    const previewUrl =
      state.uploadPreviewUrl && append
        ? state.uploadPreviewUrl
        : URL.createObjectURL(firstFile);

    setState({
      uploadFiles: combined,
      uploadName: baseName,
      uploadPreviewUrl: previewUrl,
    });
  };

  const handlePickerClick = async (e: React.MouseEvent): Promise<void> => {
    if (!isEmbeddedAndroidWebView()) {
      // Desktop, Android-Chrome, iPhone -> nativer Multi-Picker funktioniert, nichts tun
      return;
    }

    e.preventDefault(); // nur hier eingreifen, da <input multiple> nachweislich buggy ist

    let inTeams = false;
    try {
      inTeams = await isRunningInTeamsMobile();
    } catch {
      inTeams = false;
    }

    if (inTeams) {
      try {
        const files = await selectMediaFilesViaTeams();
        if (files.length === 0) return;
        handleFilesSelected(files, false);
        return;
      } catch (err) {
        console.error(
          "Teams Media Picker fehlgeschlagen, nutze Fallback:",
          err,
        );
      }
    }

    // SharePoint-App oder Teams-js-Fehlschlag -> Einzelauswahl-Fallback
    setUseFallbackPicker(true);
    document.getElementById("uploadFileInputSingle")?.click();
  };

  const handleRemoveFile = (index: number): void => {
    const updated = [...(state.uploadFiles || [])];
    updated.splice(index, 1);
    setState({ uploadFiles: updated });
  };

  return (
    <>
      {state.isUploadOpen && (
        <div className={`${styles.modalOverlay}`}>
          <div className={`${styles.modalBox} ${styles.uploadBox}`}>
            <h3>Upload</h3>
            {/* MODAL PREVIEW */}
            {state.uploadFiles?.[0] && state.uploadPreviewUrl && (
              <div className={styles.uploadPreview}>
                {(() => {
                  const file = state.uploadFiles![0];
                  const fileType = file.name.split(".").pop()?.toLowerCase();

                  const isImage = [
                    "jpg",
                    "jpeg",
                    "png",
                    "gif",
                    "webp",
                  ].includes(fileType || "");
                  const isVideo = ["mp4", "mov", "webm"].includes(
                    fileType || "",
                  );

                  if (isImage) {
                    return (
                      <img
                        src={state.uploadPreviewUrl}
                        className={styles.uploadPreviewMedia}
                      />
                    );
                  }

                  if (isVideo) {
                    return (
                      <video
                        src={state.uploadPreviewUrl}
                        controls
                        className={styles.uploadPreviewMedia}
                      />
                    );
                  }

                  return (
                    <div className={styles.uploadFallback}>📄 {file.name}</div>
                  );
                })()}
              </div>
            )}
            {/* Unsichtbarer Einzel-Input als Fallback (SharePoint-App / Teams-js-Fehler) */}
            <input
              id="uploadFileInputSingle"
              type="file"
              accept="image/*,video/*,audio/*"
              style={{ display: "none" }}
              onChange={(e) => {
                const files = e.target.files;
                if (!files || files.length === 0) return;
                handleFilesSelected(Array.from(files), true);
                e.target.value = ""; // Reset, damit dieselbe Datei erneut wählbar bleibt
              }}
            />

            {/* Normaler Multi-Input für Desktop/Browser-Kontext */}
            <input
              id="uploadFileInput"
              type="file"
              multiple={true}
              accept="image/*,video/*,audio/*"
              style={{ display: "none" }}
              onChange={(e) => {
                const files = e.target.files;
                if (!files || files.length === 0) return;
                handleFilesSelected(Array.from(files), false);
              }}
            />

            <label
              htmlFor="uploadFileInput"
              className={styles.fileSelectBtn}
              onClick={handlePickerClick}
            >
              Datei auswählen
            </label>

            {useFallbackPicker &&
              state.uploadFiles &&
              state.uploadFiles.length > 0 && (
                <button
                  type="button"
                  className={styles.fileSelectBtn}
                  onClick={() =>
                    document.getElementById("uploadFileInputSingle")?.click()
                  }
                  style={{ marginTop: 8 }}
                >
                  + Weitere Datei hinzufügen
                </button>
              )}

            <div>{state.uploadFiles?.length || 0} Dateien gewählt</div>

            {useFallbackPicker &&
              state.uploadFiles &&
              state.uploadFiles.length > 0 && (
                <ul
                  className={styles.tagList}
                  style={{ listStyle: "none", padding: 0 }}
                >
                  {state.uploadFiles.map((f, index) => (
                    <li key={`${f.name}-${index}`} className={styles.tag}>
                      {f.name}
                      <span
                        onClick={() => handleRemoveFile(index)}
                        className={styles.tagRemove}
                      >
                        ✕
                      </span>
                    </li>
                  ))}
                </ul>
              )}
            {(!state.uploadFiles || state.uploadFiles.length <= 1) && (
              <input
                className={styles.tagInput}
                type="text"
                placeholder="Name"
                value={state.uploadName}
                onChange={(e) => setState({ uploadName: e.target.value })}
              />
            )}

            <BucketDropdown
              options={state.bucketOptions}
              selected={state.uploadBucket}
              onChange={(values) =>
                setState({
                  uploadBucket: values,
                })
              }
            />

            <select
              value={state.uploadCategory}
              onChange={(e) => setState({ uploadCategory: e.target.value })}
            >
              <option value="">Kategorie wählen</option>

              {state.categoryOptions.map((cat: string) => (
                <option key={cat} value={cat}>
                  {cat}
                </option>
              ))}
            </select>

            <select
              value={state.uploadDienst}
              onChange={(e) => setState({ uploadDienst: e.target.value })}
            >
              <option value="">Dienst wählen</option>

              {state.dienstOptions?.map((dienst: string) => (
                <option key={dienst} value={dienst}>
                  {dienst}
                </option>
              ))}
            </select>

            <div>
              <div className={styles.tagList}>
                {state.uploadTags.map((tag: string, index: number) => (
                  <span
                    key={index}
                    className={`${styles.tag} ${styles.editTag}`}
                  >
                    {tag}
                    <span
                      onClick={() => {
                        const newTags = [...state.uploadTags];
                        newTags.splice(index, 1);
                        setState({ uploadTags: newTags });
                      }}
                      className={styles.tagRemove}
                    >
                      ✕
                    </span>
                  </span>
                ))}
              </div>

              <input
                type="text"
                placeholder="Tag hinzufügen + Enter"
                className={styles.tagInput}
                onKeyDown={(e) => {
                  if (e.key === "Enter") {
                    e.preventDefault();

                    const value = (e.target as HTMLInputElement).value.trim();
                    if (!value) return;

                    setState({
                      uploadTags: [...state.uploadTags, value],
                    });

                    (e.target as HTMLInputElement).value = "";
                  }
                }}
              />
            </div>

            <button onClick={onClose} className={styles.editBtn}>
              Schließen
            </button>
            {state.isUploading && (
              <div style={{ marginTop: 20 }}>
                <div>
                  Datei {state.uploadCurrentFile} von {state.uploadTotalFiles}
                </div>

                <div
                  style={{
                    width: "100%",
                    height: "12px",
                    background: "#ddd",
                    marginTop: "10px",
                  }}
                >
                  <div
                    style={{
                      width: `${state.uploadProgress}%`,
                      height: "100%",
                      background: "#ecdd04",
                      transition: "width 0.2s ease",
                    }}
                  />
                </div>

                <div style={{ marginTop: 6 }}>{state.uploadProgress}%</div>
              </div>
            )}
            <button
              onClick={onUpload}
              disabled={
                state.isUploading ||
                !state.uploadBucket ||
                state.uploadBucket.length === 0 ||
                !state.uploadCategory
              }
              className={`${styles.uploadBtn} ${
                state.isUploading ||
                !state.uploadBucket ||
                state.uploadBucket.length === 0 ||
                !state.uploadCategory
                  ? styles.disabled
                  : ""
              }`}
            >
              {state.isUploading ? "⏳ Wird hochgeladen..." : "Hochladen"}
            </button>
          </div>
        </div>
      )}
    </>
  );
};

export default UploadModal;
