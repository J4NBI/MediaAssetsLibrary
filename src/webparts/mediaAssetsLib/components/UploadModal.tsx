import * as React from "react";
import styles from "./MediaAssetsLib.module.scss";
import BucketDropdown from "./BucketDropdown";
import type { IMediaAssetsLibState } from "./MediaAssetsLib";

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
  if (!isOpen) return null;

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
            <input
              id="uploadFileInput"
              type="file"
              multiple
              style={{ display: "none" }}
              onChange={(e) => {
                const files = e.target.files;
                if (!files || files.length === 0) return;

                const fileArray = Array.from(files);

                const firstFile = fileArray[0];

                const fileName = firstFile.name;
                const baseName = fileName.includes(".")
                  ? fileName.substring(0, fileName.lastIndexOf("."))
                  : fileName;

                const previewUrl = URL.createObjectURL(firstFile);

                setState({
                  uploadFiles: fileArray,
                  uploadName: baseName,
                  uploadPreviewUrl: previewUrl,
                });
              }}
            />
            <label htmlFor="uploadFileInput" className={styles.fileSelectBtn}>
              Datei auswählen
            </label>
            <div>{state.uploadFiles?.length} Dateien gewählt</div>
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
