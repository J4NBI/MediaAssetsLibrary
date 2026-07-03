import * as React from "react";
import styles from "./MediaAssetsLib.module.scss";

interface IMediaItem {
  id: number;
  name: string;
  fileRef: string;
}

export interface IPreviewModalProps {
  item: IMediaItem;
  onClose: () => void;
}

const PreviewModal: React.FC<IPreviewModalProps> = ({ item, onClose }) => {
  const fileType = item.name.split(".").pop()?.toLowerCase();

  const isVideo =
    fileType === "mp4" || fileType === "mov" || fileType === "webm";

  const fileUrl = `${window.location.origin}${item.fileRef}`;

  const downloadUrl = `${
    window.location.origin
  }/_layouts/15/download.aspx?SourceUrl=${encodeURIComponent(fileUrl)}`;
  console.log("Name:", item.name);
  console.log("FileType:", fileType);
  console.log("IsVideo:", isVideo);
  console.log("FileUrl:", fileUrl);
  return (
    <div className={`${styles.modalOverlay} ${styles.modalOverlayPreview}`}>
      <button onClick={onClose} className={styles.modalClose}>
        ✕
      </button>

      <div className={styles.modalContent}>
        {isVideo ? (
          <video
            controls
            style={{
              maxWidth: "100%",
              maxHeight: "80vh",
            }}
          >
            <source src={fileUrl} />
          </video>
        ) : (
          <img src={fileUrl} className={styles.modalMedia} />
        )}

        <button
          onClick={() => {
            const link = document.createElement("a");

            link.href = downloadUrl;
            link.setAttribute("download", item.name || "file");

            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
          }}
          className={`${styles.downloadBtn} ${styles.modalDownload}`}
        >
          Download
        </button>
      </div>
    </div>
  );
};

export default PreviewModal;
