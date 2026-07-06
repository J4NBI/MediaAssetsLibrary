import * as React from "react";
import styles from "./MediaAssetsLib.module.scss";
import DownloadButton from "./DownloadButton";

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

        <DownloadButton
          fileUrl={fileUrl}
          fileName={item.name}
          className={`${styles.downloadBtn} ${styles.modalDownload}`}
        />
      </div>
    </div>
  );
};

export default PreviewModal;
