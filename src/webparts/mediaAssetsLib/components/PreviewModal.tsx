import * as React from "react";
import styles from "./MediaAssetsLib.module.scss";

interface IPreviewItem {
  name: string;
  fileRef: string;
}

interface IPreviewModalProps {
  item?: IPreviewItem;
  isOpen: boolean;
  onClose: () => void;
}

const PreviewModal: React.FC<IPreviewModalProps> = ({
  item,
  isOpen,
  onClose,
}) => {
  if (!isOpen || !item) {
    return null;
  }

  const fileType = item.name.split(".").pop()?.toLowerCase();

  const isVideo = fileType === "mp4" || fileType === "mov";

  const isAudio = [
    "mp3",
    "wav",
    "aiff",
    "aac",
    "flac",
    "ogg",
    "m4a",
    "wma",
  ].includes(fileType || "");

  const fileUrl = `${window.location.origin}${item.fileRef}`;

  return (
    <div className={`${styles.modalOverlay} ${styles.modalOverlayPreview}`}>
      <button onClick={onClose} className={styles.modalClose}>
        ✕
      </button>

      <div className={styles.modalContent}>
        {!isVideo && !isAudio && (
          <img src={fileUrl} className={styles.modalMedia} />
        )}

        {/* VIDEO */}
        {isVideo && (
          <video
            controls
            style={{
              maxWidth: "100%",
              maxHeight: "80vh",
            }}
          >
            <source src={fileUrl} />
          </video>
        )}

        {isAudio && (
          <audio
            controls
            style={{
              minWidth: "200px",
              marginTop: "20px",
            }}
          >
            <source src={fileUrl} />
          </audio>
        )}

        <button
          onClick={() => {
            const link = document.createElement("a");

            link.href = fileUrl;
            link.target = "_blank";
            link.download = item.name;

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
