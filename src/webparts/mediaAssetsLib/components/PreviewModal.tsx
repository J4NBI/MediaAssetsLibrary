import * as React from "react";
import styles from "./MediaAssetsLib.module.scss";
import { downloadFileToDevice } from "../utils/downloadHelpers";
import { SPHttpClient } from "@microsoft/sp-http";
import { openInExternalBrowser } from "../utils/teamsFileHelpers";

interface IPreviewItem {
  name: string;
  fileRef: string;
}

interface IPreviewModalProps {
  item?: IPreviewItem;
  isOpen: boolean;
  onClose: () => void;
  spHttpClient: SPHttpClient;
}

const PreviewModal: React.FC<IPreviewModalProps> = ({
  item,
  isOpen,
  onClose,
  spHttpClient,
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
          onClick={async () => {
            const ua = navigator.userAgent || "";
            const isIOS = /iPhone|iPad|iPod/i.test(ua);
            const isAndroid = /Android/i.test(ua);
            const isAndroidTeamsWebView =
              isAndroid && (/Teams/i.test(ua) || /; wv\)/i.test(ua));

            try {
              if (isAndroidTeamsWebView) {
                await openInExternalBrowser(fileUrl);
              } else if (isIOS || isAndroid) {
                await downloadFileToDevice(fileUrl, item.name, spHttpClient);
              } else {
                const link = document.createElement("a");
                link.href = fileUrl;
                link.download = item.name;
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
              }
            } catch (error) {
              console.error("Download fehlgeschlagen:", error);
              alert("Download fehlgeschlagen. Bitte erneut versuchen.");
            }
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
