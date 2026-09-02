import * as React from "react";
import styles from "./MediaAssetsLib.module.scss";
import { Icon } from "@fluentui/react";

interface IFileCardProps {
  item: {
    id: number;
    name: string;
    fileRef: string;
    category?: string;
    dienst?: string;
    createdBy?: string;
    created?: string;
    tags?: string[];
    thumbnailUrl?: string;
  };

  downloadingItemId?: number;

  onPreview: () => void;
  onEdit: () => void;
  onDownload: () => void;
}

const FileCard: React.FC<IFileCardProps> = ({
  item,
  downloadingItemId,
  onPreview,
  onEdit,
  onDownload,
}) => {
  const fileUrl = `${window.location.origin}${item.fileRef}`;

  const fileType = item.name?.split(".").pop()?.toLowerCase();

  const isVideo = fileType === "mp4" || fileType === "mov";

  const isAudio = ["mp3", "wav", "aiff", "aac", "flac", "ogg", "m4a"].includes(
    fileType || "",
  );

  return (
    <div className={styles.itemCard}>
      {isVideo && (
        <img
          src={item.thumbnailUrl}
          className={styles.videoImg}
          onClick={onPreview}
          style={{ cursor: "pointer" }}
        />
      )}

      {isAudio && (
        <div
          className={styles.itemImg}
          onClick={onPreview}
          style={{
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            fontSize: "50px",
            background: "#f3f2f1",
            cursor: "pointer",
          }}
        >
          🔊
        </div>
      )}

      {!isVideo && !isAudio && (
        <img src={fileUrl} className={styles.itemImg} onClick={onPreview} />
      )}

      <div className={styles.fileCardBody}>
        <div className={styles.fileCardHeader}>
          <h3 className={styles.fileCardName}>{item.name}</h3>

          <button
            className={styles.fileCardShareBtn}
            onClick={async () => {
              const shareUrl =
                `${window.location.origin}${window.location.pathname}` +
                `?item=${item.id}`;

              if (navigator.share) {
                await navigator.share({
                  title: item.name,
                  text: item.name,
                  url: shareUrl,
                });
              } else {
                await navigator.clipboard.writeText(shareUrl);
                alert("Link kopiert");
              }
            }}
          >
            <Icon iconName="Share" style={{ fontSize: "22px" }} />
          </button>
        </div>

        <div className={styles.fileMeta}>
          <span className={styles.fileMetaItem}>
            <Icon iconName="Contact" />
            {item.createdBy || "-"}
          </span>

          <span className={styles.fileMetaItem}>
            <Icon iconName="FolderHorizontal" />
            {item.category || "-"}
          </span>
        </div>

        <div className={styles.tagList}>
          {(item.tags || []).map((tag: string, i: number) => (
            <span key={i} className={`${styles.tag} ${styles.tagAccent}`}>
              {tag}
            </span>
          ))}
        </div>

        <div className={styles.itemActions} style={{ flexDirection: "row" }}>
          <button onClick={onDownload} className={styles.downloadBtn}>
            {downloadingItemId === item.id ? "⏳ Lädt..." : "Download"}
          </button>

          <button onClick={onEdit} className={styles.editBtn}>
            Bearbeiten
          </button>
        </div>
      </div>
    </div>
  );
};

export default FileCard;
