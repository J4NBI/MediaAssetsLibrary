import * as React from "react";
import styles from "./MediaAssetsLib.module.scss";

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

      <div className={styles.itemContent}>
        <div>
          <h3>{item.name}</h3>

          <p>Ersteller: {item.createdBy || "-"}</p>

          <p>Kategorie: {item.category || "-"}</p>

          <p>Dienst: {item.dienst || "-"}</p>

          <div className={styles.tagList}>
            {(item.tags || []).map((tag: string, i: number) => (
              <span key={i} className={styles.tag}>
                {tag}
              </span>
            ))}
          </div>

          <p className={styles.itemDate}>
            Erstellt am:{" "}
            {item.created ? new Date(item.created).toLocaleDateString() : "-"}
          </p>
        </div>

        <div className={styles.itemActions}>
          <button onClick={onDownload} className={styles.downloadBtn}>
            {downloadingItemId === item.id ? "⏳ Lädt..." : "Download"}
          </button>

          <button onClick={onEdit} className={styles.editBtn}>
            Editieren
          </button>
        </div>
      </div>
    </div>
  );
};

export default FileCard;
