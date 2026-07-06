import * as React from "react";
import styles from "./MediaAssetsLib.module.scss";

interface IMediaItem {
  id: number;
  name: string;
  fileRef: string;
  tags?: string[];
  created?: string;
}

export interface IMediaCardProps {
  item: IMediaItem;
  downloadingItemId?: number;

  onPreview: (item: IMediaItem) => void;
  onDownload: (item: IMediaItem) => void;
  onEdit: (item: IMediaItem) => void;
  onTagClick: (tag: string) => void;
}

const MediaCard: React.FC<IMediaCardProps> = ({
  item,
  downloadingItemId,
  onPreview,
  onDownload,
  onEdit,
  onTagClick,
}) => {
  const fileUrl = `${window.location.origin}${item.fileRef}`;

  const fileType = item.name?.split(".").pop()?.toLowerCase();

  const isVideo = fileType === "mp4" || fileType === "mov";

  return (
    <div className={styles.itemCard}>
      {isVideo ? (
        <div className={styles.itemImg} onClick={() => onPreview(item)}>
          🎬
        </div>
      ) : (
        <img
          src={fileUrl}
          className={styles.itemImg}
          onClick={() => onPreview(item)}
        />
      )}

      <div className={styles.itemContent}>
        <h3>{item.name}</h3>

        <div className={styles.tagList}>
          {(item.tags || []).map((tag, i) => (
            <span
              key={i}
              className={styles.tag}
              onClick={() => onTagClick(tag)}
            >
              {tag}
            </span>
          ))}
        </div>

        <p className={styles.itemDate}>
          Erstellt am:{" "}
          {item.created ? new Date(item.created).toLocaleDateString() : "-"}
        </p>

        <div className={styles.itemActions}>
          <button
            onClick={() => onDownload(item)}
            className={styles.downloadBtn}
          >
            {downloadingItemId === item.id ? "⏳ Lädt..." : "Download"}
          </button>

          <button onClick={() => onEdit(item)} className={styles.editBtn}>
            Editieren
          </button>
        </div>
      </div>
    </div>
  );
};

export default MediaCard;
