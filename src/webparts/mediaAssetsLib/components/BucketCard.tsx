import * as React from "react";
import styles from "./MediaAssetsLib.module.scss";

interface IMediaItem {
  id: number;
  name: string;
  fileRef: string;
}

export interface IBucketCardProps {
  bucket: string;
  count: number;
  preview?: IMediaItem;

  onOpen: (bucket: string) => void;
  onUpload: (bucket: string) => void;
}

const BucketCard: React.FC<IBucketCardProps> = ({
  bucket,
  count,
  preview,
  onOpen,
  onUpload,
}) => {
  const fileType = preview?.name?.split(".").pop()?.toLowerCase();

  const isVideo = fileType === "mp4" || fileType === "mov";

  const imageUrl = preview
    ? `${window.location.origin}/_layouts/15/getpreview.ashx?path=${encodeURIComponent(
        preview.fileRef,
      )}`
    : "";

  return (
    <div className={styles.bucketCard} onClick={() => onOpen(bucket)}>
      <button
        className={styles.bucketUploadBtn}
        onClick={(e) => {
          e.stopPropagation();
          e.preventDefault();
          onUpload(bucket);
        }}
      >
        <span className={styles.plusIcon}>+</span>
      </button>

      {preview && !isVideo && { imageUrl }}

      {preview && isVideo && (
        <div
          className={styles.itemImg}
          style={{
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            fontSize: "50px",
            background: "#f3f2f1",
          }}
        >
          🎬
        </div>
      )}

      <div className={styles.bucketContent}>
        <h3>{bucket}</h3>

        <p className={styles.bucketCount}>{count} Dateien</p>
      </div>
    </div>
  );
};

export default BucketCard;
