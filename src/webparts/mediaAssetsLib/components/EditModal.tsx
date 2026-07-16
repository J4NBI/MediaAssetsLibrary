import * as React from "react";
import styles from "./MediaAssetsLib.module.scss";
import BucketDropdown from "./BucketDropdown";

interface IMediaItem {
  id: number;
  name: string;
  fileRef: string;

  category?: string;
  tags?: string[];
  bucket?: string[];

  format?: string;

  created?: string;
  createdBy?: string;
  dienst?: string;
}

interface IEditModalProps {
  isOpen: boolean;
  item?: IMediaItem;

  editName: string;
  editTags: string[];
  editCategory: string;
  editBucket: string[];

  categoryOptions: string[];
  bucketOptions: string[];

  onClose: () => void;
  onSave: () => void;
  onDelete: () => void;

  onNameChange: (value: string) => void;
  onCategoryChange: (value: string) => void;
  onBucketChange: (values: string[]) => void;
  onTagsChange: (values: string[]) => void;

  editDienst: string;

  dienstOptions: string[];

  onDienstChange: (value: string) => void;
}

const EditModal: React.FC<IEditModalProps> = ({
  isOpen,
  item,

  editName,
  editTags,
  editCategory,
  editBucket,
  editDienst,
  dienstOptions,

  onDienstChange,

  categoryOptions,
  bucketOptions,

  onClose,
  onSave,
  onDelete,

  onNameChange,
  onCategoryChange,
  onBucketChange,
  onTagsChange,
}) => {
  if (!isOpen || !item) {
    return null;
  }

  const fileUrl = `${window.location.origin}${item.fileRef}`;

  const fileType = item.fileRef?.split(".").pop()?.toLowerCase();

  const isImage = ["jpg", "jpeg", "png", "gif", "webp"].includes(
    fileType || "",
  );

  const isVideo = ["mp4", "mov", "webm"].includes(fileType || "");

  return (
    <div className={`${styles.modalOverlay} ${styles.modalOverlayEdit}`}>
      <div className={styles.modalBox}>
        <h3>Element bearbeiten</h3>

        {/* PREVIEW */}

        {isImage && <img src={fileUrl} className={styles.editPreview} />}

        {isVideo && (
          <video src={fileUrl} controls className={styles.editPreview} />
        )}

        {!isImage && !isVideo && (
          <div className={styles.editFallback}>📄 {item.name}</div>
        )}

        {/* NAME */}

        <input
          type="text"
          className={styles.tagInput}
          value={editName}
          onChange={(e) => onNameChange(e.target.value)}
          placeholder="Name"
        />

        {/* TAGS */}

        <div>
          <div className={styles.tagList}>
            {editTags.map((tag, index) => (
              <span key={index} className={`${styles.tag} ${styles.editTag}`}>
                {tag}

                <span
                  onClick={() => {
                    const newTags = [...editTags];

                    newTags.splice(index, 1);

                    onTagsChange(newTags);
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
            className={styles.tagInput}
            placeholder="Tag hinzufügen + Enter"
            onKeyDown={(e) => {
              if (e.key === "Enter") {
                e.preventDefault();

                const value = (e.target as HTMLInputElement).value.trim();

                if (!value) {
                  return;
                }

                onTagsChange([...editTags, value]);

                (e.target as HTMLInputElement).value = "";
              }
            }}
          />
        </div>

        {/* KATEGORIE */}

        <select
          value={editCategory}
          onChange={(e) => onCategoryChange(e.target.value)}
        >
          <option value="">Kategorie wählen</option>

          {categoryOptions.map((cat) => (
            <option key={cat} value={cat}>
              {cat}
            </option>
          ))}
        </select>

        <select
          value={editDienst}
          onChange={(e) => onDienstChange(e.target.value)}
        >
          <option value="">Dienst wählen</option>

          {dienstOptions.map((dienst) => (
            <option key={dienst} value={dienst}>
              {dienst}
            </option>
          ))}
        </select>

        {/* BUCKETS */}

        <BucketDropdown
          options={bucketOptions}
          selected={editBucket}
          onChange={onBucketChange}
        />

        {/* BUTTONS */}

        <div className={styles.modalActions}>
          <button onClick={onClose} className={styles.editBtn}>
            Abbrechen
          </button>

          <button onClick={onSave} className={styles.downloadBtn}>
            Speichern
          </button>

          <button onClick={onDelete} className={styles.btnDelete}>
            Löschen
          </button>
        </div>
      </div>
    </div>
  );
};

export default EditModal;
