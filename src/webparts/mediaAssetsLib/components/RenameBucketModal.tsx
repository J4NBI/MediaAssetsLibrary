import * as React from "react";
import styles from "./MediaAssetsLib.module.scss";

export interface IRenameBucketModalProps {
  isOpen: boolean;
  newBucketName: string;
  nameExists: boolean;
  onNameChange: (value: string) => void;
  onCancel: () => void;
  onSave: () => void;
}

export default function RenameBucketModal(
  props: IRenameBucketModalProps,
): React.ReactElement {
  if (!props.isOpen) return <></>;

  const disabled = !props.newBucketName.trim() || props.nameExists;

  return (
    <div className={styles.modalOverlay}>
      <div className={styles.renameModal}>
        <h3>Ordner umbenennen</h3>
        <input
          type="text"
          value={props.newBucketName}
          onChange={(e) => props.onNameChange(e.target.value)}
        />
        {props.nameExists && (
          <p className={styles.errorText}>
            Ein Ordner mit diesem Namen existiert bereits.
          </p>
        )}
        <div className={styles.modalActions}>
          <button className={styles.editBtn} onClick={props.onCancel}>
            Abbrechen
          </button>
          <button
            className={`${styles.downloadBtn} ${disabled ? styles.disabled : ""}`}
            disabled={disabled}
            onClick={props.onSave}
          >
            Speichern
          </button>
        </div>
      </div>
    </div>
  );
}
