import * as React from "react";
import styles from "./MediaAssetsLib.module.scss";

/**
 * Props für das Modal zur Umbenennung eines Buckets.
 */
export interface IRenameBucketModalProps {
  /** Gibt an, ob das Modal sichtbar ist. */
  isOpen: boolean;
  /** Aktueller Eingabewert für den neuen Bucket-Namen. */
  newBucketName: string;
  /** Gibt an, ob der eingegebene Name bereits existiert. */
  nameExists: boolean;
  /** Callback bei Änderungen im Eingabefeld. */
  onNameChange: (value: string) => void;
  /** Callback für den Abbrechen-Button. */
  onCancel: () => void;
  /** Callback für den Speichern-Button. */
  onSave: () => void;
}

/**
 * Modal zur Umbenennung eines Buckets.
 *
 * @param {IRenameBucketModalProps} props - Zustand und Callback-Funktionen des Modals.
 * @returns {React.ReactElement} Das renderbare Modal oder ein leeres Fragment.
 */
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
