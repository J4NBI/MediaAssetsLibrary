import * as React from "react";
import styles from "./MediaAssetsLib.module.scss";

/**
 * Props für die Bucket-Auswahl-Komponente.
 * Ermöglicht die Auswahl und Pflege von Bucket-/Ordnernamen.
 */
export interface IBucketDropdownProps {
  /** Verfügbare Bucket-Optionen. */
  options: string[];
  /** Aktuell ausgewählte Bucket-Namen. */
  selected: string[];
  /** Callback, der nach Änderungen an der Auswahl aufgerufen wird. */
  onChange: (values: string[]) => void;
}

/**
 * Dropdown-basierte Auswahlkomponente für Bucket-Namen.
 * Unterstützt die Eingabe eines neuen Werts und das Entfernen vorhandener Einträge.
 *
 * @param {IBucketDropdownProps} props - Konfiguration und Callback-Funktionen.
 * @returns {React.ReactElement} Die renderbare Bucket-Auswahl.
 */
const BucketDropdown = ({
  options,
  selected,
  onChange,
}: IBucketDropdownProps): React.ReactElement => {
  const [input, setInput] = React.useState("");
  const [open, setOpen] = React.useState(false);
  const ref = React.useRef<HTMLDivElement>(null);

  React.useEffect((): (() => void) => {
    const handleClickOutside = (e: MouseEvent): void => {
      if (ref.current && !ref.current.contains(e.target as Node)) {
        setOpen(false);
      }
    };

    document.addEventListener("mousedown", handleClickOutside);

    return (): void => {
      document.removeEventListener("mousedown", handleClickOutside);
    };
  }, []);

  const filtered = options.filter((o) =>
    o.toLowerCase().includes(input.toLowerCase()),
  );

  const addValue = (value: string): void => {
    const clean = value.trim();

    if (!clean) {
      return;
    }

    if (!selected.some((s) => s.toLowerCase() === clean.toLowerCase())) {
      onChange([...selected, clean]);
    }

    setInput("");
    setOpen(false);
  };

  return (
    <div ref={ref} style={{ position: "relative", width: "100%" }}>
      <input
        className={styles.tagInput}
        type="text"
        style={{ width: "100%" }}
        value={input}
        placeholder="Ordner wählen oder hinzufügen..."
        onFocus={() => setOpen(true)}
        onChange={(e) => {
          setInput(e.target.value);
          setOpen(true);
        }}
        onKeyDown={(e) => {
          if (e.key === "Enter") {
            e.preventDefault();
            addValue(input);
          }
        }}
      />

      {open && (
        <div
          style={{
            position: "absolute",
            top: "100%",
            left: 0,
            right: 0,
            background: "white",
            border: "1px solid #ccc",
            zIndex: 1000,
            maxHeight: "200px",
            overflow: "auto",
          }}
        >
          {input.trim() &&
            !options.some(
              (o) => o.toLowerCase() === input.trim().toLowerCase(),
            ) && (
              <div
                style={{
                  padding: "8px",
                  fontWeight: "bold",
                  cursor: "pointer",
                }}
                onClick={() => addValue(input)}
              >
                {`+ "${input}" hinzufügen`}
              </div>
            )}

          {filtered.map((o) => (
            <div
              key={o}
              style={{ padding: "8px", cursor: "pointer" }}
              onClick={() => addValue(o)}
            >
              {o}
            </div>
          ))}
        </div>
      )}

      <div
        style={{
          marginTop: 6,
          display: "flex",
          flexWrap: "wrap",
          gap: 6,
        }}
      >
        {selected.map((b) => (
          <span
            key={b}
            className={styles.bucketTag}
            onClick={() => onChange(selected.filter((x) => x !== b))}
          >
            {b} ✕
          </span>
        ))}
      </div>
    </div>
  );
};

export default BucketDropdown;
