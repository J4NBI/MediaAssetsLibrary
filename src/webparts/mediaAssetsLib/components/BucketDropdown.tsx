import * as React from "react";

export interface IBucketDropdownProps {
  options: string[];
  selected: string[];
  onChange: (values: string[]) => void;
}

/********************* BUCKET DROPDOWN *****************
 * Custom Multi-Select Dropdown:
 * - Suche
 * - Mehrfachauswahl
 * - neue Buckets hinzufügen
 ******************************************************/

const BucketDropdown: React.FC<IBucketDropdownProps> = ({
  options,
  selected,
  onChange,
}) => {
  /**************** LOCAL STATE ****************/
  const [input, setInput] = React.useState("");
  const [open, setOpen] = React.useState(false);
  const ref = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) {
        setOpen(false);
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  /**************** FILTER LOGIK ****************/
  const filtered = options.filter((o) =>
    o.toLowerCase().includes(input.toLowerCase()),
  );

  /**************** ADD BUCKET ****************/
  const addValue = (value: string): void => {
    const clean = value.trim();
    if (!clean) return;

    if (!selected.some((s) => s.toLowerCase() === clean.toLowerCase())) {
      onChange([...selected, clean]);
    }

    setInput("");
    setOpen(false);
  };

  /**************** RENDER ****************/

  return (
    <div ref={ref} style={{ position: "relative", width: "100%" }}>
      {/* INPUT */}
      <label>Ordner</label>
      <input
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

      {/* DROPDOWN */}
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
          {/* NEU */}
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
                + "{input}" hinzufügen
              </div>
            )}

          {/* LISTE */}
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

      {/* AUSWAHL */}
      <div style={{ marginTop: 6, display: "flex", flexWrap: "wrap", gap: 6 }}>
        {selected.map((b) => (
          <span
            key={b}
            style={{
              background: "#ecdd04",
              color: "black",
              padding: "4px 8px",
              borderRadius: "0px",
              cursor: "pointer",
            }}
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
