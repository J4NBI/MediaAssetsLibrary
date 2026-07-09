import * as React from "react";

interface IFilterBarProps {
  filterCategory?: string;
  filterDienst?: string;
  filterCreator?: string;
  filterFormat?: string;
  filterYear?: number;
  filterMonth?: number;

  categoryOptions: string[];
  dienstOptions: string[];
  creatorOptions: string[];
  formatOptions: string[];
  yearOptions: number[];

  onChange: (values: {
    filterCategory?: string;
    filterDienst?: string;
    filterCreator?: string;
    filterFormat?: string;
    filterYear?: number;
    filterMonth?: number;
  }) => void;
}

const FilterBar: React.FC<IFilterBarProps> = ({
  filterCategory,
  filterDienst,
  filterCreator,
  filterFormat,
  filterYear,
  filterMonth,
  categoryOptions,
  dienstOptions,
  creatorOptions,
  formatOptions,
  yearOptions,
  onChange,
}) => {
  return (
    <>
      <select
        value={filterCategory || ""}
        onChange={(e) =>
          onChange({
            filterCategory: e.target.value || undefined,
          })
        }
      >
        <option value="">Kategorie</option>

        {categoryOptions.map((cat) => (
          <option key={cat} value={cat}>
            {cat}
          </option>
        ))}
      </select>

      <select
        value={filterDienst || ""}
        onChange={(e) =>
          onChange({
            filterDienst: e.target.value || undefined,
          })
        }
      >
        <option value="">Dienst</option>

        {dienstOptions.map((dienst) => (
          <option key={dienst} value={dienst}>
            {dienst}
          </option>
        ))}
      </select>

      <select
        value={filterCreator || ""}
        onChange={(e) =>
          onChange({
            filterCreator: e.target.value || undefined,
          })
        }
      >
        <option value="">Ersteller</option>

        {creatorOptions.map((creator) => (
          <option key={creator} value={creator}>
            {creator}
          </option>
        ))}
      </select>

      <select
        value={filterFormat || ""}
        onChange={(e) =>
          onChange({
            filterFormat: e.target.value || undefined,
          })
        }
      >
        <option value="">Format</option>

        {formatOptions.map((format) => (
          <option key={format} value={format}>
            {format}
          </option>
        ))}
      </select>

      <select
        value={filterYear || ""}
        onChange={(e) =>
          onChange({
            filterYear: e.target.value ? Number(e.target.value) : undefined,
          })
        }
      >
        <option value="">Jahr</option>

        {yearOptions.map((year) => (
          <option key={year} value={year}>
            {year}
          </option>
        ))}
      </select>

      <select
        value={filterMonth || ""}
        onChange={(e) =>
          onChange({
            filterMonth: e.target.value ? Number(e.target.value) : undefined,
          })
        }
      >
        <option value="">Monat</option>

        {[
          { v: 1, n: "Jan" },
          { v: 2, n: "Feb" },
          { v: 3, n: "Mär" },
          { v: 4, n: "Apr" },
          { v: 5, n: "Mai" },
          { v: 6, n: "Jun" },
          { v: 7, n: "Jul" },
          { v: 8, n: "Aug" },
          { v: 9, n: "Sep" },
          { v: 10, n: "Okt" },
          { v: 11, n: "Nov" },
          { v: 12, n: "Dez" },
        ].map((m) => (
          <option key={m.v} value={m.v}>
            {m.n}
          </option>
        ))}
      </select>
    </>
  );
};

export default FilterBar;
