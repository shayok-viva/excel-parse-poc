import React, { useRef, useState } from "react";
import * as XLSX from "xlsx";

/**
 * ExcelRenderer
 * Renders a single worksheet with exact layout, merges, and styles.
 */
function ExcelRenderer({ sheet }) {
  // 1) Build matrix of values
  const matrix = React.useMemo(
    () =>
      XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" }),
    [sheet]
  );

  // 2) Prepare merges map
  const merges = sheet["!merges"] || [];
  const mergeMap = React.useMemo(() => {
    const map = {};
    merges.forEach(({ s, e }) => {
      const key = `${s.r},${s.c}`;
      map[key] = { rowSpan: e.r - s.r + 1, colSpan: e.c - s.c + 1 };
      for (let r = s.r; r <= e.r; r++) {
        for (let c = s.c; c <= e.c; c++) {
          const rk = `${r},${c}`;
          if (rk !== key) map[rk] = "skip";
        }
      }
    });
    return map;
  }, [merges]);

  // 3) Style map
  const styleMap = React.useMemo(() => {
    const map = {};
    Object.entries(sheet).forEach(([addr, cell]) => {
      if (addr.startsWith("!") || !cell.s) return;
      const { r, c } = XLSX.utils.decode_cell(addr);
      map[`${r},${c}`] = cell.s;
    });
    return map;
  }, [sheet]);

  // 4) Dimensions
  const colWidths = (sheet["!cols"] || []).map((c) => c?.wpx || 100);
  const rowHeights = (sheet["!rows"] || []).map((r) => r?.hpx || 30);

  // 5) Style helper
  const toCSS = React.useCallback(
    (r, c) => {
      const s = styleMap[`${r},${c}`];
      if (!s) return {};
      const css = { boxSizing: "border-box" };
      if (s.fill?.fgColor?.rgb)
        css.backgroundColor = `#${s.fill.fgColor.rgb.slice(-6)}`;
      if (s.font) {
        if (s.font.sz) css.fontSize = `${s.font.sz}pt`;
        if (s.font.color?.rgb) css.color = `#${s.font.color.rgb.slice(-6)}`;
        if (s.font.bold) css.fontWeight = "bold";
      }
      if (s.alignment) {
        if (s.alignment.horizontal) css.textAlign = s.alignment.horizontal;
        if (s.alignment.wrapText) css.whiteSpace = "pre-wrap";
      }
      return css;
    },
    [styleMap]
  );

  // Render table
  return (
    <div style={{ overflow: "auto" }}>
      <table style={{ borderCollapse: "collapse", tableLayout: "fixed" }}>
        <tbody>
          {matrix.map((row, r) => (
            <tr key={r} style={{ height: `${rowHeights[r]}px` }}>
              {row.map((val, c) => {
                const mk = `${r},${c}`;
                if (mergeMap[mk] === "skip") return null;
                const props = {};
                const m = mergeMap[mk];
                if (m) {
                  props.rowSpan = m.rowSpan;
                  props.colSpan = m.colSpan;
                }
                const cellStyle = toCSS(r, c);
                return (
                  <td
                    key={c}
                    {...props}
                    style={{
                      border: "1px solid #000",
                      width: `${colWidths[c]}px`,
                      padding: 4,
                      ...cellStyle,
                    }}
                  >
                    {val}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

/**
 * ExcelUploader
 * Handles file upload, parses Excel, and renders all sheets dynamically.
 */
export default function ExcelUploader() {
  const [sheets, setSheets] = useState([]);
  const [names, setNames] = useState([]);
  const [active, setActive] = useState(0);
  const [file, setFile] = useState(null);
  const fileInput = useRef();
  const handleUpload = async (e) => {
    const file = e.target.files[0];
    setFile(file);
    if (!file) return;

    const data = await file.arrayBuffer();
    const wb = XLSX.read(data, {
      type: "array",
      cellStyles: true,
      sheetStubs: true,
    });

    setNames(wb.SheetNames);
    setSheets(wb.SheetNames.map((name) => wb.Sheets[name]));
    setActive(0);
  };

  return (
    <div>
      <div
        style={{
          display: "flex",
          flexDirection: "column",
          alignItems: "center",
        }}
      >
        <div
          style={{
            display: "flex",
            flexDirection: "column",
            alignItems: "center",
            justifyContent: "center",
          }}
        >
          <p>
            {file ? `Selected file: ${file.name}` : "Upload an Excel file"}
          </p>
          <button type="button" onClick={() => fileInput.current.click()}>
            {sheets.length ? " Change" : "Upload"}
          </button>
        </div>
        <input
          ref={fileInput}
          style={{ display: "none" }}
          type="file"
          accept=".xlsx,.xls"
          onChange={handleUpload}
        />
      </div>

      {sheets.length > 0 && (
        <div>
          <div style={{ margin: "8px 0" }}>
            <h2>Sheet Names:</h2>
            {names.map((n, i) => (
              <button
                key={n}
                onClick={() => setActive(i)}
                style={{ marginRight: 4 }}
              >
                {n}
              </button>
            ))}
          </div>
          <ExcelRenderer sheet={sheets[active]} />
        </div>
      )}
    </div>
  );
}
