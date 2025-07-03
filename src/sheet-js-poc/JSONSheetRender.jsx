import React from 'react';
import * as XLSX from 'xlsx';

/**
 * JSONSheetRenderer.jsx
 * Dynamically renders any Excel sheet with exact layout, merges, cell styles, row heights, and column widths.
 *
 * Props:
 *   sheet: XLSX.WorkSheet — a SheetJS worksheet (parsed with cellStyles: true, sheetStubs: true).
 *   className?: string — optional CSS class for the table container.
 */
export default function JSONSheetRenderer({ sheet, className }) {
  // 1) Extract cell matrix: preserve blank cells
  const matrix = React.useMemo(
    () => XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: '' }),
    [sheet]
  );

  // 2) Build merge lookup map
  const merges = sheet['!merges'] || [];
  const mergeMap = React.useMemo(() => {
    const map = {};
    merges.forEach(({ s, e }) => {
      const startKey = `${s.r},${s.c}`;
      map[startKey] = { rowSpan: e.r - s.r + 1, colSpan: e.c - s.c + 1 };
      for (let rr = s.r; rr <= e.r; rr++) {
        for (let cc = s.c; cc <= e.c; cc++) {
          const key = `${rr},${cc}`;
          if (key !== startKey) map[key] = 'skip';
        }
      }
    });
    return map;
  }, [merges]);

  // 3) Build style map from cell style objects
  const styleMap = React.useMemo(() => {
    const map = {};
    Object.entries(sheet).forEach(([cell, obj]) => {
      if (cell.startsWith('!') || !obj.s) return;
      const { r, c } = XLSX.utils.decode_cell(cell);
      map[`${r},${c}`] = obj.s;
    });
    return map;
  }, [sheet]);

  // 4) Column widths (pixels) and row heights (pixels)
  const colWidths = (sheet['!cols'] || []).map(c => (c && c.wpx ? c.wpx : 80));
  const rowHeights = (sheet['!rows'] || []).map(r => (r && r.hpx ? r.hpx : 24));

  // Helper: convert SheetJS style object to inline CSS
  const styleFor = React.useCallback((r, c) => {
    const s = styleMap[`${r},${c}`];
    if (!s) return {};
    const css = {
      boxSizing: 'border-box',
    };
    // Fill color
    if (s.fill?.fgColor?.rgb) {
      css.backgroundColor = `#${s.fill.fgColor.rgb.slice(-6)}`;
    }
    // Font settings
    if (s.font) {
      if (s.font.sz) css.fontSize = `${s.font.sz}pt`;
      if (s.font.color?.rgb) css.color = `#${s.font.color.rgb.slice(-6)}`;
      if (s.font.bold) css.fontWeight = 'bold';
      if (s.font.italic) css.fontStyle = 'italic';
      if (s.font.name) css.fontFamily = s.font.name;
      if (s.font.underline) css.textDecoration = 'underline';
    }
    // Alignment
    if (s.alignment) {
      if (s.alignment.horizontal) css.textAlign = s.alignment.horizontal;
      if (s.alignment.vertical) css.verticalAlign = s.alignment.vertical;
      if (s.alignment.wrapText) css.whiteSpace = 'pre-wrap';
    }
    return css;
  }, [styleMap]);

  return (
    <div className={className} style={{ overflow: 'auto' }}>
      <table style={{ borderCollapse: 'collapse', tableLayout: 'fixed' }}>
        <tbody>
          {matrix.map((row, rIdx) => (
            <tr key={rIdx} style={{ height: `${rowHeights[rIdx] || 24}px` }}>
              {row.map((cellValue, cIdx) => {
                const key = `${rIdx},${cIdx}`;
                const m = mergeMap[key];
                if (m === 'skip') return null;
                const props = {};
                if (m) {
                  if (m.rowSpan > 1) props.rowSpan = m.rowSpan;
                  if (m.colSpan > 1) props.colSpan = m.colSpan;
                }
                const cellStyle = styleFor(rIdx, cIdx);
                const width = colWidths[cIdx] || 80;
                return (
                  <td
                    key={cIdx}
                    {...props}
                    style={{
                      border: '1px solid #000',
                      padding: '2px 4px',
                      width: `${width}px`,
                      ...cellStyle,
                    }}
                  >
                    {cellValue}
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
