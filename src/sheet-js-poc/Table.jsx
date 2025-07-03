import React from 'react';

/**
 * JSONTable
 * Renders a HTML table from an array of row‑objects.
 *
 * Props:
 *   data: Array<Object> — each object is a “row”, keys are column IDs.
 *   className?: string — optional CSS class for the <table>.
 */
export function JSONTable({ data, className }) {
  // 1) Collect all unique column keys in order of first appearance:
  const columns = React.useMemo(() => {
    const seen = new Set();
    const cols = [];
    data.forEach(row => {
      Object.keys(row).forEach(key => {
        if (!seen.has(key)) {
          seen.add(key);
          cols.push(key);
        }
      });
    });
    return cols;
  }, [data]);

  return (
    <table className={className} style={{ borderCollapse: 'collapse', width: '100%' }}>
      <thead>
        <tr>
          {columns.map(col => (
            <th
              key={col}
              style={{
                border: '1px solid #ccc',
                padding: '4px 8px',
                background: '#f5f5f5',
                textAlign: 'left',
              }}
            >
              {col}
            </th>
          ))}
        </tr>
      </thead>
      <tbody>
        {data.map((row, rIdx) => (
          <tr key={rIdx}>
            {columns.map(col => (
              <td
                key={col}
                style={{
                  border: '1px solid #ddd',
                  padding: '4px 8px',
                }}
              >
                {/* Render empty string if undefined/null */}
                {row[col] != null ? row[col].toString() : ''}
              </td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  );
}
