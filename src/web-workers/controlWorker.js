// src/web-workers/controlWorker.js
import * as XLSX from 'xlsx';

self.onmessage = async (e) => {
  try {
    const { buffer, bookingRef } = e.data;
    const refLower = bookingRef?.toString().trim().toLowerCase();
    if (!refLower) throw new Error('No bookingRef provided');

    // Read workbook
    const wb = XLSX.read(buffer, { type: 'array' });
    const sht = wb.Sheets[wb.SheetNames[0]];
    const sheetRef = sht['!ref'];
    if (!sheetRef) throw new Error('Sheet reference not found');

    // Decode range
    const range = XLSX.utils.decode_range(sheetRef);

    // Header row is 3rd row (0-based index = range.s.r + 2)
    const headerRowIndex = range.s.r + 2;
    const headerMap = {};
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const addr = XLSX.utils.encode_cell({ r: headerRowIndex, c: C });
      const cell = sht[addr];
      const key = cell && cell.v ? cell.v.toString().trim().toLowerCase() : '';
      if (key) headerMap[key] = C;
    }

    // Ensure necessary columns exist
    const idxManu = headerMap['manuref'];
    if (idxManu === undefined) throw new Error('Manuref column not found');

    // Map other columns, fallback to known indices
    const idxA  = headerMap['po number'] ?? 0;
    const idxAC = headerMap['po type'] ?? 28;
    const idxAR = headerMap['ship mode'] ?? 43;
    const idxAW = headerMap['original planned po delivery date'] ?? 48;

    // Collect matches
    const matches = [];
    for (let R = headerRowIndex + 1; R <= range.e.r; ++R) {
      const manuAddr = XLSX.utils.encode_cell({ r: R, c: idxManu });
      const manuCell = sht[manuAddr];
      const manuVal = manuCell && manuCell.v ? manuCell.v.toString().trim().toLowerCase() : '';
      if (manuVal === refLower) {
        const getValAt = (col) => {
          const a = XLSX.utils.encode_cell({ r: R, c: col });
          const c = sht[a];
          return c && c.v != null ? c.v : '';
        };
        matches.push({
          'PO number': getValAt(idxA),
          'PO type': getValAt(idxAC),
          'Ship mode': getValAt(idxAR),
          'Original Planned PO delivery date': typeof getValAt(idxAW) === 'number'
            ? XLSX.SSF.format('dd-mmm-yyyy', getValAt(idxAW))
            : getValAt(idxAW)
        });
      }
    }

    self.postMessage({ success: true, data: matches });
  } catch (err) {
    self.postMessage({ success: false, error: err.message });
  }
};
