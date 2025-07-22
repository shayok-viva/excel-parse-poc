import { useRef, useState, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";

export default function BookingFormWithControl() {
  // Single source of truth for all row data
  const [businessData, setBusinessData] = useState([]);
  const [loadingControl, setLoadingControl] = useState(false);

  // Manual-entry inputs
  const [packSize, setPackSize] = useState("");
  const [unitCost, setUnitCost] = useState("");

  const bookingInput = useRef();
  const controlInput = useRef();
  const controlWorkerRef = useRef(null);

  const headers = [
    "Department",
    "Supplier",
    "Factory Name",
    "Season",
    "Phase",
    "Description",
    "Color",
    "Booking Ref",
    "Booking form date",
    "Ship date in booking form",
    "Selling Unit",
    "Lot",
    "PO number",
    "PO type",
    "Ship mode",
    "Original Planned PO delivery date",
    "Manufacturing Units",
    "Total Value",
    "Total Value after 1%",
  ];

  const formatDate = (serial) =>
    typeof serial === "number"
      ? XLSX.SSF.format("dd-mmm-yyyy", serial)
      : serial || "";

  const getCell = (sheet, addr) => sheet[addr]?.v || "";

  // --- 1) Parse Booking Form and seed businessData ---
  const handleBookingUpload = async (e) => {
    setBusinessData([]); // clear any prior data
    setPackSize("");
    setUnitCost("");

    const file = e.target.files[0];
    if (!file) return;

    const data = await file.arrayBuffer();
    const wb = XLSX.read(data, {
      type: "array",
      cellStyles: true,
      sheetStubs: true,
    });
    const sht = wb.Sheets[wb.SheetNames[0]];

    const bookingRef = getCell(sht, "D23");
    if (getCell(sht, "A23") !== "Ref" || !bookingRef) {
      alert("Booking ref not present at D23 or A23 mismatch");
      return;
    }

    // Base fields from booking form
    const base = {
      Department: getCell(sht, "R12"),
      Supplier: getCell(sht, "D26") || "",
      "Factory Name": getCell(sht, "D15"),
      Season: getCell(sht, "R13"),
      Phase: getCell(sht, "R14"),
      Description: getCell(sht, "D21"),
      Color: getCell(sht, "D27"),
      "Booking Ref": bookingRef,
      "Booking form date": formatDate(getCell(sht, "L14")),
      "Ship date in booking form": formatDate(getCell(sht, "J21")),
    };

    // Collect Seller Units & Lots
    const unitCells = ["J24", "K24", "L24", "M24"];
    const lotCells = ["J25", "K25", "L25", "M25"];
    let bulkCount = 0;
    const rows = unitCells.reduce((acc, addr, idx) => {
      const unit = getCell(sht, addr);
      const rawLot = getCell(sht, lotCells[idx]);
      if (unit !== "") {
        let lot;
        if (rawLot === "A") lot = 1;
        else if (rawLot.toLowerCase() === "bulk") lot = ++bulkCount;
        else lot = rawLot;
        acc.push({
          ...base,
          "Selling Unit": unit,
          Lot: lot,
          // leave control & computed fields empty for now
          "PO number": "",
          "PO type": "",
          "Ship mode": "",
          "Original Planned PO delivery date": "",
        });
      }
      return acc;
    }, []);

    // Seed businessData with booking-only rows
    setBusinessData(rows);
  };

  // --- 2) Parse Control Tower and merge into businessData ---
  const handleControlUpload = async (e) => {
    setLoadingControl(true);
    const file = e.target.files[0];
    if (!file || !businessData.length) {
      setLoadingControl(false);
      return;
    }

    if (!controlWorkerRef.current) {
      controlWorkerRef.current = new Worker(
        new URL("../web-workers/controlWorker.js", import.meta.url),
        { type: "module" }
      );
      controlWorkerRef.current.onmessage = (evt) => {
        const { success, data: controlRows, error } = evt.data;
        if (!success) {
          alert("Control parse failed: " + error);
          setLoadingControl(false);
          return;
        }

        // Merge controlRows (by index) into each businessData row
        setBusinessData((prev) =>
          prev.map((r, i) => {
            const cr = controlRows[i] || {};
            return {
              ...r,
              "PO number": cr["po_number"] || "",
              "PO type": cr["po_type"] || "",
              "Ship mode": cr["ship_mode"] || "",
              "Original Planned PO delivery date":
                cr["original_planned_po_delivery_date"] || "",
            };
          })
        );

        setLoadingControl(false);
        controlWorkerRef.current.terminate();
        controlWorkerRef.current = null;
      };
    }

    const buffer = await file.arrayBuffer();
    controlWorkerRef.current.postMessage({
      buffer,
      bookingRef: businessData[0]["Booking Ref"],
    });
  };

  // Reset flow and re-upload booking form
  const resetAndUploadBooking = () => {
    setBusinessData([]);
    setLoadingControl(false);
    setPackSize("");
    setUnitCost("");
    bookingInput.current.click();
    if (bookingInput.current) bookingInput.current.value = "";
  };

  const finalRows = useMemo(() => {
    const ps = parseFloat(packSize) || 0;
    const uc = parseFloat(unitCost) || 0;
    return businessData.map((row) => {
      const su = parseFloat(row["Selling Unit"]) || 0;
      const manu = su * ps;
      const tot = su * uc;
      const aft = tot * 0.99;

      // Return a brand‑new object with your computed fields included
      return {
        ...row,
        "Manufacturing Units": manu,
        "Total Value": tot.toFixed(2),
        "Total Value after 1%": aft.toFixed(2),
      };
    });
  }, [businessData, packSize, unitCost]);

  // --- 3) Render ---
  return (
    <div>
      {/* Upload Buttons */}
      <div style={{ textAlign: "center", margin: 20 }}>
        <button onClick={resetAndUploadBooking}>Upload Booking Form</button>
        <input
          ref={bookingInput}
          type="file"
          accept=".xlsx,.xls"
          style={{ display: "none" }}
          onChange={handleBookingUpload}
        />

        <button
          onClick={() => controlInput.current.click()}
          disabled={!businessData.length || loadingControl}
          style={{ marginLeft: 10 }}
        >
          {loadingControl
            ? "Parsing Control Tower..."
            : "Upload Control Tower File"}
        </button>
        {loadingControl && <span style={{ marginLeft: 10 }}>⏳</span>}
        <input
          ref={controlInput}
          type="file"
          accept=".xlsx,.xls"
          style={{ display: "none" }}
          onChange={handleControlUpload}
        />
      </div>

      {/* Manual inputs appear once both parsing steps done */}
      {businessData.length > 0 && !loadingControl && (
        <div style={{ textAlign: "center", margin: "1rem 0" }}>
          <label style={{ marginRight: 12 }}>
            Pack Size:&nbsp;
            <input
              type="number"
              value={packSize}
              onChange={(e) => setPackSize(e.target.value)}
              style={{ width: 80 }}
            />
          </label>
          <label>
            Unit Cost:&nbsp;
            <input
              type="number"
              value={unitCost}
              onChange={(e) => setUnitCost(e.target.value)}
              style={{ width: 80 }}
            />
          </label>
        </div>
      )}

      {/* Final table */}
      <div style={{ overflow: "auto", marginTop: 20 }}>
        <table style={{ borderCollapse: "collapse", width: "100%" }}>
          <thead>
            <tr>
              {headers.map((h) => (
                <th
                  key={h}
                  style={{
                    border: "1px solid #000",
                    padding: 8,
                    backgroundColor: "#f0f0f0",
                  }}
                >
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {finalRows.map((row, i) => (
              <tr key={i}>
                {headers.map((h) => (
                  <td key={h} style={{ border: "1px solid #000", padding: 8 }}>
                    {row[h] || ""}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
