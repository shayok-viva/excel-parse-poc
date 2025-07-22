import { useRef, useState, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";

export default function BookingFormWithControl() {
  const [businessData, setBusinessData] = useState([]);
  const [loadingControl, setLoadingControl] = useState(false);
  const [packSize, setPackSize] = useState("");
  const [unitCost, setUnitCost] = useState("");

  const bookingInput = useRef();
  const controlInput = useRef();
  const controlWorkerRef = useRef(null);

  // 1) Define columns with display header and snake_case accessor
  const columns = [
    { header: "Department", accessor: "department" },
    { header: "Supplier", accessor: "supplier" },
    { header: "Factory Name", accessor: "factory_name" },
    { header: "Season", accessor: "season" },
    { header: "Phase", accessor: "phase" },
    { header: "Description", accessor: "description" },
    { header: "Color", accessor: "color" },
    { header: "Booking Ref", accessor: "booking_ref" },
    { header: "Booking Form Date", accessor: "booking_form_date" },
    {
      header: "Ship Date From Booking Form",
      accessor: "ship_date_from_booking_form",
    },
    { header: "Selling Unit", accessor: "selling_unit" },
    { header: "Lot", accessor: "lot" },
    { header: "PO Number", accessor: "po_number" },
    { header: "PO Type", accessor: "po_type" },
    { header: "Ship Mode", accessor: "ship_mode" },
    {
      header: "Original Planned PO Delivery Date",
      accessor: "original_planned_po_delivery_date",
    },
    { header: "Manufacturing Units", accessor: "manufacturing_units" },
    { header: "Total Value", accessor: "total_value" },
    { header: "Total Value After 1%", accessor: "total_value_after_1pct" },
  ];

  const formatDate = (serial) =>
    typeof serial === "number"
      ? XLSX.SSF.format("dd-mmm-yyyy", serial)
      : serial || "";

  const getCell = (sheet, addr) => sheet[addr]?.v || "";

  // 2) Booking form parser seeds businessData with snake_case keys
  const handleBookingUpload = async (e) => {
    setBusinessData([]);
    setPackSize("");
    setUnitCost("");

    const file = e.target.files[0];
    if (!file) return;

    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, {
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

    const base = {
      department: getCell(sht, "R12"),
      supplier: getCell(sht, "D26") || "",
      factory_name: getCell(sht, "D15"),
      season: getCell(sht, "R13"),
      phase: getCell(sht, "R14"),
      description: getCell(sht, "D21"),
      color: getCell(sht, "D27"),
      booking_ref: bookingRef,
      booking_form_date: formatDate(getCell(sht, "L14")),
      ship_date_from_booking_form: formatDate(getCell(sht, "J21")),
    };

    const unitCells = ["J24", "K24", "L24", "M24"];
    const lotCells = ["J25", "K25", "L25", "M25"];
    let bulkCount = 0;

    const rows = unitCells.reduce((arr, addr, idx) => {
      const unit = getCell(sht, addr);
      const rawLot = getCell(sht, lotCells[idx]);
      if (unit !== "") {
        let lot;
        if (rawLot === "A") lot = 1;
        else if (rawLot.toLowerCase() === "bulk") lot = ++bulkCount;
        else lot = rawLot;
        arr.push({
          ...base,
          selling_unit: unit,
          lot,
          po_number: "",
          po_type: "",
          ship_mode: "",
          original_planned_po_delivery_date: "",
        });
      }
      return arr;
    }, []);

    setBusinessData(rows);
  };

  // 3) Control parser merges into same businessData
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
      controlWorkerRef.current.onmessage = ({ data }) => {
        const { success, data: controlRows, error } = data;
        if (!success) {
          alert("Control parse failed: " + error);
          setLoadingControl(false);
          return;
        }
        setBusinessData((prev) =>
          prev.map((r, i) => {
            const cr = controlRows[i] || {};
            return {
              ...r,
              po_number: cr["po_number"] || "",
              po_type: cr["po_type"] || "",
              ship_mode: cr["ship_mode"] || "",
              original_planned_po_delivery_date:
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
      bookingRef: businessData[0].booking_ref,
    });
  };

  // 4) Reset and restart
  const resetAndUploadBooking = () => {
    // clear businessData so inputs reappear fresh
    setBusinessData([]);
    setLoadingControl(false);
    setPackSize("");
    setUnitCost("");
    // reset file input so onChange always fires
    bookingInput.current.value = "";
    controlInput.current.value = "";
    bookingInput.current.click();
  };

  // 5) Compute final rows with manual inputs
  const finalRows = useMemo(() => {
    const ps = parseFloat(packSize) || 0;
    const uc = parseFloat(unitCost) || 0;
    return businessData.map((row) => ({
      ...row,
      manufacturing_units: row.selling_unit * ps,
      total_value: (row.selling_unit * uc).toFixed(2),
      total_value_after_1pct: (row.selling_unit * uc * 0.99).toFixed(2),
    }));
  }, [businessData, packSize, unitCost]);

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

      {/* Manual inputs */}
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

      {/* Final Table */}
      <div style={{ overflow: "auto", marginTop: 20 }}>
        <table style={{ borderCollapse: "collapse", width: "100%" }}>
          <thead>
            <tr>
              {columns.map((col) => (
                <th
                  key={col.accessor}
                  style={{
                    border: "1px solid #000",
                    padding: 8,
                    backgroundColor: "#f0f0f0",
                  }}
                >
                  {col.header}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {finalRows
              ? finalRows.map((row, i) => (
                  <tr key={i}>
                    {columns.map((col) => (
                      <td
                        key={col.accessor}
                        style={{ border: "1px solid #000", padding: 8 }}
                      >
                        {row[col.accessor] ?? ""}
                      </td>
                    ))}
                  </tr>
                ))
              : columns.header.map((header,idx) => (
                  <tr key={idx}>
                    <td key={idx}>
                      {""}
                    </td>
                  </tr>
                ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
