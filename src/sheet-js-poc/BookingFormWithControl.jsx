import { useRef, useState, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";

export default function BookingFormWithControl() {
  const [businessData, setBusinessData] = useState([]);
  const [loadingControl, setLoadingControl] = useState(false);

  const bookingInput = useRef();
  const controlInput = useRef();
  const controlWorkerRef = useRef(null);

  // 1) Define columns, including two manual‐entry columns
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
    // new manual columns
    { header: "Pack Size", accessor: "pack_size" },
    { header: "Unit Cost", accessor: "unit_cost" },
    // computed columns
    { header: "Manufacturing Units", accessor: "manufacturing_units" },
    { header: "Total Value", accessor: "total_value" },
    { header: "Total Value after 1%", accessor: "total_value_after_1pct" },
  ];

  const formatDate = (serial) =>
    typeof serial === "number"
      ? XLSX.SSF.format("dd-mmm-yyyy", serial)
      : serial || "";

  const getCell = (sheet, addr) => sheet[addr]?.v || "";

  // 2) Parse Booking Form → seed businessData
  const handleBookingUpload = async (e) => {
    setBusinessData([]); // clear
    const file = e.target.files[0];
    if (!file) return;

    const wb = XLSX.read(await file.arrayBuffer(), {
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

    // Each row carries pack_size/unit_cost initially empty
    const rows = unitCells.reduce((arr, addr, idx) => {
      const unit = getCell(sht, addr);
      const rawLot = getCell(sht, lotCells[idx]);
      if (unit !== "") {
        let lot =
          rawLot === "A"
            ? 1
            : rawLot.toLowerCase() === "bulk"
            ? ++bulkCount
            : rawLot;
        arr.push({
          ...base,
          selling_unit: unit,
          lot,
          po_number: "",
          po_type: "",
          ship_mode: "",
          original_planned_po_delivery_date: "",
          pack_size: "",
          unit_cost: "",
          manufacturing_units: "",
          total_value: "",
          total_value_after_1pct: "",
        });
      }
      return arr;
    }, []);

    setBusinessData(rows);
  };

  // 3) Parse Control Tower → merge PO columns
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
        const { success, data: crs, error } = data;
        if (!success) {
          alert("Control parse failed: " + error);
          setLoadingControl(false);
          return;
        }
        setBusinessData((prev) =>
          prev.map((row, i) => {
            const cr = crs[i] || {};
            return {
              ...row,
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

  // 4) Reset flow
  const resetAndUploadBooking = () => {
    setBusinessData([]);
    setLoadingControl(false);
    bookingInput.current.value = "";
    controlInput.current.value = "";
    bookingInput.current.click();
  };

  // 5) Compute per-row derived fields
  useEffect(() => {
    setBusinessData((prev) =>
      prev.map((row) => {
        const ps = parseFloat(row.pack_size) || 0;
        const uc = parseFloat(row.unit_cost) || 0;
        const su = parseFloat(row.selling_unit) || 0;
        const manu = su * ps;
        const tot = su * uc;
        const aft = tot * 0.99;
        return {
          ...row,
          manufacturing_units: manu,
          total_value: tot.toFixed(2),
          total_value_after_1pct: aft.toFixed(2),
        };
      })
    );
  }, [
    businessData.map((r) => r.pack_size).join(),
    businessData.map((r) => r.unit_cost).join(),
  ]);

  // --- Render ---
  // if no rows yet, show one empty row
  const rowsToShow = businessData.length
    ? businessData
    : [columns.reduce((obj, c) => ({ ...obj, [c.accessor]: "" }), {})];
  console.log({ businessData });
  return (
    <div>
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

      <div style={{ overflow: "auto", marginTop: 20 }}>
        <table style={{ borderCollapse: "collapse", width: "100%" }}>
          <thead>
            <tr>
              {columns.map((c) => (
                <th
                  key={c.accessor}
                  style={{
                    border: "1px solid #000",
                    padding: 8,
                    backgroundColor: "#f0f0f0",
                  }}
                >
                  {c.header}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rowsToShow.map((row, ri) => (
              <tr key={ri}>
                {columns.map((c) => (
                  <td
                    key={c.accessor}
                    style={{ border: "1px solid #000", padding: 8 }}
                  >
                    {c.accessor === "pack_size" ||
                    c.accessor === "unit_cost" ? (
                      <input
                        type="number"
                        value={row[c.accessor]}
                        onChange={(e) => {
                          const v = e.target.value;
                          setBusinessData((prev) => {
                            const next = [...prev];
                            next[ri] = { ...next[ri], [c.accessor]: v };
                            return next;
                          });
                        }}
                        style={{ width: 60,background:"transparent", border:'0.5px solid gray' }}
                      />
                    ) : (
                      row[c.accessor]
                    )}
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
