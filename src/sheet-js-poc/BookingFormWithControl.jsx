import { useRef, useState, useEffect } from "react";
import * as XLSX from "xlsx";

export default function BookingFormWithControl() {
  const [bookingRows, setBookingRows] = useState([]);
  const [controlRows, setControlRows] = useState([]);
  const [mergedRows, setMergedRows] = useState([]);
  const [loadingControl, setLoadingControl] = useState(false);

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
  ];

  const formatDate = (serial) =>
    typeof serial === "number"
      ? XLSX.SSF.format("dd-mmm-yyyy", serial)
      : serial || "";

  const getCell = (sheet, addr) => sheet[addr]?.v || "";

  // --- Booking Form Handler (Main Thread) ---
  const handleBookingUpload = async (e) => {
    setBookingRows([]);
    setMergedRows([]);
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
        acc.push({ ...base, "Selling Unit": unit, Lot: lot });
      }
      return acc;
    }, []);

    setBookingRows(rows);
  };

  // --- Control Tower Handler (Web Worker) ---
  const handleControlUpload = async (e) => {
    setControlRows([]);
    setMergedRows([]);
    setLoadingControl(true);
    const file = e.target.files[0];
    if (!file || !bookingRows.length) {
      setLoadingControl(false);
      return;
    }

    if (!controlWorkerRef.current) {
      controlWorkerRef.current = new Worker(
        new URL("../web-workers/controlWorker.js", import.meta.url),
        { type: "module" }
      );
      controlWorkerRef.current.onmessage = (evt) => {
        const { success, data: rows, error } = evt.data;
        if (success) {
          setControlRows(rows);
        } else {
          alert("Control parse failed: " + error);
        }
        setLoadingControl(false);
        controlWorkerRef.current.terminate();
        controlWorkerRef.current = null;
      };
    }
    const buffer = await file.arrayBuffer();
    controlWorkerRef.current.postMessage({
      buffer,
      bookingRef: bookingRows[0]["Booking Ref"],
    });
  };
  function handleUploadBookingformButtonClick() {
    setBookingRows([]);
    setControlRows([]);
    setMergedRows([]);
    bookingInput.current.click();
  }
  // Merge when both sets update
  useEffect(() => {
    if (!bookingRows.length) return;
    if (!controlRows.length) {
      setMergedRows(
        bookingRows.map((br) => ({
          ...br,
          "PO number": "",
          "PO type": "",
          "Ship mode": "",
          "Original Planned PO delivery date": "",
        }))
      );
      return;
    }
    const merged = bookingRows.map((br, idx) => ({
      ...br,
      "PO number": controlRows[idx].po_number || "",
      "PO type": controlRows[idx].po_type || "",
      "Ship mode": controlRows[idx].ship_mode || "",
      "Original Planned PO delivery date":
        controlRows[idx].original_planned_po_delivery_date || "",
    }));
    setMergedRows(merged);
  }, [bookingRows, controlRows]);

  return (
    <div>
      <div style={{ textAlign: "center", margin: 20 }}>
        <button onClick={handleUploadBookingformButtonClick}>
          Upload Booking Form
        </button>
        <input
          ref={bookingInput}
          type="file"
          accept=".xlsx,.xls"
          style={{ display: "none" }}
          onChange={handleBookingUpload}
        />

        <button
          onClick={() => controlInput.current.click()}
          disabled={!bookingRows.length || loadingControl}
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
          {mergedRows.length > 0 ? (
            <tbody>
              {mergedRows.map((row, i) => (
                <tr key={i}>
                  {headers.map((h) => (
                    <td
                      key={h}
                      style={{
                        width: "fit-content",
                        border: "1px solid #000",
                        padding: 8,
                      }}
                    >
                      {row[h] || ""}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          ) : (
            <tbody>
              <tr>
                {headers.map((h) => (
                  <td
                    key={h}
                    style={{
                      height: "100px",
                      border: "1px solid #000",
                      padding: 8,
                    }}
                  >
                    {""}
                  </td>
                ))}
              </tr>
            </tbody>
          )}
        </table>
      </div>
    </div>
  );
}
