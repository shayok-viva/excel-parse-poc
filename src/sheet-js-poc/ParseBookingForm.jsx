import React, { useRef, useState } from "react";
import * as XLSX from "xlsx";

// ParseBookingForm component handles file upload and displays data table
export default function ParseBookingForm() {
  const [excelData, setExcelData] = useState([]);
  const fileInput = useRef();

  // Define the column headers (excluding Selling Unit for now)
  const headers = [
    "Department",
    "Supplier",
    "Factory Name",
    "Season",
    "Phase",
    "Description",
    "Color",
    "Booking Ref",
    "Lot",
    "Booking form date",
    "Ship date in booking form",
  ];
  //for parsing excel dates
  const formatDate = (serial) => {
    if (typeof serial === "number") {
      return XLSX.SSF.format("dd-mmm-yyyy", serial);
    }
    return serial; // fallback for non-date
  };

  // Handle file upload and parsing
  const handleUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const data = await file.arrayBuffer(); // Read file as array buffer:contentReference[oaicite:4]{index=4}
    const workbook = XLSX.read(data, {
      type: "array",
      cellStyles: true,
      sheetStubs: true,
    });

    // Use the first sheet (or adjust if needed)
    const firstSheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheetName];

    // Helper to safely get cell value (or empty string if undefined)
    const getVal = (cellAddress) => {
      const cell = sheet[cellAddress];
      return cell ? cell.v : "";
    };

    // Extract fixed values from their cells
    const department = getVal("R12");
    const supplier = getVal("D26");
    const factory = getVal("D15"); // assuming D15 contains the full factory name
    const season = getVal("R13");
    const phase = getVal("R14");
    const description = getVal("D21");
    const color = getVal("D27");
    const bookingRef = getVal("D23");
    const bookingDate = formatDate(getVal("L14"));
    const shipDate = formatDate(getVal("J21"));

    // Collect all non-empty "Lot" values from I27 through I54
    const lots = [];
    for (let row = 27; row <= 54; row++) {
      const lotVal = getVal(`I${row}`);
      if (lotVal !== "") {
        lots.push(lotVal);
      }
    }

    // Build the data rows, repeating fixed fields for each lot
    const rows = lots.map((lot) => ({
      Department: department,
      Supplier: supplier,
      "Factory Name": factory,
      Season: season,
      Phase: phase,
      Description: description,
      Color: color,
      "Booking Ref": bookingRef,
      Lot: lot,
      "Booking form date": bookingDate,
      "Ship date in booking form": shipDate,
    }));

    setExcelData(rows);
  };
  return (
    <div>
      <div style={{ textAlign: "center", margin: "20px" }}>
        {/* Display selected file name or prompt */}
        <p>{excelData.length ? "Data loaded." : "Upload an Excel file"}</p>
        <button onClick={() => fileInput.current.click()}>
          {excelData.length ? "Load Another File" : "Select File"}
        </button>
        <input
          type="file"
          accept=".xlsx,.xls"
          ref={fileInput}
          style={{ display: "none" }}
          onChange={handleUpload}
        />
      </div>

      {excelData.length > 0 && (
        <div style={{ overflow: "auto", marginTop: "20px" }}>
          <table style={{ borderCollapse: "collapse", width: "100%" }}>
            <thead>
              <tr>
                {headers.map((h, idx) => (
                  <th
                    key={idx}
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
              {excelData.map((row, i) => (
                <tr key={i}>
                  {headers.map((h, j) => (
                    <td
                      key={j}
                      style={{ border: "1px solid #000", padding: 8 }}
                    >
                      {row[h] || ""}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
