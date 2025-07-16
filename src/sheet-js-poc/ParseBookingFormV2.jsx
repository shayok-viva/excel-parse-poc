import React, { useRef, useState } from "react";
import * as XLSX from "xlsx";

// React component to upload an Excel file and render specified fields
export default function ParseBookingFormV2() {
  const [excelData, setExcelData] = useState([]);
  const fileInput = useRef();

  // Define table headers (excluding future fields)
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
  ];

  // Format Excel serial date to readable string
  const formatDate = (serial) => {
    if (typeof serial === "number") {
      return XLSX.SSF.format("dd-mmm-yyyy hh:MM", serial);
    }
    return serial || "";
  };

  // Safely get the cell value
  const getCell = (sheet, addr) => {
    const cell = sheet[addr];
    return cell ? cell.v : "";
  };

  // Handle file selection and parsing
  const handleUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const data = await file.arrayBuffer();
    const wb = XLSX.read(data, {
      type: "array",
      cellStyles: true,
      sheetStubs: true,
    });

    const sheet = wb.Sheets[wb.SheetNames[0]];

    // Extract fixed fields
    const base = {
      Department: getCell(sheet, "R12"),
      Supplier: getCell(sheet, "D26")||'',
      "Factory Name": getCell(sheet, "D15"),
      Season: getCell(sheet, "R13"),
      Phase: getCell(sheet, "R14"),
      Description: getCell(sheet, "D21"),
      Color: getCell(sheet, "D27"),
      "Booking Ref": getCell(sheet, "D23"),
      "Booking form date": formatDate(getCell(sheet, "L14")),
      "Ship date in booking form": formatDate(getCell(sheet, "J21")),
    };

    // Coordinates for Selling Unit and Type
    const unitCells = ["J24", "K24", "L24", "M24"];
    const typeCells = ["J25", "K25", "L25", "M25"];

    // Build data rows based on number of selling unit entries
    const rows = unitCells.reduce((acc, addr, idx) => {
      const unitVal = getCell(sheet, addr);
      if (unitVal !== "") {
        acc.push({
          ...base,
          "Selling Unit": unitVal,
          Lot: getCell(sheet, typeCells[idx]),
        });
      }
      return acc;
    }, []);

    setExcelData(rows);
  };

  return (
    <div>
      <div style={{ textAlign: "center", margin: 20 }}>
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
        <div style={{ overflow: "auto", marginTop: 20 }}>
          <table style={{ borderCollapse: "collapse", width: "100%" }}>
            <thead>
              <tr>
                {headers.map((h) => (
                  <th
                    key={h}
                    style={{ border: "1px solid #000", padding: 8, backgroundColor: "#f0f0f0" }}
                  >
                    {h}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {excelData.map((row, i) => (
                <tr key={i}>
                  {headers.map((h) => (
                    <td key={h} style={{ border: "1px solid #000", padding: 8 }}>
                      {row[h]}
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
