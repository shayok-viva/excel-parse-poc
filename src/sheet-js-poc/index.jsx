import { useState } from "react";
import { read, utils } from "xlsx";
import { JSONTable } from "./Table";
import JSONSheetRenderer from "./JSONSheetRender";

const SheetJsParser = () => {
  const [data, setData] = useState();
  const [sheet, setSheet] = useState();
  const [file, setFile] = useState(null);
  const handleFileChange = (e) => {
    const file = e.target.files[0];
    if (file) {
      setFile(file);
      const reader = new FileReader();
      reader.onload = (event) => {
        const binaryString = event.target.result;
        const workbook = read(binaryString, { type: "binary" });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = utils.sheet_to_json(worksheet);
        setData(jsonData);
      };
      reader.readAsBinaryString(file);
      setSheet(
        read(binaryString, { type: "binary" }).Sheets[workbook.SheetNames[0]]
      );
      const jsonData = utils.sheet_to_json(sheet, { header: 1 });
      setData(jsonData);
    }
    return;
  };
  return (
    <div>
      <input type="file" accept=".xlsx, .xls" onChange={handleFileChange} />
      {data && <JSONTable data={data} />}
      <h2>SheetJS Parser</h2>
      <p>
        This component allows you to upload an Excel file, parse its contents,
        and download the parsed data as a new Excel file.
      </p>
      <p>It uses the SheetJS library to read and write Excel files.</p>
      <p>
        To use this component, simply upload an Excel file, and it will display
        the parsed data in JSON format. You can then download the parsed data as
        a new Excel file.
      </p>
      <p>Make sure to have the necessary dependencies installed:</p>
      <pre>npm install xlsx file-saver</pre>
      <p>Enjoy parsing your Excel files!</p>
    </div>
  );
};

export default SheetJsParser;
