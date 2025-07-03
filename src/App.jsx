import { useState } from "react";
import reactLogo from "./assets/react.svg";
import viteLogo from "/vite.svg";
import "./App.css";
import ExcelJsParser from "./excel-js-poc";
import SheetJsParser from "./sheet-js-poc";
import ExcelUploader from "./sheet-js-poc/ParseSheetV2";

function App() {
  const [count, setCount] = useState(0);

  return (
    <>
      {/* <ExcelJsParser/> */}
      {/* <SheetJsParser/> */}
      <ExcelUploader />
    </>
  );
}

export default App;
