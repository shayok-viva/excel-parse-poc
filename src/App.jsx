import { useState } from "react";
import "./App.css";
import ExcelUploader from "./sheet-js-poc/ParseSheetV2";

function App() {
  const [count, setCount] = useState(0);

  return (
    <>
      <ExcelUploader />
    </>
  );
}

export default App;
