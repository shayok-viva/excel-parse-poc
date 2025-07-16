import { useState } from "react";
import "./App.css";
import ExcelUploader from "./sheet-js-poc/ParseSheetV2";
import ParseBookingForm from "./sheet-js-poc/ParseBookingForm";
import ParseBookingFormV2 from "./sheet-js-poc/ParseBookingFormV2";

function App() {
  const [count, setCount] = useState(0);

  return (
    <>
      {/* <ExcelUploader /> */}
      {/* <ParseBookingForm/> */}
      <ParseBookingFormV2/>
    </>
  );
}

export default App;
