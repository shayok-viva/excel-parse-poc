import { useState } from "react";
import "./App.css";
// import ExcelUploader from "./sheet-js-poc/ParseSheetV2";
// import ParseBookingForm from "./sheet-js-poc/ParseBookingForm";
// import ParseBookingFormV2 from "./sheet-js-poc/ParseBookingFormV2";
import BookingFormWithControl from "./sheet-js-poc/BookingFormWithControl";

function App() {
  const [count, setCount] = useState(0);

  return (
    <>
      {/* <ExcelUploader /> */}
      {/* <ParseBookingForm/> */}
      {/* <ParseBookingFormV2/> */}
      <img src="/css-logo.png" />
      <BookingFormWithControl />
    </>
  );
}

export default App;
