import React from "react";
import { transformExcelToFortune } from "./main";

export default function ImportHelper(props) {
    const {
        setSheets, setKey, sheetRef
    } = props;
    return (
        <input
        type="file"
        id="ImportHelper"
        onChange={async (e) => {
          await transformExcelToFortune(e, setSheets, setKey, sheetRef);
        }}
        hidden
      />
    );
}