import React from "react";
import { Sheet } from "@fortune-sheet/core";
import { Workbook } from "@fortune-sheet/react";
import "@fortune-sheet/react/dist/index.css";
import { exportToolBarItem, importToolBarItem } from "../ToolbarItem.js";
import { ImportHelper } from "../ImportHelper.jsx";

export const Page: React.FC = () => {
  const [key, setKey] = React.useState<number>(0);
  const [sheets, setSheets] = React.useState<Sheet[]>([{ name: "Sheet1" }]);
  const sheetRef: any = React.useRef(null);

  return (
    <div
      style={{
        display: "flex",
        flexDirection: "column",
        width: "100%",
        height: "100vh",
      }}
    >
      <ImportHelper setKey={setKey} setSheets={setSheets} sheetRef={sheetRef} />
      <Workbook
        key={key}
        data={sheets}
        ref={sheetRef}
        customToolbarItems={[importToolBarItem(), exportToolBarItem(sheetRef)]}
      />
    </div>
  );
};
