import ExportIcon from "./icons/ExportIcon.js";
import ImportIcon from "./icons/ImportIcon.js";
import { transformFortuneToExcel } from "./Transform.js";

export const exportToolBarItem = (sheetRef:any) => {
  return {
    key: "export",
    tooltip: "export .xlsl",
    icon: ExportIcon(),
    onClick: async (e:any) => {
      await transformFortuneToExcel(sheetRef.current);
    },
  };
};

export const importToolBarItem = () => {
  return {
    key: "import",
    tooltip: "import .xlsl",
    icon: ImportIcon(),
    onClick: (e:any) => {
      document.getElementById("ImportHelper")?.click();
    },
  };
};