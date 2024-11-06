import ExportIcon from "./icons/ExportIcon";
import ImportIcon from "./icons/ImportIcon";
import { transformFortuneToExcel } from "./main";

export const exportToolBarItem = (sheetRef) => {
  return {
    key: "export",
    tooltip: "export .xlsl",
    icon: ExportIcon(),
    onClick: async (e) => {
      await transformFortuneToExcel(sheetRef.current);
    },
  };
};

export const importToolBarItem = () => {
  return {
    key: "import",
    tooltip: "import .xlsl",
    icon: ImportIcon(),
    onClick: (e) => {
      document.getElementById("ImportHelper")?.click();
    },
  };
};