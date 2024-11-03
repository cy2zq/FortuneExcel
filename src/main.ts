import { FortuneFile } from "./ToFortuneSheet/FortuneFile.js";
import type { FortuneFileBase } from "./ToFortuneSheet/FortuneBase.ts";
import { HandleZip } from "./HandleZip.js";
import { WorkbookInstance } from "@fortune-sheet/react";
import { exportSheetExcel } from './ToExcel/ExcelFile.js';

export const transformExcelToFortune = async (
  excelFile: File,
): Promise<FortuneFileBase> => {
  const files = await new HandleZip(excelFile).unzipFile();
  const fortuneFile = new FortuneFile(files, excelFile.name);
  fortuneFile.Parse();
  return fortuneFile.serialize();
};

export const transformFortuneToExcel = async (
  luckysheetRef: WorkbookInstance,
  name = "file"
) => {
  await exportSheetExcel(luckysheetRef, name);
};