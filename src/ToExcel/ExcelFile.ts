import ExcelJS from "exceljs";
import * as fileSaver from "file-saver";
import { setStyleAndValue } from "./ExcelStyle";
import { setMerge } from "../common/method";
import { setImages } from "./ExcelImage";
import { setBorder } from "./ExcelBorder";
import { setDataValidations } from "./ExcelValidation";
import { setHiddenRowCol } from "./ExcelConfig";

export async function exportSheetExcel(luckysheetRef: any) {
  const luckysheet = luckysheetRef.getAllSheets();
  const workbook = new ExcelJS.Workbook();
  luckysheet.every(function (table: any) {
    if (table?.data?.length === 0) return true;
    const worksheet = workbook.addWorksheet(table.name);
    setStyleAndValue(table, worksheet);
    setMerge(table?.config?.merge, worksheet);
    setBorder(table, worksheet);
    setImages(table, worksheet, workbook);
    setDataValidations(table, worksheet);
    setHiddenRowCol(table, worksheet);
    return true;
  });
  const buffer = await workbook.xlsx.writeBuffer();
  const fileData = new Blob([buffer]);
  fileSaver.saveAs(fileData, `${luckysheetRef.getSheet().name}.xlsx`);
}

export async function exportSheetExcelCy(name: any, customData: any) {
  const workbook = new ExcelJS.Workbook();
  customData.every(function (table: any) {
    if (table?.data?.length === 0) return true;
    const worksheet = workbook.addWorksheet(table.name);
    setStyleAndValue(table, worksheet);
    setMerge(table?.config?.merge, worksheet);
    setBorder(table, worksheet);
    setImages(table, worksheet, workbook);
    setDataValidations(table, worksheet);
    setHiddenRowCol(table, worksheet);
    return true;
  });
  const buffer = await workbook.xlsx.writeBuffer();
  const fileData = new Blob([buffer]);
  fileSaver.saveAs(fileData, `${name}.xlsx`);
}

export async function exportSheetExcelByWorkbook(
  name: any,
  customData: any,
  workbook: any
) {
  // const workbook = new ExcelJS.Workbook();
  customData.every(function (table: any) {
    if (table?.data?.length === 0) return true;
    const worksheet = workbook.addWorksheet(table.name);
    setStyleAndValue(table, worksheet);
    setMerge(table?.config?.merge, worksheet);
    setBorder(table, worksheet);
    setImages(table, worksheet, workbook);
    setDataValidations(table, worksheet);
    setHiddenRowCol(table, worksheet);
    return true;
  });
  const buffer = await workbook.xlsx.writeBuffer();
  const fileData = new Blob([buffer]);
  fileSaver.saveAs(fileData, `${name}.xlsx`);
}
