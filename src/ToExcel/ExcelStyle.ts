import ExcelJS, { CellValue } from "exceljs";
import { fillConvert, fontConvert, alignmentConvert } from "./ExcelConvert.js";

const isTime = (d:string) => {
  return d === "hh:mm";
};

var setStyleAndValue = function (
  luckysheet: any,
  cellArr: any,
  worksheet: ExcelJS.Worksheet
) {
  if (!Array.isArray(cellArr)) return;

  cellArr.forEach(function (row, rowid) {
    const dbrow = worksheet.getRow(rowid + 1);
    //设置单元格行高,默认乘以1.2倍
    dbrow.height = luckysheet.getRowHeight([rowid])[rowid] / 1.2;
    row.every(function (cell: any, columnid: any) {
      if (!cell) return true;
      if (rowid == 0) {
        const dobCol = worksheet.getColumn(columnid + 1);
        //设置单元格列宽除以8
        dobCol.width = luckysheet.getColumnWidth([columnid])[columnid] / 8;
      }
      let fill = fillConvert(cell.bg);
      let font = fontConvert(
        cell.ff as string,
        cell.fc,
        cell.bl,
        cell.it,
        cell.fs,
        cell.cl,
        cell.un
      );
      let alignment = alignmentConvert(
        cell.vt,
        cell.ht,
        cell.tb && parseInt(cell.tb, 10),
        cell.tr && parseInt(cell.tr, 10)
      );
      let value: CellValue;

      var v: number | string | boolean | Date = "";
      var numFmt: string = undefined;
      // TODO: check and add support for currency, boolean, date format
      if (cell.ct && cell.ct.t == "inlineStr") {
        var s = cell.ct.s;
        s.forEach(function (val: any, num: any) {
          v += val.v;
        });
      } else if (cell.ct && cell.ct.t == "n") {
        v = +cell.v;
        if (cell.ct !== "General") numFmt = cell.ct.fa;
      } else if (cell.ct.t == "d") {
        const mockDate = isTime(cell.ct.fa) ? "2000-01-01 " : "";
        v = new Date(mockDate + cell.m);
        numFmt = cell.ct.fa;
      } else {
        v = cell.v as string;
      }
      if (cell.f) {
        value = {
          formula: cell.f.startsWith("=") ? cell.f.slice(1) : cell.f,
          result: v,
        };
      } else {
        value = v;
      }
      let target = worksheet.getCell(rowid + 1, columnid + 1);
      target.fill = fill;
      target.font = font;
      target.alignment = alignment;
      target.value = value;
      target.numFmt = numFmt;
      return true;
    });
  });
};

export { setStyleAndValue };
