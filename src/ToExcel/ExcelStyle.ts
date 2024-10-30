import ExcelJS from "exceljs";
import { WorkbookInstance } from "@fortune-sheet/react";
import { CellMatrix } from "@fortune-sheet/core";
import { fillConvert, fontConvert, alignmentConvert } from "./ExcelConvert.js";

var setStyleAndValue = function (
  luckysheet: WorkbookInstance,
  cellArr: CellMatrix,
  worksheet: ExcelJS.Worksheet
) {
  if (!Array.isArray(cellArr)) return;

  cellArr.forEach(function (row, rowid) {
    const dbrow = worksheet.getRow(rowid + 1);
    //设置单元格行高,默认乘以1.2倍
    dbrow.height = luckysheet.getRowHeight([rowid])[rowid] * 1.2;
    row.every(function (cell, columnid) {
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
      let value;

      var v = "";
      if (cell.ct && cell.ct.t == "inlineStr") {
        var s = cell.ct.s;
        s.forEach(function (val: any, num: any) {
          v += val.v;
        });
      } else {
        v = cell.v as string;
      }
      if (cell.f) {
        value = { formula: cell.f, result: v };
      } else {
        value = v;
      }
      let target = worksheet.getCell(rowid + 1, columnid + 1);
      target.fill = fill;
      target.font = font;
      target.alignment = alignment;
      target.value = value;
      return true;
    });
  });
};

export { setStyleAndValue };
