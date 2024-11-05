# FortuneExcel

FortuneExcel is an import/export library for [FortuneSheet](https://github.com/ruilisi/fortune-sheet/).
It only supports .xlsx format files (not .xls).

It is a fork of (now archived) FortuneSheetExcel.

## Features

Supports the following spreadsheet features in import/export:

- Cell style
- Cell border
- Cell format, such as number format, date, percentage, etc.
- Formula

## Usage

> NOTE: to be modified as a plugin for FortuneSheet

For best results, import and export a single sheet at a time. Although you can force FortuneExcel to handle multiple sheets, certain configurations may break.

```js
import { transformExcelToFortune } from "FortuneSheetExcel";

// e.g. got a file input change event
const xls = await e.target.files[0].arrayBuffer();
const fsh = await transformExcelToFortune(xls);
setData(fsh.sheets); // use this as the Workbook data
```

Interactively in a node repl:

```js
f = await (
  await import("node:fs/promises")
).readFile("/home/val/Downloads/Silkscreen.xlsx");
console.log(
  (
    await (
      await import("FortuneSheetExcel")
    ).FortuneExcel.transformExcelToFortune(f)
  ).toJsonString()
);
// in dev: console.log((await (await import("./dist/main.js")).FortuneExcel.transformExcelToFortune(f)).toJsonString())
```

## TODO

1. Add plugin support in FortuneSheet
2. Add plugin support in this project
3. publish FortuneExcel as an npm package

## Authors and acknowledgment

- [@Corbe30](https://github.com/Corbe30)
- [@wbfsa](https://github.com/wbfsa)
- [@wpxp123456](https://github.com/wpxp123456)
- [@Dushusir](https://github.com/Dushusir)
- [@xxxDeveloper](https://github.com/xxxDeveloper)
- [@mengshukeji](https://github.com/mengshukeji)

## License

[MIT](http://opensource.org/licenses/MIT)
