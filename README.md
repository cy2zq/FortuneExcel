# FortuneExcel

<p>
<a href="http://npmjs.com/package/@corbe30/fortune-excel" alt="fortuneExcel on npm">
<img src="https://img.shields.io/npm/v/@corbe30/fortune-excel" /></a>

<a href="http://npmjs.com/package/@corbe30/fortune-excel" alt="fortuneExcel downloads">
<img src="https://img.shields.io/npm/d18m/%40corbe30%2Ffortune-excel" /></a>
</p>

FortuneExcel is an import/export library for [FortuneSheet](https://github.com/ruilisi/fortune-sheet/). It only supports .xlsx files (not .xls).

## Features

Supports the following spreadsheet features in import/export:

- Cell style
- Cell border
- Cell format, such as number format, date, percentage, etc.
- Formula

## Usage

For best results, import and export a single sheet at a time. Although you can force FortuneExcel to handle multiple sheets, certain configurations may break.

1. Install the package:
```js
npm i @corbe30/fortune-excel
```

2. Add import/export toolbar item in fortune-sheet
> `<ImportHelper />` is a hidden component and only required when using `importToolBarItem()`.
```js
import { importToolBarItem, ImportHelper, exportToolBarItem } from "fortune-excel";

function App() {
  const workbookRef = useRef();
  const [key, setKey] = useState(0);
  const [sheets, setSheets] = useState(data);

  return (
    <>
      <ImportHelper setKey={setKey} setSheets={setSheets} sheetRef={workbookRef} />
      <Workbook
        key={key}
        data={sheets}
        ref={workbookRef}
        customToolbarItems={[exportToolBarItem(workbookRef), importToolBarItem()]}
      />
    </>
  );
}
```

## Authors and acknowledgment

- [@Corbe30](https://github.com/Corbe30)
- [@wbfsa](https://github.com/wbfsa)
- [@wpxp123456](https://github.com/wpxp123456)
- [@Dushusir](https://github.com/Dushusir)
- [@xxxDeveloper](https://github.com/xxxDeveloper)
- [@mengshukeji](https://github.com/mengshukeji)

## License

[MIT](http://opensource.org/licenses/MIT)
