# FortuneExcel

FortuneExcel is an import/export library for [FortuneSheet](https://github.com/ruilisi/fortune-sheet/).
It only supports .xlsx format files (not .xls).

It is a fork of (now archived) [FortuneSheetExcel](https://github.com/zenmrp/FortuneSheetExcel).

## Features

Supports the following spreadsheet features in import/export:

- Cell style
- Cell border
- Cell format, such as number format, date, percentage, etc.
- Formula

## Usage

For best results, import and export a single sheet at a time. Although you can force FortuneExcel to handle multiple sheets, certain configurations may break.

### React frontend
`ImportHelper` is a hidden component and only required when using `importToolBarItem`.
```js
import { importToolBarItem, ImportHelper, exportToolBarItem } from "fortune-excel";

function App() {
  const workbookRef = useRef();
  const [key, setKey] = useState(0);
  const [sheets, setSheets] = useState(data);

  return (
    <>
      <ImportHelper setKey={setKey} setSheets={setSheets} sheetRef={workbookRef1} />
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

### Node backend
to be updated

## Authors and acknowledgment

- [@Corbe30](https://github.com/Corbe30)
- [@wbfsa](https://github.com/wbfsa)
- [@wpxp123456](https://github.com/wpxp123456)
- [@Dushusir](https://github.com/Dushusir)
- [@xxxDeveloper](https://github.com/xxxDeveloper)
- [@mengshukeji](https://github.com/mengshukeji)

## License

[MIT](http://opensource.org/licenses/MIT)
