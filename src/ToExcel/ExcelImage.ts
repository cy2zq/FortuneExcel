//获取图片在单元格的位置
var getImagePosition = function (num: number, arr: number[]) {
    let index = 0;
    let minIndex;
    let maxIndex;
    for (let i = 0; i < arr.length; i++) {
      if (num < arr[i]) {
        index = i;
        break;
      }
    }
  
    if (index == 0) {
      minIndex = 0;
      maxIndex = 1;
      return Math.abs((num - 0) / (arr[maxIndex] - arr[minIndex])) + index;
    } else if (index == arr.length - 1) {
      minIndex = arr.length - 2;
      maxIndex = arr.length - 1;
    } else {
      minIndex = index - 1;
      maxIndex = index;
    }
    let min = arr[minIndex];
    let max = arr[maxIndex];
    let radio = Math.abs((num - min) / (max - min)) + index;
    return radio;
  };
  
  var setImages = function (table: any, worksheet: any, workbook: any) {
    const localTable = { ...table };
    let {
      images,
      visibledatacolumn, //所有行的位置
      visibledatarow, //所有列的位置
    } = localTable;
    if (typeof images != "object") return;
    for (let key in images) {
      // 通过 base64  将图像添加到工作簿
      const myBase64Image = images[key].src;
      //开始行 开始列 结束行 结束列
      const item = images[key];
      const imageId = workbook.addImage({
        base64: myBase64Image,
        extension: "png",
      });
  
      const col_st = getImagePosition(item.default.left, visibledatacolumn);
      const row_st = getImagePosition(item.default.top, visibledatarow);
  
      //模式1，图片左侧与luckysheet位置一样，像素比例保持不变，但是，右侧位置可能与原图所在单元格不一致
      worksheet.addImage(imageId, {
        tl: { col: col_st, row: row_st },
        ext: { width: item.default.width, height: item.default.height },
      });
      //模式2,图片四个角位置没有变动，但是图片像素比例可能和原图不一样
      // const w_ed = item.default.left+item.default.width;
      // const h_ed = item.default.top+item.default.height;
      // const col_ed = getImagePosition(w_ed,visibledatacolumn);
      // const row_ed = getImagePosition(h_ed,visibledatarow);
      // worksheet.addImage(imageId, {
      //   tl: { col: col_st, row: row_st},
      //   br: { col: col_ed, row: row_ed},
      // });
    }
  };
  
  export { setImages };
  