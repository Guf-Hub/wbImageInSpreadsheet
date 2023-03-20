function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu(`🔽 МЕНЮ`)
    .addItem("🔄 Полчить фото", "getImg")
    .addToUi();
}

const CONFIG = {
  sheetName: "НАЗВАНИЕ ВАШЕГО ЛИСТ", // лист книги для вставки
  wbIdColumn: "ЧИСЛО", // номер столбца с nmId WB
  pastColumn: "ЧИСЛО", // номер столбца для вставки ссылок на фото
};

// main
function getImg() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    CONFIG.sheetName
  );
  img(sh, CONFIG.wbIdColumn, CONFIG.pastColumn);
}

/**
 * Создание и вставка картинки карточки WB
 * @param {SpreadsheetApp.Sheet} sheet лист книги
 * @param {number} wbIdColumn столбец с nmId
 * @param {number} pastColumn столбец для вставки ссылок на фото
 */
function img(sheet, wbIdColumn, pastColumn) {
  const data = sheet.getDataRange().getValues();
  let result = [["Фото"]];
  data.map((r, i) => {
    if (i > 0) {
      if (
        r[wbIdColumn] &&
        typeof r[wbIdColumn] === "number" &&
        r[wbIdColumn] > 0
      ) {
        result.push([`=IMAGE("${new GenerateImgUrl(r[wbIdColumn]).url()}")`]);
      } else {
        result.push([""]);
      }
    }
  });

  sheet.getRange(1, pastColumn, result.length, 1).setValues(result);
}

class GenerateImgUrl {
  constructor(nmId, photoSize, photoNumber) {
    if (typeof nmId !== "number" || nmId < 0) {
      throw new Error("Invalid nmId value");
    }
    this.nmId = nmId;
    this.size = photoSize || "c246x328";
    this.number = photoNumber || 1;
  }

  getUrlPart(id) {
    if (id >= 0 && id <= 143) return "//basket-01.wb.ru";
    if (id >= 144 && id <= 287) return "//basket-02.wb.ru";
    if (id >= 288 && id <= 431) return "//basket-03.wb.ru";
    if (id >= 432 && id <= 719) return "//basket-04.wb.ru";
    if (id >= 720 && id <= 1007) return "//basket-05.wb.ru";
    if (id >= 1008 && id <= 1061) return "//basket-06.wb.ru";
    if (id >= 1062 && id <= 1115) return "//basket-07.wb.ru";
    if (id >= 1116 && id <= 1169) return "//basket-08.wb.ru";
    if (id >= 1170 && id <= 1313) return "//basket-09.wb.ru";
    return "//basket-10.wb.ru";
  }

  url() {
    const vol = Math.floor(this.nmId / 100000);
    const part = Math.floor(this.nmId / 1000);
    return `https:${this.getUrlPart(vol)}/vol${vol}/part${part}/${this.nmId
      }/images/${this.size}/${this.number}.jpg`;
  }
}
