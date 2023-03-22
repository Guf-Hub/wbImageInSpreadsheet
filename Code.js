/**
 * @OnlyCurrentDoc
 * @author Mikhail Nosaev <m.nosaev@gmail.com>
 * @see {@link https://t.me/nosaev_m Telegram} разработка Google таблиц и GAS скриптов
 * @license MIT
 */

function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu(`🔽 МЕНЮ`)
    .addItem("🔄 Полчить фото", "getImg")
    .addToUi();
}

/**
 * @type {Object.<string|number>}
 * @const
 */
const CONFIG = {
  sheetName: "НАЗВАНИЕТ", // название листа книги для вставки
  wbIdColumn: ЧИСЛО, // номер столбца с nmId Wildberries
  pastColumn: ЧИСЛО, // номер столбца для вставки ссылок на изображения
};

// main
function getImg() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    CONFIG.sheetName
  );
  img(sh, CONFIG.wbIdColumn - 1, CONFIG.pastColumn);
}

/**
 * Создание ссылки и вставка изображений
 * @param {SpreadsheetApp.Sheet} sheet лист книги
 * @param {number} wbIdColumn столбец с nmId Wildberries
 * @param {number} pastColumn столбец для вставки ссылок на изображения
 */
function img(sheet, wbIdColumn, pastColumn) {
  const data = sheet.getDataRange().getValues();
  const wbIdColumnIndex = wbIdColumn - 1;
  let result = [["Фото"]];
  data.map((r, i) => {
    if (i > 0) {
      if (
        r[wbIdColumnIndex] &&
        typeof r[wbIdColumnIndex] === "number" &&
        r[wbIdColumnIndex] > 0
      ) {
        result.push([
          `=IMAGE("${new GenerateImgUrl(r[wbIdColumnIndex]).url()}")`,
        ]);
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
    this.nmId = parseInt(nmId, 10);
    this.size = photoSize || "c246x328";
    this.number = photoNumber || 1;
  }

  getHost(id) {
    if (id >= 0 && id <= 143) return "//basket-01.wb.ru";
    if (id >= 144 && id <= 287) return "//basket-02.wb.ru";
    if (id >= 288 && id <= 431) return "//basket-03.wb.ru";
    if (id >= 432 && id <= 719) return "//basket-04.wb.ru";
    if (id >= 720 && id <= 1007) return "//basket-05.wb.ru";
    if (id >= 1008 && id <= 1061) return "//basket-06.wb.ru";
    if (id >= 1062 && id <= 1115) return "//basket-07.wb.ru";
    if (id >= 1116 && id <= 1169) return "//basket-08.wb.ru";
    if (id >= 1170 && id <= 1313) return "//basket-09.wb.ru";
    if (id >= 1314 && id <= 1601) return "//basket-10.wb.ru";
    if (id >= 1602 && id <= 1655) return "//basket-11.wb.ru";
    return "//basket-12.wb.ru";
  }

  url() {
    const vol = ~~(this.nmId / 1e5),
      part = ~~(this.nmId / 1e3);
    return `https:${this.getHost(vol)}/vol${vol}/part${part}/${
      this.nmId
    }/images/${this.size}/${this.number}.jpg`;
  }
}
