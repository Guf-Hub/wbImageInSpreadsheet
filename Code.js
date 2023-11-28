/**
 * @OnlyCurrentDoc
 * @author Mikhail Nosaev <m.nosaev@gmail.com>
 * @see {@link https://t.me/nosaev_m Telegram} разработка Google таблиц и GAS скриптов
 * @license MIT
 */

function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu(`🔽 МЕНЮ`)
    .addItem("🔄 Получить фото", "getImg")
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

function getImg() {
  const { sheetName, wbIdColumn, pastColumn } = CONFIG;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  img_(sheet, wbIdColumn, pastColumn);
}

/**
 * Создание ссылки и вставка изображений
 * @param {SpreadsheetApp.Sheet} sheet лист книги
 * @param {number} wbIdColumn столбец с nmId Wildberries
 * @param {number} pastColumn столбец для вставки ссылок на изображения
 */
function img_(sheet, wbIdColumn, pastColumn) {
  const data = sheet.getDataRange().getValues();
  const wbIdColumnIndex = wbIdColumn - 1;

  const result = [["Фото"]];
  data.map((r, i) => {
    console.log(r);
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
  constructor(nmId, photoSize, photoNumber, format) {
    if (typeof nmId !== "number" || nmId < 0) {
      throw new Error("Invalid nmId value");
    }
    this.nmId = parseInt(nmId, 10);
    this.size = photoSize || "big"; //"c246x328";
    this.number = photoNumber || 1;
    this.format = format || "webp"; //"jpg";
  }

  getHost(id) {
    const urlParts = [
      { range: [0, 143], url: "//basket-01.wb.ru" },
      { range: [144, 287], url: "//basket-02.wb.ru" },
      { range: [288, 431], url: "//basket-03.wb.ru" },
      { range: [432, 719], url: "//basket-04.wb.ru" },
      { range: [720, 1007], url: "//basket-05.wb.ru" },
      { range: [1008, 1061], url: "//basket-06.wb.ru" },
      { range: [1062, 1115], url: "//basket-07.wb.ru" },
      { range: [1116, 1169], url: "//basket-08.wb.ru" },
      { range: [1170, 1313], url: "//basket-09.wb.ru" },
      { range: [1314, 1601], url: "//basket-10.wb.ru" },
      { range: [1602, 1655], url: "//basket-11.wb.ru" },
      { range: [1656, 1919], url: "//basket-12.wb.ru" },
      { range: [1920, 2045], url: "//basket-13.wb.ru" },
      { range: [2046, Infinity], url: "//basket-14.wb.ru" },
    ];

    const { url } = urlParts.find(
      ({ range }) => id >= range[0] && id <= range[1]
    );
    return url;
  }

  url() {
    const vol = ~~(this.nmId / 1e5),
      part = ~~(this.nmId / 1e3);
    return `https:${this.getHost(vol)}/vol${vol}/part${part}/${
      this.nmId
    }/images/${this.size}/${this.number}.${this.format}`;
  }
}
