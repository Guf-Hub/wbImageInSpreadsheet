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
  sheetName: "НАЗВАНИЕ ЛИСТА", // название листа книги для вставки
  wbIdColumn: ЧИСЛО, // номер столбца, по порядку, с nmId Wildberries (формат столбца должен быть число)
  pastColumn: ЧИСЛО, // номер столбца, по порядку, для вставки ссылок на изображения
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
 
  id(id) {
    if (id <= 143) return "01";
    if (id <= 287) return "02";
    if (id <= 431) return "03";
    if (id <= 719) return "04";
    if (id <= 1007) return "05";
    if (id <= 1061) return "06";
    if (id <= 1115) return "07";
    if (id <= 1169) return "08";
    if (id <= 1313) return "09";
    if (id <= 1601) return "10";
    if (id <= 1655) return "11";
    if (id <= 1919) return "12";
    if (id <= 2045) return "13";
    if (id <= 2189) return "14";
    if (id <= 2405) return "15";
    if (id <= 2621) return "16";
    if (id <= 2837) return "17";
    if (id <= 3053) return "18";
    if (id <= 3269) return "19";
    if (id <= 3485) return "20";
    if (id <= 3701) return "21";
    return "22";
  }


  url() {
    const vol = ~~(this.nmId / 1e5);
    const part = ~~(this.nmId / 1e3);
    return `https:${this.id(vol)}/vol${vol}/part${part}/${
      this.nmId
    }/images/${this.size}/${this.number}.${this.format}`;
  }
}
