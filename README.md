# wbPhotoLink

Получение изображений в Google таблицу по коду товара (nmId) Wildberries.</br>

[![Donate](https://img.shields.io/badge/Donate-Yoomoney-green.svg)](https://yoomoney.ru/to/410019620244262)
![GitHub code size in bytes](https://img.shields.io/github/languages/code-size/Guf-Hub/wbPhotoLink)
![javascript](https://img.shields.io/badge/lang-javascript-red)
![GAS](https://img.shields.io/badge/google-apps%20script-red)

### Использование:
* Полностью копируйте [скрипт](https://github.com/Guf-Hub/wbPhotoLink/blob/main/Code.js) в свою таблицу;
* Заполните столбец кодами **nmId**;
* Укажите свои данные в CONFIG, пример:
```JavaScript
const CONFIG = {
  sheetName: "Товары", // название листа книги для вставки
  wbIdColumn: 1, // номер столбца с nmId Wildberries
  pastColumn: 2, // номер столбца для вставки ссылок на фото
};
```
* Запустите скрипт > **🔽 МЕНЮ**;
* [Пройдите авторизацию](https://dzen.ru/media/excelifehack/kak-avtorizovat-skript-v-google-tablicah-61a943694333203e458eb600).
