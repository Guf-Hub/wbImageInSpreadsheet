# wbImageInSpreadsheet
<!-- HEADER START -->
<p style='text-align: center;'>
  <a href="https://openapi.wb.ru/"><img src="https://github.com/Guf-Hub/Wildberries/blob/main/src/wildberries.png"></a>
</p>
<hr />
<!-- HEADER END -->

Получение изображений в Google таблицу по коду товара (nmId) Wildberries.</br>

[![Donate](https://img.shields.io/badge/Donate-Yoomoney-green.svg)](https://yoomoney.ru/to/410019620244262)
![GitHub code size in bytes](https://img.shields.io/github/languages/code-size/Guf-Hub/WildberriesImageInGoogleSpreadsheet)
![javascript](https://img.shields.io/badge/lang-javascript-red)
![GAS](https://img.shields.io/badge/google-apps%20script-red)

[**Feedback**](https://t.me/nosaev_m)<br/>

## Использование:
* Полностью копируйте [скрипт](https://github.com/Guf-Hub/wbPhotoLink/blob/main/Code.js) в свою таблицу;
* Заполните столбец кодами **nmId**;
* Укажите свои данные в **CONFIG**:
```JavaScript
const CONFIG = {
  sheetName: "Товары", // название листа книги для вставки
  wbIdColumn: 1, // номер столбца с nmId Wildberries
  pastColumn: 2, // номер столбца для вставки ссылок на фото
};
```
* Сохраните изменения;
* Запустите скрипт > **🔽 МЕНЮ**;
* [Пройдите авторизацию](https://dzen.ru/media/excelifehack/kak-avtorizovat-skript-v-google-tablicah-61a943694333203e458eb600).

## Copyright & License

[MIT License](LICENSE)

Copyright (©) 2022 by [Mikhail Nosaev](https://github.com/Guf-Hub)

Настоящим предоставляется бесплатное разрешение любому лицу, получившему копию этого программного обеспечения и связанных с ним файлов документации («Программное обеспечение»), работать с Программным обеспечением без ограничений, включая, помимо прочего, права на использование, копирование, изменение, слияние. Публиковать, распространять, сублицензировать и/или продавать копии Программного обеспечения, а также разрешать лицам, которым предоставляется Программное обеспечение, делать это при соблюдении следующих условий:

Приведенное выше уведомление об авторских правах и это уведомление о разрешении должны быть включены во все копии или существенные части Программного обеспечения.

ПРОГРАММНОЕ ОБЕСПЕЧЕНИЕ ПРЕДОСТАВЛЯЕТСЯ «КАК ЕСТЬ», БЕЗ КАКИХ-ЛИБО ГАРАНТИЙ, ЯВНЫХ ИЛИ ПОДРАЗУМЕВАЕМЫХ, ВКЛЮЧАЯ, ПОМИМО ПРОЧЕГО, ГАРАНТИИ КОММЕРЧЕСКОЙ ПРИГОДНОСТИ, ПРИГОДНОСТИ ДЛЯ ОПРЕДЕЛЕННОЙ ЦЕЛИ И НЕНАРУШЕНИЯ ПРАВ. НИ ПРИ КАКИХ ОБСТОЯТЕЛЬСТВАХ АВТОРЫ ИЛИ ОБЛАДАТЕЛИ АВТОРСКИМ ПРАВОМ НЕ НЕСУТ ОТВЕТСТВЕННОСТИ ЗА ЛЮБЫЕ ПРЕТЕНЗИИ, УЩЕРБ ИЛИ ИНУЮ ОТВЕТСТВЕННОСТЬ, БУДУТ СВЯЗАННЫЕ С ДОГОВОРОМ, ДЕЛОМ ИЛИ ИНЫМ ОБРАЗОМ, ВОЗНИКАЮЩИЕ ИЗ ПРОГРАММНОГО ОБЕСПЕЧЕНИЯ ИЛИ ИСПОЛЬЗОВАНИЯ ИЛИ ИНЫХ СДЕЛОК В СВЯЗИ С ПРОГРАММНЫМ ОБЕСПЕЧЕНИЕМ ИЛИ ИСПОЛЬЗОВАНИЕМ ПРОГРАММНОГО ОБЕСПЕЧЕНИЯ.
