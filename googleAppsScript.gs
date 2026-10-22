/**
 * ============================================================
 *  googleAppsScript.gs — приём заявок с лендинга «Нон-Стоп Дубна»
 * ============================================================
 *  КАК ПОДКЛЮЧИТЬ:
 *  1. Откройте таблицу Google Таблиц (или создайте новую).
 *  2. Расширения → Apps Script.
 *  3. Удалите весь код по умолчанию и вставьте содержимое этого файла.
 *  4. Нажмите «Сохранить» (иконка дискеты).
 *  5. Справа вверху кнопка «Развернуть» → «Новое развертывание».
 *     - Тип: «Веб-приложение»
 *     - Кто имеет доступ: «Все (анонимные пользователи)»
 *     - Нажмите «Развернуть».
 *     (При первом развертывании спросит права — согласитесь.)
 *  6. Скопируйте «URL веб-приложения» (вида
 *     https://script.google.com/macros/s/AKfyc.../exec)
 *     и вставьте его в index.html → LEAD_CONFIG → SHEETS_URL.
 *  7. Готово: все заявки будут появляться строками в Google Таблице.
 *
 *  ВАЖНО: новая заявка появляется в таблице почти сразу,
 *  но обновляйте страницу таблицы, чтобы увидеть её.
 *
 *  ПРАВИЛО «+»: телефон приходит строкой вида "+7 (926) 027-76-96".
 *  Если бы он попал в ячейку обычным числом, Таблицы приняли бы его за
 *  формулу и показали #ERROR!. Поэтому колонка «Телефон родителя»
 *  всегда приводится к формату «Текст» (@) — см. PHONE_COL.
 *  Если старые ячейки уже испорчены, запустите функцию fixPhoneColumn().
 * ============================================================
 */

var SHEET_NAME = "Заявки"; // имя вкладки, куда пишутся заявки
var PHONE_COL  = 6;        // номер колонки «Телефон родителя» (A=1 … F=6)

/* Шапка таблицы: 8 колонок. Колонка «Страница» убрана — она была лишней.
   Чтобы изменить состав колонок, правьте этот массив и запустите
   функцию rebuildHeader() (пересоберёт шапку и её цвет). */
var HEADERS = [
  "Дата", "Время",
  "Имя ребенка", "Фамилия ребенка",
  "Имя родителя", "Телефон родителя",
  "Возрастная группа", "Направление"
];
var HEADER_BG = "#D9EAD3"; // светло-зелёный фон шапки

/**
 * Обработчик POST-запроса от формы.
 * Сайт отправляет JSON вида:
 * { childName, childLastName, parentName, phone, group, ageGroup, page }
 *  - group    — направление (например «Хип-хоп»)
 *  - ageGroup — возрастная группа (например «Дети (6–7 лет)»)
 */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);

    // Если вкладки ещё нет — создаём
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    // Шапка всегда приводится к актуальному виду: 8 колонок + светло-зелёный фон.
    // Так исправляется и старая таблица, где была лишняя колонка «Страница».
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length)
         .setBackground(HEADER_BG)
         .setFontWeight("bold")
         .setVerticalAlignment("middle");

    // Если справа остались лишние колонки (например, от прежней «Страницы») —
    // очищаем их названия, чтобы шапка выглядела аккуратно.
    var extraCols = sheet.getLastColumn() - HEADERS.length;
    if (extraCols > 0) {
      sheet.getRange(1, HEADERS.length + 1, 1, extraCols)
           .clearContent()
           .setBackground(null);
    }

    // Колонка телефона ВСЕГДА в формате «Текст» (@).
    // Иначе значение вида "+7 (926) 027-76-96" Таблицы принимает за формулу
    // и в ячейке появляется #ERROR!.
    sheet.getRange(1, PHONE_COL, sheet.getMaxRows(), 1).setNumberFormat("@");

    var now = new Date();
    var tz = ss.getSpreadsheetTimeZone();

    // Телефон ВСЕГДА строка и всегда в текстовом формате: сайт присылает его
    // уже отформатированным как "+7 (926) 027-76-96", поэтому Google Таблицы
    // не могут превратить его в формулу и выдать #ERROR!.
    var phone = String(data.phone || "").replace(/^\s*\+\s*/, "+");

    sheet.appendRow([
      Utilities.formatDate(now, tz, "dd.MM.yyyy"),
      Utilities.formatDate(now, tz, "HH:mm"),
      String(data.childName     || ""),
      String(data.childLastName || ""),
      String(data.parentName    || ""),
      phone,
      String(data.ageGroup || ""),
      String(data.group      || "")
    ]);

    // Дублируем запись в колонку телефона как ТЕКСТ (appendRow уже отработал,
    // но номер строки фиксируем явно — так надёжнее при параллельных заявках).
    sheet.getRange(sheet.getLastRow(), PHONE_COL)
         .setNumberFormat("@")
         .setValue(phone);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Проверка связи: откройте URL веб-приложения в браузере —
 * увидите инструкцию вместо ошибки.
 */
function doGet() {
  var html = "Скрипт работает. Заявки записываются во вкладку «" + SHEET_NAME + "».";
  return HtmlService.createHtmlOutput("<b>" + html + "</b>");
}

/**
 *  ОДИН РАЗ ПОСЛЕ ОБНОВЛЕНИЯ СКРИПТА:
 *  чинит уже испорченные ячейки телефона (те, где вместо номера стоит #ERROR!)
 *  и переводит всю колонку «Телефон родителя» в формат «Текст».
 *
 *  Как запустить: в редакторе Apps Script выберите функцию fixPhoneColumn
 *  в списке функций сверху и нажмите «Запустить». Ответ появится в журнале
 *  выполнений (Выполнение → Журнал выполнений).
 */
function fixPhoneColumn() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error("Вкладка «" + SHEET_NAME + "» не найдена");
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    sheet.getRange(1, PHONE_COL, 1, 1).setNumberFormat("@");
    return "Заявок пока нет — колонка переведена в формат «Текст».";
  }

  var range = sheet.getRange(2, PHONE_COL, lastRow - 1, 1);
  var formulas = range.getFormulas();   // что реально лежит в ячейке
  var fixed = 0;

  for (var i = 0; i < formulas.length; i++) {
    var raw = String(formulas[i] || "").trim();
    if (!raw) continue;

    // Значение, которое Таблицы посчитала формулой: начинается с "+" или "=+"
    var cleaned = raw.replace(/^=+/, "").trim();
    if (cleaned.charAt(0) === "+") {
      cleaned = "+" + cleaned.substring(1).trim();
    }

    // Приводим к единому виду: без лишних пробелов после плюса
    cleaned = cleaned.replace(/^\+\s+/, "+");
    if (cleaned) {
      range.getCell(i + 1, 1).setNumberFormat("@").setValue(cleaned);
      fixed++;
    }
  }

  // Заголовок и вся колонка — текст
  sheet.getRange(1, PHONE_COL, sheet.getMaxRows(), 1).setNumberFormat("@");
  return "Исправлено ячеек с телефоном: " + fixed;
}

/**
 *  Пересобирает шапку таблицы: 8 колонок (без «Страницы»),
 *  светло-зелёный фон, жирный текст. Лишние колонки справа удаляются
 *  вместе с их данными.
 *
 *  Как запустить: выберите функцию rebuildHeader в списке функций сверху
 *  и нажмите «Запустить».
 */
function rebuildHeader() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // Удаляем всё, что правее нужных 8 колонок (прежняя «Страница» и её данные)
  var extraCols = sheet.getLastColumn() - HEADERS.length;
  if (extraCols > 0) {
    sheet.deleteColumns(HEADERS.length + 1, extraCols);
  }

  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.getRange(1, 1, 1, HEADERS.length)
       .setBackground(HEADER_BG)
       .setFontWeight("bold")
       .setVerticalAlignment("middle");

  // Телефон — текст
  sheet.getRange(1, PHONE_COL, sheet.getMaxRows(), 1).setNumberFormat("@");

  return "Шапка обновлена: " + HEADERS.length + " колонок, фон " + HEADER_BG + ".";
}