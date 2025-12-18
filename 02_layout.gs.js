/**********************
 * Нормализация заголовков (лечит пробелы/варианты)
 **********************/
function normHeader_(h) {
  const t = String(h || '').trim().replace(/\s+/g, ' ');
  if (/^Покупка\b/i.test(t)) return 'Покупка (дней)';
  if (t === 'Покупка (дней с момента покупки)') return 'Покупка (дней)';
  if (t === 'Решение ') return 'Решение';
  if (t === 'Сообщение ') return 'Сообщение';
  return t;
}

/**********************
 * ЛЕЙАУТ + МИГРАЦИЯ СТРУКТУРЫ (FIX)
 **********************/
function ensureSheetLayout_(sheet) {
  const need = HEADERS_MAIN.length;

  // 0) гарантируем минимум колонок (не удаляем лишние до миграции)
  const maxCol0 = sheet.getMaxColumns();
  if (maxCol0 < need) sheet.insertColumnsAfter(maxCol0, need - maxCol0);

  const lastRow = sheet.getLastRow();
  const lastColExisting = Math.max(sheet.getLastColumn(), need);

  // 1) миграция по заголовкам при несовпадении
  if (lastRow >= 1) {
    const currentAll = sheet.getRange(1, 1, 1, lastColExisting).getDisplayValues()[0].map(normHeader_);
    const expected = HEADERS_MAIN.map(normHeader_);
    const currentFirst = currentAll.slice(0, need);
    const match = expected.every((h, i) => h === currentFirst[i]);

    if (!match && lastRow >= 2) {
      rebuildSheetByHeaders_(sheet, lastColExisting);
    }
  }

  // 2) теперь приводим кол-во колонок к need
  const maxCol = sheet.getMaxColumns();
  if (maxCol > need) sheet.deleteColumns(need + 1, maxCol - need);

  // 3) заголовки строго
  sheet.getRange(1, 1, 1, need).setValues([HEADERS_MAIN]).setFontWeight('bold');

  // 4) форматы
  const dataLast = getDataLastRow_(sheet);
  const rows = Math.max(1, dataLast - 1);

  try {
    sheet.getRange(2, COL.DT, rows, 1).setNumberFormat('dd.MM.yyyy HH:mm');
    sheet.getRange(2, COL.DEADLINE, rows, 1).setNumberFormat('dd.MM.yyyy HH:mm');
    sheet.getRange(2, COL.PURCHASE_DAYS, rows, 1).setNumberFormat('0');
    sheet.getRange(2, COL.WARRANTY, rows, 1).setNumberFormat('@');
    sheet.getRange(2, COL.DECISION, rows, 1).setNumberFormat('@');
    sheet.getRange(2, COL.MESSAGE, rows, 1).setNumberFormat('@');
  } catch (e) {}

  // ✅ выпадающий список решений (колонка L)
  try {
    const dv = SpreadsheetApp.newDataValidation()
      .requireValueInList(DECISIONS, true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(2, COL.DECISION, rows, 1).setDataValidation(dv);
  } catch (e) {}

  // 5) фильтр
  try {
    if (!sheet.getFilter()) sheet.getRange(1, 1, 1, need).createFilter();
  } catch (e) {}

  // 6) скрываем служебную
  try { sheet.hideColumns(COL.FOREIGN_BRAND); } catch (e) {}
}

/**
 * Пересборка листа по заголовкам
 */
function rebuildSheetByHeaders_(sheet, lastColExisting) {
  const need = HEADERS_MAIN.length;
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(lastColExisting || sheet.getLastColumn(), need);

  const old = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const oldHeaders = (old[0] || []).map(normHeader_);

  const map = new Map();
  for (let c = 0; c < oldHeaders.length; c++) {
    const name = oldHeaders[c];
    if (name && !map.has(name)) map.set(name, c + 1);
  }

  const out = new Array(lastRow);
  out[0] = HEADERS_MAIN.slice();

  for (let r = 1; r < lastRow; r++) {
    const row = new Array(need).fill('');
    for (let i = 0; i < need; i++) {
      const h = normHeader_(HEADERS_MAIN[i]);
      const src = map.get(h);
      row[i] = src ? old[r][src - 1] : '';
    }
    out[r] = row;
  }

  sheet.clear({ contentsOnly: true });
  sheet.getRange(1, 1, lastRow, need).setValues(out);
}

/**********************
 * Последняя строка данных
 **********************/
function getDataLastRowByColumn_(sheet, colIdx) {
  const last = sheet.getLastRow();
  if (last < 2) return 1;

  const vals = sheet.getRange(1, colIdx, last, 1).getValues().flat();
  for (let i = vals.length - 1; i >= 1; i--) {
    if (String(vals[i] || '').trim() !== '') return i + 1;
  }
  return 1;
}

function getDataLastRow_(sheet) {
  return getDataLastRowByColumn_(sheet, COL.CLAIM_ID);
}

/**********************
 * СЕРВИС
 **********************/
function getOrCreateSheet_(name) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  ensureSheetLayout_(sh);
  return sh;
}

function safeParseDate_(s) {
  if (!s) return null;
  const d = (s instanceof Date) ? s : new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

/**********************
 * Автосортировка по дате создания (A)
 **********************/
function autoSortByDate_(sheet) {
  const lastRow = getDataLastRow_(sheet);
  if (lastRow < 3) return;

  const filter = sheet.getFilter();
  if (!filter) return;

  const range = sheet.getRange(2, 1, lastRow - 1, HEADERS_MAIN.length);
  range.sort({ column: COL.DT, ascending: true });
}
