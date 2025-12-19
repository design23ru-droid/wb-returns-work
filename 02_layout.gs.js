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
function ensureSheetLayout_(sheet, opts) {
  opts = opts || {};
  const allowRebuild = !!opts.allowRebuild;

  const need = HEADERS_MAIN.length;

  // 0) гарантируем минимум колонок (не удаляем лишние до миграции)
  const maxCol0 = sheet.getMaxColumns();
  if (maxCol0 < need) sheet.insertColumnsAfter(maxCol0, need - maxCol0);

  const lastRow = sheet.getLastRow();
  const lastColExisting = Math.max(sheet.getLastColumn(), need);

  // 1) проверка/миграция по заголовкам при несовпадении
  if (lastRow >= 1) {
    const currentAll = sheet.getRange(1, 1, 1, lastColExisting).getDisplayValues()[0].map(normHeader_);
    const expected = HEADERS_MAIN.map(normHeader_);
    const currentFirst = currentAll.slice(0, need);
    const match = expected.every((h, i) => h === currentFirst[i]);

    // если данные уже есть — любые перестановки только по кнопке "Миграция"
    if (!match && lastRow >= 2) {
      if (allowRebuild) {
        rebuildSheetByHeaders_(sheet, lastColExisting);
      } else {
        throw new Error(
          'Структура листа не совпадает с текущей схемой колонок.\n\n' +
          'Чтобы не перемещать колонки неожиданно, автоматическая пересборка отключена.\n' +
          'Открой меню "WB · Возвраты" → "Администрирование" → "Миграция".'
        );
      }
    }
  }

  // 2) приводим кол-во колонок к need
  //    (удаление лишних — только при миграции или когда лист пустой)
  const maxCol = sheet.getMaxColumns();
  if ((allowRebuild || lastRow <= 1) && maxCol > need) {
    sheet.deleteColumns(need + 1, maxCol - need);
  }

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
    if (COL.REASON) sheet.getRange(2, COL.REASON, rows, 1).setNumberFormat('@');
    sheet.getRange(2, COL.DECISION, rows, 1).setNumberFormat('@');
    sheet.getRange(2, COL.MESSAGE, rows, 1).setNumberFormat('@');
  } catch (e) {}

  // ✅ выпадающий список решений (колонка "Решение")
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

  // 6) Скрытие служебной колонки — ТОЛЬКО по заголовку, чтобы не скрыть "Сообщение"
  try {
    const scanCols = Math.max(sheet.getLastColumn(), need);
    const hdr = sheet.getRange(1, 1, 1, scanCols).getDisplayValues()[0].map(normHeader_);

    let fbCol = 0;
    let msgCol = 0;

    for (let i = 0; i < hdr.length; i++) {
      const h = hdr[i];
      if (!msgCol && h === 'Сообщение') msgCol = i + 1;
      if (!fbCol && h === '_foreignBrand') fbCol = i + 1;
    }

    if (msgCol) sheet.showColumns(msgCol);
    if (fbCol) sheet.hideColumns(fbCol);
  } catch (e) {}
}

/**
 * Пересборка листа по заголовкам (перенос данных в новый порядок HEADERS_MAIN)
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

  // ✅ Алиасы для переименований (чтобы данные не терялись при смене названий колонок)
  const aliasPairs = [
    ['Создано', 'Дата создания'],
    ['Артикул', 'NM ID'],
    ['Тип возврата', 'Возврат'],
    ['Тип возврата', 'Возврат (правило: Заявка / 14 дней)'],
    ['Давность', 'Покупка (дней)'],
    ['Давность', 'Покупка (дней с момента покупки)'],
    ['Кол-во', 'Кол-во (кол-во отзывов)'],
    ['Оценка', 'Оценка (⭐ средняя)'],
    ['Отзыв', 'Отзыв (строгая склейка -> ⭐N)']
  ];

  for (let i = 0; i < aliasPairs.length; i++) {
    const to = normHeader_(aliasPairs[i][0]);
    const from = normHeader_(aliasPairs[i][1]);
    if (!map.has(to) && map.has(from)) map.set(to, map.get(from));
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

/**
 * Ремонт: если "Сообщение" случайно оказалось в колонке "_foreignBrand",
 * переносим текст обратно, а "_foreignBrand" очищаем.
 */
function repairMessageMovedToForeignBrand_(sheet) {
  try {
    const lastCol = Math.max(sheet.getLastColumn(), HEADERS_MAIN.length);
    const hdr = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(normHeader_);

    let msgCol = 0;
    let fbCol = 0;

    for (let i = 0; i < hdr.length; i++) {
      const h = hdr[i];
      if (!msgCol && h === 'Сообщение') msgCol = i + 1;
      if (!fbCol && h === '_foreignBrand') fbCol = i + 1;
    }

    if (!msgCol || !fbCol || msgCol === fbCol) return;

    const dataLast = getDataLastRow_(sheet);
    if (dataLast < 2) return;

    const rows = dataLast - 1;
    const msgVals = sheet.getRange(2, msgCol, rows, 1).getValues();
    const fbVals = sheet.getRange(2, fbCol, rows, 1).getValues();

    let changed = false;

    for (let i = 0; i < rows; i++) {
      const m = String(msgVals[i][0] || '').trim();
      const fRaw = fbVals[i][0];
      const f = String(fRaw || '').trim();

      const isBoolLike =
        f === '' ||
        f === 'TRUE' || f === 'FALSE' ||
        f === 'true' || f === 'false' ||
        fRaw === true || fRaw === false;

      if (!m && f && !isBoolLike) {
        msgVals[i][0] = fbVals[i][0];
        fbVals[i][0] = '';
        changed = true;
      }
    }

    if (changed) {
      sheet.getRange(2, msgCol, rows, 1).setValues(msgVals);
      sheet.getRange(2, fbCol, rows, 1).setValues(fbVals);
    }
  } catch (e) {}
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

/**********************
 * МИГРАЦИЯ СХЕМЫ (осознанно, вручную)
 **********************/
function migrateSchemaToCurrentHeaders_() {
  withLock_(() => {
    const ss = SpreadsheetApp.getActive();
    const ui = SpreadsheetApp.getUi();

    const sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) {
      ui.alert('Лист "' + SHEET_NAME + '" не найден.');
      return;
    }

    const answer = ui.alert(
      'Миграция структуры',
      'Будет применена НОВАЯ структура колонок.\n' +
      'Данные будут перенесены по заголовкам.\n\n' +
      'Рекомендуется сделать копию таблицы.\n\n' +
      'Продолжить?',
      ui.ButtonSet.OK_CANCEL
    );
    if (answer !== ui.Button.OK) return;

    // --- бэкап листа (страховка) ---
    const backupName =
      SHEET_NAME + ' · backup ' +
      Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm');
    sh.copyTo(ss).setName(backupName);

    // --- убираем фильтр перед перестановкой ---
    try {
      const f = sh.getFilter();
      if (f) f.remove();
    } catch (e) {}

    toast_(ss, 'WB · MIGRATE: пересборка структуры…', 5);

    const lastCol = sh.getLastColumn();
    rebuildSheetByHeaders_(sh, lastCol);

    // финальная нормализация (в миграции разрешаем подрезку лишних колонок)
    ensureSheetLayout_(sh, { allowRebuild: true });

    SpreadsheetApp.flush();
    clearToast_(ss);

    ui.alert(
      'Миграция завершена',
      'Структура обновлена.\n' +
      'Данные сохранены.\n\n' +
      'Резервная копия листа создана:\n' +
      backupName,
      ui.ButtonSet.OK
    );
  });
}
