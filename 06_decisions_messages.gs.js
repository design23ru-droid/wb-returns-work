/**********************
 * РЕШЕНИЯ + СООБЩЕНИЯ + onEdit (ускорение подстановки)
 **********************/

/**
 * РЕШЕНИЕ -> СООБЩЕНИЕ (лист "Сообщения")
 * Заполняем только если "Сообщение" пустое.
 */
function fillDecisionMessages_(sheet) {
  const map = getDecisionMessageMap_(); // Map(решение -> сообщение)

  const dataLastRow = getDataLastRow_(sheet);
  if (dataLastRow < 2) return;
  const rows = dataLastRow - 1;

  const decisions = sheet.getRange(2, COL.DECISION, rows, 1)
    .getDisplayValues().flat().map(v => String(v || '').trim());

  const curMsg = sheet.getRange(2, COL.MESSAGE, rows, 1)
    .getDisplayValues().flat().map(v => String(v || '').trim());

  const out = new Array(rows);
  const rowsToNormalize = [];

  for (let i = 0; i < rows; i++) {
    const msg = String(curMsg[i] || '').trim();
    if (msg) {
      out[i] = [curMsg[i]];
      continue;
    }

    const d = String(decisions[i] || '').trim();
    if (!d) {
      out[i] = [''];
      continue;
    }

    const v = map.get(d) || '';
    out[i] = [v];

    if (v) rowsToNormalize.push(i + 2); // A
  }

  sheet.getRange(2, COL.MESSAGE, rows, 1).setValues(out);
  try { sheet.getRange(2, COL.MESSAGE, rows, 1).setNumberFormat('@'); } catch (e) {}

  // A) автоподстановка → нормализуем высоту строки
  normalizeRowsHeight_(sheet, rowsToNormalize);
}


/**
 * Создать/поддержать лист "Сообщения"
 */
function ensureMessagesSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(MESSAGES_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(MESSAGES_SHEET_NAME);

  const headers = ['Решение', 'Сообщение'];
  sh.getRange(1, 1, 1, 2).setValues([headers]).setFontWeight('bold');

  const last = sh.getLastRow();
  const existing = new Set();
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, 1).getValues().flat().forEach(v => {
      const k = String(v || '').trim();
      if (k) existing.add(k);
    });
  }

  const toAdd = [];
  DECISIONS.forEach(d => {
    if (!existing.has(d)) toAdd.push([d, '']);
  });

  if (toAdd.length) {
    sh.getRange(sh.getLastRow() + 1, 1, toAdd.length, 2).setValues(toAdd);
  }

  try { sh.autoResizeColumns(1, 2); } catch (e) {}

  // прогреем runtime/cache карты сообщений (чтобы onEdit был быстрым)
  try { warmDecisionMessageMap_(); } catch (e) {}
}

/**
 * Выпадающий список "Решение"
 */
function applyDecisionDropdown_(sheet) {
  const lastRow = getDataLastRow_(sheet);
  if (lastRow < 2) return;

  const rows = lastRow - 1;

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DECISIONS, true)
    .setAllowInvalid(true)
    .build();

  sheet.getRange(2, COL.DECISION, rows, 1).setDataValidation(rule);
}

/**
 * onEdit:
 * 1) Если редактируют лист "Сообщения" — сбрасываем кэш (чтобы новые тексты сразу подхватывались).
 * 2) Если редактируют колонку "Решение" на основном листе — подставляем текст в "Сообщение" (если оно пустое).
 */
function onEdit(e) {
  if (!e || !e.range) return;

  const range = e.range;
  const sh = range.getSheet();
  const shName = sh.getName();

  // --- 1) Любая правка словаря сообщений -> сброс кэша
  if (shName === MESSAGES_SHEET_NAME) {
    const cols = _msgSheetCols_(sh); // {decisionCol, messageCol}
    if (!cols) return;

    const r = range.getRow();
    const c1 = range.getColumn();
    const c2 = c1 + range.getNumColumns() - 1;

    if (r >= 2 && !(c2 < cols.decisionCol || c1 > cols.messageCol)) {
      clearDecisionMessageCache_();
    }
    return;
  }

  // --- 2) Основной лист
  if (!_isReturnsSheet_(sh)) return;

  const rowStart = range.getRow();
  const colStart = range.getColumn();
  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();

  if (rowStart < 2) return;

  const decisionCol = COL.DECISION;
  const messageCol = COL.MESSAGE;
  const colEnd = colStart + numCols - 1;

  // B) Ручная правка "Сообщение" → нормализуем высоту строки
  if (messageCol >= colStart && messageCol <= colEnd) {
    const rowsToNormalize = [];
    for (let i = 0; i < numRows; i++) rowsToNormalize.push(rowStart + i);
    normalizeRowsHeight_(sh, rowsToNormalize);
  }

  // если "Решение" не задето — дальше не идём
  if (decisionCol < colStart || decisionCol > colEnd) return;

  const decisionRange = sh.getRange(rowStart, decisionCol, numRows, 1);
  const msgRange = sh.getRange(rowStart, messageCol, numRows, 1);

  const decisions = decisionRange.getDisplayValues().flat().map(v => String(v || '').trim());
  const currentMsgs = msgRange.getValues().flat().map(v => String(v || '').trim());

  const map = getDecisionMessageMap_();

  let changed = false;
  const out = new Array(numRows);
  const rowsToNormalizeAfterAuto = [];

  for (let i = 0; i < numRows; i++) {
    const d = String(decisions[i] || '').trim();
    const curMsg = String(currentMsgs[i] || '').trim();

    if (curMsg) {
      out[i] = [currentMsgs[i]];
      continue;
    }

    if (!d) {
      out[i] = [''];
      continue;
    }

    const v = map.get(d) || '';
    out[i] = [v];

    if (v) rowsToNormalizeAfterAuto.push(rowStart + i); // A
    if (v !== curMsg) changed = true;
  }

  if (changed) {
    msgRange.setValues(out);
    normalizeRowsHeight_(sh, rowsToNormalizeAfterAuto);
  }
}


/**********************
 * Внутренний кэш "Решение -> Сообщение"
 **********************/
const DEC_MSG_CACHE_KEY = 'WB_DECISION_MSG_MAP_V2';
let DEC_MSG_RUNTIME = null;

function clearDecisionMessageCache_() {
  DEC_MSG_RUNTIME = null;
  try { CacheService.getScriptCache().remove(DEC_MSG_CACHE_KEY); } catch (e) {}
}

function getDecisionMessageMap_() {
  if (DEC_MSG_RUNTIME) return DEC_MSG_RUNTIME;

  const cache = CacheService.getScriptCache();
  const cached = cache.get(DEC_MSG_CACHE_KEY);

  if (cached) {
    try {
      const obj = JSON.parse(cached) || {};
      const m = new Map();
      Object.keys(obj).forEach(k => m.set(k, obj[k]));
      DEC_MSG_RUNTIME = m;
      return DEC_MSG_RUNTIME;
    } catch (e) {}
  }

  return rebuildDecisionMessageMap_();
}

function warmDecisionMessageMap_() {
  getDecisionMessageMap_();
}

function rebuildDecisionMessageMap_() {
  const ss = SpreadsheetApp.getActive();
  const msgSheet = ss.getSheetByName(MESSAGES_SHEET_NAME);

  const m = new Map();
  if (msgSheet) {
    const cols = _msgSheetCols_(msgSheet) || { decisionCol: 1, messageCol: 2 };
    const last = msgSheet.getLastRow();
    if (last >= 2) {
      const w = Math.max(cols.messageCol, cols.decisionCol);
      const data = msgSheet.getRange(2, 1, last - 1, w).getValues();
      for (let i = 0; i < data.length; i++) {
        const k = String(data[i][cols.decisionCol - 1] || '').trim();
        const v = String(data[i][cols.messageCol - 1] || '').trim();
        if (k && !m.has(k)) m.set(k, v);
      }
    }
  }

  // в CacheService (объектом)
  const obj = {};
  m.forEach((v, k) => obj[k] = v);

  try { CacheService.getScriptCache().put(DEC_MSG_CACHE_KEY, JSON.stringify(obj), 6 * 60 * 60); } catch (e) {}

  DEC_MSG_RUNTIME = m;
  return m;
}

/**********************
 * ВСПОМОГАТЕЛЬНЫЕ
 **********************/

// Определяем, что это именно лист возвратов, даже если его переименовали
function _isReturnsSheet_(sheet) {
  try {
    if (sheet.getName() === SHEET_NAME) return true;

    const needCol = Math.max(COL.MESSAGE, COL.DECISION);
    const hdr = sheet.getRange(1, 1, 1, needCol).getDisplayValues()[0];
    const hDec = normHeader_(hdr[COL.DECISION - 1]);
    const hMsg = normHeader_(hdr[COL.MESSAGE - 1]);

    return (hDec === 'Решение' && hMsg === 'Сообщение');
  } catch (e) {
    return false;
  }
}

// Находим колонки "Решение"/"Сообщение" на листе-словаре (если их сдвинули)
function _msgSheetCols_(sheet) {
  try {
    const lastCol = sheet.getLastColumn();
    if (lastCol < 2) return { decisionCol: 1, messageCol: 2 };

    const hdr = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(v => String(v || '').trim());
    let decisionCol = 0;
    let messageCol = 0;

    for (let i = 0; i < hdr.length; i++) {
      const h = hdr[i];
      if (!decisionCol && normHeader_(h) === 'Решение') decisionCol = i + 1;
      if (!messageCol && normHeader_(h) === 'Сообщение') messageCol = i + 1;
    }

    if (!decisionCol) decisionCol = 1;
    if (!messageCol) messageCol = 2;

    // гарантируем порядок
    if (messageCol < decisionCol) {
      const t = decisionCol;
      decisionCol = messageCol;
      messageCol = t;
    }

    return { decisionCol, messageCol };
  } catch (e) {
    return null;
  }
}

// Нормализует высоту указанных строк до стандартной (как у пустой строки)
function normalizeRowsHeight_(sheet, rowNumbers) {
  if (!rowNumbers || !rowNumbers.length) return;

  const defaultHeight = sheet.getDefaultRowHeight();
  const uniq = Array.from(new Set(rowNumbers));

  for (let i = 0; i < uniq.length; i++) {
    try {
      sheet.setRowHeight(uniq[i], defaultHeight);
    } catch (e) {}
  }
}

