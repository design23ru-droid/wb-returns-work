/**********************
 * 1) ЗАГРУЗКА ВОЗВРАТОВ + meta  (+ счётчик новых заявок)
 **********************/
function loadReturns_() {
  const sheet = getOrCreateSheet_(SHEET_NAME);
  ensureSheetLayout_(sheet);

  const dataLastRow = getDataLastRow_(sheet);

  const existingIds = new Set();
  if (dataLastRow > 1) {
    sheet
      .getRange(2, COL.CLAIM_ID, dataLastRow - 1, 1)
      .getValues()
      .forEach(r => r[0] && existingIds.add(String(r[0]).trim()));
  }

  const BASE_URL = 'https://returns-api.wildberries.ru/api/v1/claims';
  const LIMIT = 200;

  let offset = 0;
  let total = Infinity;

  const rowsToAdd = [];
  const claimsMeta = {};

  const token = getTokenCached_(TOKEN_KEYS.RETURNS);

  while (offset < total) {
    throttle_('returns');

    const url = `${BASE_URL}?is_archive=false&limit=${LIMIT}&offset=${offset}`;
    const data = fetchJsonWithRetry_(url, token);

    const claims = data.claims || [];
    total = data.total || 0;
    if (!claims.length) break;

    for (let i = 0; i < claims.length; i++) {
      const c = claims[i];
      const id = String(c.id);

      claimsMeta[id] = {
        id,
        nm_id: String(c.nm_id || '').trim(),
        order_dt: c.order_dt || '',
        dt: c.dt || ''
      };

      if (!existingIds.has(id)) {
        const row = new Array(HEADERS_MAIN.length).fill('');

        row[COL.DT - 1] = c.dt ? new Date(c.dt) : '';
        row[COL.NM_ID - 1] = c.nm_id || '';
        row[COL.CLAIM_ID - 1] = id;

        rowsToAdd.push(row);
      }
    }

    offset += claims.length;
  }

  // ✅ НОВОЕ: удалить строки, которых больше нет в активных
  const removedCount = pruneMissingClaims_(sheet, claimsMeta);

  const newCount = rowsToAdd.length;

  if (newCount) {
    const insertRow = getDataLastRow_(sheet) + 1;
    sheet
      .getRange(insertRow, 1, newCount, HEADERS_MAIN.length)
      .setValues(rowsToAdd);
  }

  try { sheet.hideColumns(COL.FOREIGN_BRAND); } catch (e) {}

  try {
    const ss = SpreadsheetApp.getActive();
    const msg = removedCount
      ? `Новых заявок: ${newCount} · Удалено: ${removedCount}`
      : `Новых заявок: ${newCount}`;
    ss.toast(msg, 'WB · Возвраты', 5);
  } catch (e) {}

  return { sheet, claimsMeta, newCount, removedCount };
}


/**********************
 * 1a) СИНХРОНИЗАЦИЯ: удалить заявки, которых нет в active
 **********************/
function pruneMissingClaims_(sheet, activeClaimsOrSet) {
  if (!sheet) return 0;

  const dataLastRow = getDataLastRow_(sheet);
  if (dataLastRow < 2) return 0;

  let activeSet = null;

  if (activeClaimsOrSet && typeof activeClaimsOrSet.has === 'function') {
    // Set
    activeSet = activeClaimsOrSet;
  } else {
    // Object (claimsMeta)
    const obj = activeClaimsOrSet || {};
    activeSet = new Set(Object.keys(obj));
  }

  // Если по какой-то причине нет данных — ничего не трогаем
  if (!activeSet || typeof activeSet.size !== 'number') return 0;

  const rows = dataLastRow - 1;

  const claimIds = sheet
    .getRange(2, COL.CLAIM_ID, rows, 1)
    .getValues()
    .flat()
    .map(v => String(v || '').trim());

  const rowsToDelete = [];
  const removedIds = [];

  for (let i = 0; i < claimIds.length; i++) {
    const id = claimIds[i];
    if (!id) continue;
    if (!activeSet.has(id)) {
      rowsToDelete.push(i + 2); // + header
      removedIds.push(id);
    }
  }

  if (!rowsToDelete.length) return 0;

  // Сжимаем в блоки подряд и удаляем снизу вверх
  rowsToDelete.sort((a, b) => a - b);

  const blocks = [];
  let start = rowsToDelete[0];
  let prev = rowsToDelete[0];
  let len = 1;

  for (let i = 1; i < rowsToDelete.length; i++) {
    const r = rowsToDelete[i];
    if (r === prev + 1) {
      len++;
      prev = r;
    } else {
      blocks.push({ start, len });
      start = r;
      prev = r;
      len = 1;
    }
  }
  blocks.push({ start, len });

  for (let i = blocks.length - 1; i >= 0; i--) {
    sheet.deleteRows(blocks[i].start, blocks[i].len);
  }

  // (необязательно, но полезно) подчистим кэш отзывов по удалённым claimId
  try {
    const props = PropertiesService.getScriptProperties();
    for (let i = 0; i < removedIds.length; i++) {
      const k = `claimfb:${removedIds[i]}`;
      props.deleteProperty(k);
      props.deleteProperty(k + PROP_CACHE_TS_SUFFIX);
    }
  } catch (e) {}

  return rowsToDelete.length;
}





/**********************
 * meta-only
 **********************/
function fetchClaimsMeta_() {
  const BASE_URL = 'https://returns-api.wildberries.ru/api/v1/claims';
  const LIMIT = 200;

  let offset = 0;
  let total = Infinity;
  const claimsMeta = {};

  const token = getTokenCached_(TOKEN_KEYS.RETURNS);

  while (offset < total) {
    throttle_('returns');

    const url = `${BASE_URL}?is_archive=false&limit=${LIMIT}&offset=${offset}`;
    const data = fetchJsonWithRetry_(url, token); // ✅ если недоступно — упадём и НЕ затрём вычисляемые колонки

    const claims = data.claims || [];
    total = data.total || 0;
    if (!claims.length) break;

    for (let i = 0; i < claims.length; i++) {
      const c = claims[i];
      const id = String(c.id);
      claimsMeta[id] = {
        id,
        nm_id: String(c.nm_id || '').trim(),
        order_dt: c.order_dt || '',
        dt: c.dt || ''
      };
    }

    offset += claims.length;
  }

  return claimsMeta;
}
