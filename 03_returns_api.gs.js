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

  while (offset < total) {
    throttle_('returns');

    const url = `${BASE_URL}?is_archive=false&limit=${LIMIT}&offset=${offset}`;
    let data;
    try {
      data = fetchJsonWithRetry_(url, getTokenCached_(TOKEN_KEYS.RETURNS));
    } catch (e) {
      throw new Error('WB Returns API error: ' + (e && e.message ? e.message : e));
    }

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
        rowsToAdd.push([
          c.dt ? new Date(c.dt) : '', // A
          '',                         // B
          id,                         // C
          c.nm_id || '',              // D
          '',                         // E
          '',                         // F
          '',                         // G
          '',                         // H
          '',                         // I
          '',                         // J
          '',                         // K
          '',                         // L Решение
          '',                         // M Сообщение
          ''                          // N _foreignBrand
        ]);
      }
    }

    offset += claims.length;
  }

  const newCount = rowsToAdd.length;

  if (newCount) {
    const insertRow = getDataLastRow_(sheet) + 1;
    sheet
      .getRange(insertRow, 1, newCount, HEADERS_MAIN.length)
      .setValues(rowsToAdd);
  }

  try { sheet.hideColumns(COL.FOREIGN_BRAND); } catch (e) {}

  // 👉 СЧЁТЧИК НОВЫХ ЗАЯВОК (вариант A)
  try {
    const ss = SpreadsheetApp.getActive();
    ss.toast(`Новых заявок: ${newCount}`, 'WB · Возвраты', 5);
  } catch (e) {}

  return { sheet, claimsMeta, newCount };
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

  while (offset < total) {
    throttle_('returns');

    const url = `${BASE_URL}?is_archive=false&limit=${LIMIT}&offset=${offset}`;
    let data;
    try {
      data = fetchJsonWithRetry_(url, getTokenCached_(TOKEN_KEYS.RETURNS));
    } catch (e) {
      throw new Error('WB Returns API error (meta): ' + (e && e.message ? e.message : e));
    }

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

