/**********************
 * 1.5) БРЕНД ПО NM ID (кэш)
 **********************/
function fillBrands_(sheet) {
  const props = PropertiesService.getScriptProperties();
  const allProps = props.getProperties();

  const dataLastRow = getDataLastRow_(sheet);
  if (dataLastRow < 2) return;

  const rows = dataLastRow - 1;

  const nmIds = sheet.getRange(2, COL.NM_ID, rows, 1).getValues().flat().map(x => String(x || '').trim());
  const curBrands = sheet.getRange(2, COL.BRAND, rows, 1).getValues().flat().map(x => String(x || '').trim());

  const need = [];
  const needSet = new Set();

  for (let i = 0; i < rows; i++) {
    const nmId = nmIds[i];
    if (!nmId) continue;

    const current = curBrands[i];
    if (current && current !== 'Без бренда') continue;

    const cacheKey = `brand:${nmId}`;
    if (allProps[cacheKey]) continue;

    if (!needSet.has(nmId)) {
      needSet.add(nmId);
      need.push(nmId);
    }
  }

  for (let i = 0; i < need.length; i++) {
    const nmId = need[i];
    const mini = getCardMiniByNmId_(nmId);
    const brand = (mini && String(mini.brand || '').trim()) ? String(mini.brand).trim() : 'Без бренда';
    props.setProperty(`brand:${nmId}`, brand);
  }

  const out = new Array(rows);
  const allAfter = props.getProperties();

  for (let i = 0; i < rows; i++) {
    const nmId = nmIds[i];
    const current = curBrands[i];

    if (!nmId) { out[i] = [current || '']; continue; }
    if (current && current !== 'Без бренда') { out[i] = [current]; continue; }

    const cached = allAfter[`brand:${nmId}`];
    out[i] = [cached || current || ''];
  }

  sheet.getRange(2, COL.BRAND, rows, 1).setValues(out);
}

/**********************
 * N = TRUE если бренда нет в листе "Бренды" (A)
 **********************/
function fillForeignBrandFlags_(sheet) {
  const brandsSheet = SpreadsheetApp.getActive().getSheetByName(BRANDS_SHEET_NAME);
  if (!brandsSheet) throw new Error('Лист "Бренды" не найден');

  const last = brandsSheet.getLastRow();
  const list = (last >= 1) ? brandsSheet.getRange(1, 1, last, 1).getValues().flat() : [];
  const allowed = new Set(list.map(v => String(v || '').trim()).filter(Boolean));

  const dataLastRow = getDataLastRow_(sheet);
  if (dataLastRow < 2) return;

  const rows = dataLastRow - 1;
  const brands = sheet.getRange(2, COL.BRAND, rows, 1).getValues().flat();

  const out = new Array(rows);
  for (let i = 0; i < rows; i++) {
    const b = String(brands[i] || '').trim();
    out[i] = [b ? !allowed.has(b) : ''];
  }

  sheet.getRange(2, COL.FOREIGN_BRAND, rows, 1).setValues(out);
  try { sheet.hideColumns(COL.FOREIGN_BRAND); } catch (e) {}
}

/**********************
 * 2) РЕЙТИНГИ
 **********************/
function loadRatings_(sheet) {
  const dataLastRow = getDataLastRow_(sheet);
  if (dataLastRow < 2) return;

  const rows = dataLastRow - 1;

  const nmIds = sheet.getRange(2, COL.NM_ID, rows, 1).getValues().flat().map(x => String(x || '').trim());
  const curE = sheet.getRange(2, COL.RATING, rows, 1).getDisplayValues().flat().map(x => String(x || '').trim());
  const curF = sheet.getRange(2, COL.RATING_COUNT, rows, 1).getValues().flat();

  const need = [];
  const needSet = new Set();

  for (let i = 0; i < rows; i++) {
    const nmId = nmIds[i];
    if (!nmId) continue;

    const hasE = !!curE[i];
    const hasF = (curF[i] !== '' && curF[i] !== null && typeof curF[i] !== 'undefined');
    if (hasE && hasF) continue;

    if (!needSet.has(nmId)) {
      needSet.add(nmId);
      need.push(nmId);
    }
  }

  for (let i = 0; i < need.length; i++) {
    const nmId = need[i];
    if (RUNTIME.ratingByNm[nmId]) continue;
    RUNTIME.ratingByNm[nmId] = getRatingStatsForNmId_(nmId);
  }

  const outE = new Array(rows);
  const outF = new Array(rows);

  for (let i = 0; i < rows; i++) {
    const nmId = nmIds[i];
    const stats = nmId ? (RUNTIME.ratingByNm[nmId] || null) : null;

    if (curE[i]) {
      outE[i] = [curE[i]];
    } else if (stats && stats.count > 0 && stats.avg !== '') {
      const v = Number(stats.avg);
      outE[i] = [isFinite(v) ? `⭐ ${v.toLocaleString('ru-RU', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}` : ''];
    } else {
      outE[i] = [''];
    }

    const alreadyCount = curF[i];
    if (alreadyCount !== '' && alreadyCount !== null && typeof alreadyCount !== 'undefined' && alreadyCount !== 0) {
      outF[i] = [alreadyCount];
    } else if (stats) {
      outF[i] = [stats.count || ''];
    } else {
      outF[i] = [''];
    }
  }

  sheet.getRange(2, COL.RATING, rows, 1).setValues(outE);
  sheet.getRange(2, COL.RATING_COUNT, rows, 1).setValues(outF);
}

/**********************
 * 3) УСЛОВИЕ ВОЗВРАТА (лист "Категории") -> G
 **********************/
function fillReturnConditions_(sheet) {
  const props = PropertiesService.getScriptProperties();
  const cached = props.getProperties();

  const catSheet = SpreadsheetApp.getActive().getSheetByName(CATEGORIES_SHEET_NAME);
  if (!catSheet) throw new Error('Лист "Категории" не найден');

  const catLastRow = catSheet.getLastRow();
  const categoryMap = new Map();
  if (catLastRow >= 2) {
    catSheet.getRange(2, 1, catLastRow - 1, 2).getValues().forEach(r => {
      if (r[0] && r[1]) categoryMap.set(String(r[0]).trim(), String(r[1]).trim());
    });
  }

  const dataLastRow = getDataLastRow_(sheet);
  if (dataLastRow < 2) return;

  const rows = dataLastRow - 1;

  const nmIds = sheet.getRange(2, COL.NM_ID, rows, 1).getValues().flat().map(x => String(x || '').trim());
  const curG = sheet.getRange(2, COL.RETURN_RULE, rows, 1).getDisplayValues().flat().map(x => String(x || '').trim());

  const need = [];
  const needSet = new Set();

  for (let i = 0; i < rows; i++) {
    const nmId = nmIds[i];
    if (!nmId) continue;
    if (curG[i]) continue;

    const cacheKey = `rule:${nmId}`;
    const cachedValue = cached[cacheKey];
    if (cachedValue === 'Заявка' || cachedValue === '14 дней') continue;

    if (!needSet.has(nmId)) {
      needSet.add(nmId);
      need.push(nmId);
    }
  }

  for (let i = 0; i < need.length; i++) {
    const nmId = need[i];
    const mini = getCardMiniByNmId_(nmId);
    const category = mini ? String(mini.subjectName || '').trim() : '';
    const fullRule = categoryMap.get(category) || '';
    const shortRule = normalizeReturnRule_(fullRule);

    if (shortRule === 'Заявка' || shortRule === '14 дней') {
      props.setProperty(`rule:${nmId}`, shortRule);
    }
  }

  const cachedAfter = props.getProperties();
  const out = new Array(rows);

  for (let i = 0; i < rows; i++) {
    const nmId = nmIds[i];
    const current = curG[i];

    if (current) { out[i] = [current]; continue; }
    if (!nmId) { out[i] = ['']; continue; }

    out[i] = [cachedAfter[`rule:${nmId}`] || ''];
  }

  sheet.getRange(2, COL.RETURN_RULE, rows, 1).setValues(out);
}
