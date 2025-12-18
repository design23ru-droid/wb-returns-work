/**********************
 * 4) ОТЗЫВ (строгая склейка) — H
 **********************/
function fillReturnFeedbacks_(sheet, claimsMeta) {
  const props = PropertiesService.getScriptProperties();
  const cache = props.getProperties();

  const dataLastRow = getDataLastRow_(sheet);
  if (dataLastRow < 2) return;

  const rows = dataLastRow - 1;

  const claimIds = sheet.getRange(2, COL.CLAIM_ID, rows, 1).getValues().flat().map(x => String(x || '').trim());
  const curH = sheet.getRange(2, COL.RETURN_FB, rows, 1).getDisplayValues().flat().map(x => String(x || '').trim());

  const group = new Map();

  for (let i = 0; i < rows; i++) {
    const claimId = claimIds[i];
    const current = curH[i];
    if (!claimId || current) continue;

    const cacheKey = `claimfb:${claimId}`;
    if (cache.hasOwnProperty(cacheKey)) continue;

    const meta = claimsMeta && claimsMeta[claimId];
    if (!meta || !meta.nm_id || !meta.order_dt || !meta.dt) continue;

    const nmId = String(meta.nm_id).trim();
    const orderDt = safeParseDate_(meta.order_dt);
    const claimDt = safeParseDate_(meta.dt);
    if (!nmId || !orderDt || !claimDt) continue;

    if (!group.has(nmId)) group.set(nmId, []);
    group.get(nmId).push({ rowIndex: i, claimId, orderDt, claimDt });
  }

  const out = new Array(rows);
  for (let i = 0; i < rows; i++) {
    const claimId = claimIds[i];
    const current = curH[i];

    if (current) { out[i] = [current]; continue; }

    const cacheKey = `claimfb:${claimId}`;
    if (claimId && cache.hasOwnProperty(cacheKey)) { out[i] = [cache[cacheKey]]; continue; }

    out[i] = [''];
  }

  const STRICT_MINUTES = 120;

  for (const [nmId, items] of group.entries()) {
    let minFrom = new Date(items[0].orderDt.getTime() - 24 * 3600 * 1000);
    let maxTo   = new Date(items[0].claimDt.getTime() + 24 * 3600 * 1000);

    items.forEach(it => {
      const from = new Date(it.orderDt.getTime() - 24 * 3600 * 1000);
      const to   = new Date(it.claimDt.getTime() + 24 * 3600 * 1000);
      if (from < minFrom) minFrom = from;
      if (to > maxTo) maxTo = to;
    });

    const feedbacks = fetchFeedbacksForNmIdWindow_(nmId, minFrom, maxTo);
    const idx = buildFeedbackIndex_(feedbacks);

    items.forEach(it => {
      const stars = pickClosestStars_(idx, it.orderDt.getTime(), STRICT_MINUTES);
      const val = (stars >= 1 && stars <= 5) ? `⭐${stars}` : '';
      props.setProperty(`claimfb:${it.claimId}`, val);
      out[it.rowIndex][0] = val;
    });
  }

  sheet.getRange(2, COL.RETURN_FB, rows, 1).setValues(out);
}

function buildFeedbackIndex_(feedbacks) {
  const arr = [];
  (feedbacks || []).forEach(fb => {
    const d = safeParseDate_(fb.lastOrderCreatedAt);
    if (!d) return;
    const v = Number(fb.productValuation);
    if (!(v >= 1 && v <= 5)) return;
    arr.push([d.getTime(), v]);
  });
  arr.sort((a, b) => a[0] - b[0]);
  return arr;
}

function pickClosestStars_(idxArr, targetMs, strictMinutes) {
  if (!idxArr || !idxArr.length) return NaN;

  let lo = 0, hi = idxArr.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (idxArr[mid][0] < targetMs) lo = mid + 1;
    else hi = mid;
  }

  const cand = [];
  for (let k = -2; k <= 2; k++) {
    const i = lo + k;
    if (i >= 0 && i < idxArr.length) cand.push(idxArr[i]);
  }

  let bestDiff = Infinity;
  let bestVal = NaN;

  for (let i = 0; i < cand.length; i++) {
    const diff = Math.abs(cand[i][0] - targetMs);
    if (diff < bestDiff) {
      bestDiff = diff;
      bestVal = cand[i][1];
    }
  }

  return (bestDiff <= strictMinutes * 60000) ? bestVal : NaN;
}

/**********************
 * 🛒 ПОКУПКА (дней) — I
 **********************/
function fillPurchaseDays_(sheet, claimsMeta) {
  const dataLastRow = getDataLastRow_(sheet);
  if (dataLastRow < 2) return;

  const rows = dataLastRow - 1;

  const claimIds = sheet
    .getRange(2, COL.CLAIM_ID, rows, 1)
    .getValues()
    .flat()
    .map(v => String(v || '').trim());

  const claimDtVals = sheet.getRange(2, COL.DT, rows, 1).getValues().flat();

  const DAY_MS = 24 * 3600 * 1000;
  const out = new Array(rows);

  for (let i = 0; i < rows; i++) {
    const claimId = claimIds[i];
    const meta = claimsMeta && claimsMeta[claimId];

    const orderRaw = meta ? meta.order_dt : '';
    if (!orderRaw) { out[i] = ['']; continue; }

    const orderDt = safeParseDate_(orderRaw);
    if (!orderDt) { out[i] = ['']; continue; }

    let claimDt = meta && meta.dt ? safeParseDate_(meta.dt) : null;
    if (!claimDt) {
      const v = claimDtVals[i];
      claimDt = (v instanceof Date) ? v : safeParseDate_(v);
    }
    if (!claimDt) { out[i] = ['']; continue; }

    const order0 = new Date(orderDt.getTime()); order0.setHours(0, 0, 0, 0);
    const claim0 = new Date(claimDt.getTime()); claim0.setHours(0, 0, 0, 0);

    const diffDays = Math.floor((claim0.getTime() - order0.getTime()) / DAY_MS);
    if (!isFinite(diffDays) || diffDays < 0) { out[i] = ['']; continue; }

    if (diffDays <= 14) out[i] = [diffDays];
    else out[i] = [14.0001]; // видно 14, но условие >14 сработает
  }

  sheet.getRange(2, COL.PURCHASE_DAYS, rows, 1).setValues(out);
  try { sheet.getRange(2, COL.PURCHASE_DAYS, rows, 1).setNumberFormat('0'); } catch (e) {}
}

/**********************
 * 🛡 ГАРАНТИЯ — J
 **********************/
function fillWarrantyStatus_(sheet, claimsMeta) {
  const dataLastRow = getDataLastRow_(sheet);
  if (dataLastRow < 2) return;

  const rows = dataLastRow - 1;

  const props = PropertiesService.getScriptProperties();
  const cached = props.getProperties();

  const claimIds = sheet.getRange(2, COL.CLAIM_ID, rows, 1).getValues().flat().map(v => String(v || '').trim());
  const curW = sheet.getRange(2, COL.WARRANTY, rows, 1).getDisplayValues().flat().map(v => String(v || '').trim());
  const claimDtVals = sheet.getRange(2, COL.DT, rows, 1).getValues().flat();

  const out = new Array(rows);
  for (let i = 0; i < rows; i++) out[i] = [curW[i] || ''];

  const needNm = [];
  const needSet = new Set();

  for (let i = 0; i < rows; i++) {
    if (out[i][0]) continue;
    const claimId = claimIds[i];
    const meta = claimsMeta && claimsMeta[claimId];
    const nmId = meta && meta.nm_id ? String(meta.nm_id).trim() : '';
    if (!nmId) continue;

    const cacheKey = `warrantyM:${nmId}`;
    if (cached.hasOwnProperty(cacheKey)) continue;

    if (!needSet.has(nmId)) {
      needSet.add(nmId);
      needNm.push(nmId);
    }
  }

  for (let i = 0; i < needNm.length; i++) {
    const nmId = needNm[i];
    const mini = getCardMiniByNmId_(nmId);
    const months = mini ? mini.warrantyMonths : null;
    props.setProperty(`warrantyM:${nmId}`, (months && months > 0) ? String(months) : 'NA');
  }

  const cachedAfter = props.getProperties();

  for (let i = 0; i < rows; i++) {
    if (out[i][0]) continue;

    const claimId = claimIds[i];
    const meta = claimsMeta && claimsMeta[claimId];
    if (!meta) continue;

    const nmId = meta.nm_id ? String(meta.nm_id).trim() : '';
    const orderRaw = meta.order_dt || '';
    if (!nmId || !orderRaw) continue;

    const orderDt = safeParseDate_(orderRaw);
    if (!orderDt) continue;

    let claimDt = meta.dt ? safeParseDate_(meta.dt) : null;
    if (!claimDt) {
      const v = claimDtVals[i];
      claimDt = (v instanceof Date) ? v : safeParseDate_(v);
    }
    if (!claimDt) continue;

    const key = `warrantyM:${nmId}`;
    const v = cachedAfter[key];
    if (!v || v === 'NA') continue;

    const months = Number(v);
    if (!isFinite(months) || months <= 0) continue;

    const order0 = new Date(orderDt.getTime()); order0.setHours(0,0,0,0);
    const claim0 = new Date(claimDt.getTime()); claim0.setHours(0,0,0,0);

    const end = addMonths_(order0, months);
    const status = (claim0.getTime() <= end.getTime()) ? 'На гарантии' : 'Гарантия вышла';

    out[i][0] = status;
  }

  sheet.getRange(2, COL.WARRANTY, rows, 1).setValues(out);
  try { sheet.getRange(2, COL.WARRANTY, rows, 1).setNumberFormat('@'); } catch (e) {}
}

function addMonths_(date, months) {
  const d = new Date(date.getTime());
  const day = d.getDate();
  d.setMonth(d.getMonth() + months);
  if (d.getDate() < day) d.setDate(0);
  d.setHours(0,0,0,0);
  return d;
}

function parseWarrantyMonths_(rawText) {
  const s0 = String(rawText || '').trim();
  if (!s0) return null;

  const s = s0.toLowerCase().replace(/\s+/g, ' ');
  if (/(отсутств|нет|без\s*гарант|не\s*предусмотр)/i.test(s)) return null;

  const yearMatch = s.match(/(\d+)\s*(год|года|лет|г\b)/i);
  const monMatch  = s.match(/(\d+)\s*(мес|месяц|месяцев|м\b)/i);

  let years = yearMatch ? Number(yearMatch[1]) : 0;
  let mons  = monMatch  ? Number(monMatch[1])  : 0;

  if (!isFinite(years)) years = 0;
  if (!isFinite(mons)) mons = 0;

  let total = 0;
  if (years || mons) {
    total = years * 12 + mons;
  } else {
    const n = s.match(/(\d+)/);
    if (!n) return null;
    total = Number(n[1]);
  }

  if (!isFinite(total) || total <= 0) return null;
  if (total > 120) return null;
  return total;
}

/**********************
 * ✅ ДЕДЛАЙН — K
 **********************/
function fillDeadlines_(sheet) {
  const dataLastRow = getDataLastRow_(sheet);
  if (dataLastRow < 2) return;

  const rows = dataLastRow - 1;
  const dtVals = sheet.getRange(2, COL.DT, rows, 1).getValues().flat();

  const out = new Array(rows);

  for (let i = 0; i < rows; i++) {
    const d0 = dtVals[i] instanceof Date ? dtVals[i] : safeParseDate_(dtVals[i]);
    if (!d0) { out[i] = ['']; continue; }

    const plus = new Date(d0.getTime() + DEADLINE_ADD_DAYS * 24 * 3600 * 1000);

    out[i] = [new Date(
      plus.getFullYear(),
      plus.getMonth(),
      plus.getDate(),
      DEADLINE_SET_HOUR, DEADLINE_SET_MIN, DEADLINE_SET_SEC
    )];
  }

  sheet.getRange(2, COL.DEADLINE, rows, 1).setValues(out);
  try { sheet.getRange(2, COL.DEADLINE, rows, 1).setNumberFormat('dd.MM.yyyy HH:mm'); } catch (e) {}
}
