/**********************
 * 4) ОТЗЫВ (строгая склейка) — H
 **********************/
function fillReturnFeedbacks_(sheet, claimsMeta) {
  const props = PropertiesService.getScriptProperties();
  const cache = props.getProperties();
  const nowMs = Date.now();
  const touchTs = [];

  const dataLastRow = getDataLastRow_(sheet);
  if (dataLastRow < 2) return;

  const rows = dataLastRow - 1;

  const claimIds = sheet.getRange(2, COL.CLAIM_ID, rows, 1).getValues().flat().map(x => String(x || '').trim());
  const curH = sheet.getRange(2, COL.RETURN_FB, rows, 1).getDisplayValues().flat().map(x => String(x || '').trim());

  const group = new Map();
  const out = new Array(rows);

  // 1) Сначала пробуем кэш (с TTL). Если нет/протух — ставим задачу на поиск.
  for (let i = 0; i < rows; i++) {
    const claimId = claimIds[i];
    const current = curH[i];

    if (current) { out[i] = [current]; continue; }
    if (!claimId) { out[i] = ['']; continue; }

    const cacheKey = `claimfb:${claimId}`;

    if (Object.prototype.hasOwnProperty.call(cache, cacheKey)) {
      const raw = String(cache[cacheKey] || '');
      const isFound = raw.indexOf('⭐') === 0;

      const ttl = isFound ? CACHE_TTL.CLAIMFB_FOUND_MS : CACHE_TTL.CLAIMFB_MISS_MS;
      const r = isFound
        ? propCacheGetFromAll_(cache, cacheKey, ttl, nowMs)
        : propCacheGetFromAllStrictTs_(cache, cacheKey, ttl, nowMs);

      if (r.exists && r.fresh) {
        if (r.needsTouch) touchTs.push(cacheKey);
        const v = String(r.value || '');
        out[i] = [(v === PROP_CACHE_MISS || v === PROP_CACHE_NA) ? '' : v];
        continue;
      }
    }

    out[i] = [''];

    const meta = claimsMeta && claimsMeta[claimId];
    if (!meta || !meta.nm_id || !meta.order_dt || !meta.dt) continue;

    const nmId = String(meta.nm_id).trim();
    const orderDt = safeParseDate_(meta.order_dt);
    const claimDt = safeParseDate_(meta.dt);
    if (!nmId || !orderDt || !claimDt) continue;

    if (!group.has(nmId)) group.set(nmId, []);
    group.get(nmId).push({ rowIndex: i, claimId, orderDt, claimDt });
  }

  // миграция старого кэша (⭐N без ts)
  try { propCacheTouchTs_(props, touchTs, nowMs); } catch (e) {}

  const DAY_MS = 24 * 3600 * 1000;

  for (const [nmId, items] of group.entries()) {
    let minFrom = new Date(items[0].orderDt.getTime() - FEEDBACK_WINDOW_DAYS_BEFORE_ORDER * DAY_MS);
    let maxToStrict = new Date(items[0].claimDt.getTime() + FEEDBACK_WINDOW_DAYS_AFTER_CLAIM_STRICT * DAY_MS);
    let maxToFallback = new Date(items[0].claimDt.getTime() + FEEDBACK_WINDOW_DAYS_AFTER_CLAIM_FALLBACK * DAY_MS);

    items.forEach(it => {
      const from = new Date(it.orderDt.getTime() - FEEDBACK_WINDOW_DAYS_BEFORE_ORDER * DAY_MS);
      const toStrict = new Date(it.claimDt.getTime() + FEEDBACK_WINDOW_DAYS_AFTER_CLAIM_STRICT * DAY_MS);
      const toFallback = new Date(it.claimDt.getTime() + FEEDBACK_WINDOW_DAYS_AFTER_CLAIM_FALLBACK * DAY_MS);
      if (from < minFrom) minFrom = from;
      if (toStrict > maxToStrict) maxToStrict = toStrict;
      if (toFallback > maxToFallback) maxToFallback = toFallback;
    });

    // 2) Строгий поиск
    const fbStrict = fetchFeedbacksForNmIdWindow_(nmId, minFrom, maxToStrict);
    const idxStrict = buildFeedbackIndex_(fbStrict);

    const unresolved = [];

    items.forEach(it => {
      const stars = pickClosestStars_(idxStrict, it.orderDt.getTime(), FEEDBACK_STRICT_MINUTES);
      const val = (stars >= 1 && stars <= 5) ? `⭐${stars}` : '';
      if (val) {
        propCacheSet_(props, `claimfb:${it.claimId}`, val, nowMs);
        out[it.rowIndex][0] = val;
      } else {
        unresolved.push(it);
      }
    });

    // 3) Fallback (только если не нашли в строгом)
    if (unresolved.length) {
      const idxWide = (maxToFallback.getTime() <= maxToStrict.getTime())
        ? idxStrict
        : buildFeedbackIndex_(fetchFeedbacksForNmIdWindow_(nmId, minFrom, maxToFallback));

      unresolved.forEach(it => {
        const stars = pickClosestStars_(idxWide, it.orderDt.getTime(), FEEDBACK_FALLBACK_MINUTES);
        const val = (stars >= 1 && stars <= 5) ? `⭐${stars}` : '';
        if (val) {
          propCacheSet_(props, `claimfb:${it.claimId}`, val, nowMs);
          out[it.rowIndex][0] = val;
        } else {
          // MISS кэшируем ненадолго, чтобы отзыв мог подтянуться в следующий прогон
          propCacheSet_(props, `claimfb:${it.claimId}`, PROP_CACHE_MISS, nowMs);
          out[it.rowIndex][0] = '';
        }
      });
    }
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

    // сохраняем фактическое число дней (не обрезаем до 14),
    // а подсветку >14 делаем правилом условного форматирования
    out[i] = [diffDays];
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
  const nowMs = Date.now();
  const touchTs = [];

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
    const r = propCacheGetFromAll_(cached, cacheKey, CACHE_TTL.WARRANTY_MS, nowMs);
    if (r.needsTouch) touchTs.push(cacheKey);
    const cachedValue = (r.exists && r.fresh) ? String(r.value || '') : '';
    if (cachedValue) continue;

    if (!needSet.has(nmId)) {
      needSet.add(nmId);
      needNm.push(nmId);
    }
  }

  for (let i = 0; i < needNm.length; i++) {
    const nmId = needNm[i];
    const mini = getCardMiniByNmId_(nmId);
    const months = mini ? mini.warrantyMonths : null;
    propCacheSet_(props, `warrantyM:${nmId}`, (months && months > 0) ? String(months) : PROP_CACHE_NA, nowMs);
  }

  // миграция старого кэша (без ts): считаем «свежим», но проставляем ts
  try { propCacheTouchTs_(props, touchTs, nowMs); } catch (e) {}

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
    const r = propCacheGetFromAll_(cachedAfter, key, CACHE_TTL.WARRANTY_MS, nowMs);
    if (r.needsTouch) touchTs.push(key);
    const v = (r.exists && r.fresh) ? String(r.value || '') : '';
    if (!v || v === PROP_CACHE_NA) continue;

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

  try { propCacheTouchTs_(props, touchTs, nowMs); } catch (e) {}
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
