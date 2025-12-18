/**********************
 * СБРОС КЭША
 **********************/
function resetBrandCache_() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();

  let removed = 0;
  Object.keys(all).forEach(k => {
    if (k.startsWith('brand:')) {
      props.deleteProperty(k);
      removed++;
    }
  });

  SpreadsheetApp.getUi().alert('Кэш брендов очищен: ' + removed);
}

function resetAllCache_(silent) {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();

  const prefixes = ['brand:', 'rule:', 'claimfb:', 'warrantyM:'];
  let removed = 0;

  Object.keys(all).forEach(k => {
    if (prefixes.some(p => k.startsWith(p))) {
      props.deleteProperty(k);
      removed++;
    }
  });

  if (!silent) SpreadsheetApp.getUi().alert('Весь кэш очищен: ' + removed);
}

/**********************
 * SCRIPT PROPERTIES CACHE (TTL)
 **********************/
const PROP_CACHE_TS_SUFFIX = ':ts';
const PROP_CACHE_MISS = '__MISS__';
const PROP_CACHE_NA = 'NA';

function _propCacheGetFromAll_(allProps, key, ttlMs, nowMs, missingTsIsFresh) {
  if (!allProps || !Object.prototype.hasOwnProperty.call(allProps, key)) {
    return { exists: false, fresh: false, value: null, needsTouch: false };
  }

  const value = allProps[key];
  const tsKey = key + PROP_CACHE_TS_SUFFIX;
  const tsRaw = allProps[tsKey];

  if (!tsRaw) {
    if (missingTsIsFresh) {
      return { exists: true, fresh: true, value: value, needsTouch: true };
    }
    return { exists: true, fresh: false, value: value, needsTouch: false };
  }

  const ts = Number(tsRaw);
  if (!isFinite(ts) || ts <= 0) {
    if (missingTsIsFresh) {
      return { exists: true, fresh: true, value: value, needsTouch: true };
    }
    return { exists: true, fresh: false, value: value, needsTouch: false };
  }

  const age = (nowMs || Date.now()) - ts;
  const fresh = (!ttlMs || ttlMs <= 0) ? true : (age <= ttlMs);
  return { exists: true, fresh: fresh, value: value, needsTouch: false };
}

function propCacheGetFromAll_(allProps, key, ttlMs, nowMs) {
  return _propCacheGetFromAll_(allProps, key, ttlMs, nowMs, true);
}

function propCacheGetFromAllStrictTs_(allProps, key, ttlMs, nowMs) {
  return _propCacheGetFromAll_(allProps, key, ttlMs, nowMs, false);
}

function propCacheSet_(props, key, value, nowMs) {
  props.setProperty(key, String(value));
  props.setProperty(key + PROP_CACHE_TS_SUFFIX, String(nowMs || Date.now()));
}

function propCacheTouchTs_(props, keys, nowMs) {
  if (!keys || !keys.length) return;
  const ts = String(nowMs || Date.now());
  const uniq = {};
  keys.forEach(k => { uniq[k] = true; });
  Object.keys(uniq).forEach(k => {
    try { props.setProperty(k + PROP_CACHE_TS_SUFFIX, ts); } catch (e) {}
  });
}

function toastActive_(msg, sec) {
  try { toast_(SpreadsheetApp.getActive(), msg, sec || 6); } catch (e) {}
}

function log_(msg) {
  try { Logger.log(String(msg || '')); } catch (e) {}
}


/**********************
 * Lock + Toast helpers
 **********************/
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { fn(); }
  finally { try { lock.releaseLock(); } catch (e) {} }
}

function toast_(ss, msg, sec) {
  if (!ENABLE_TOASTS) return;
  try { ss.toast(String(msg || ''), 'WB · Возвраты', sec || 5); } catch (e) {}
}

function clearToast_(ss) {
  try { ss.toast('', '', 1); } catch (e) {}
}

/**********************
 * Throttle
 **********************/
function throttle_(key) {
  const r = RATE[key];
  if (!r) return;
  const now = Date.now();
  const wait = (r.last + r.minMs) - now;
  if (wait > 0) Utilities.sleep(wait);
  r.last = Date.now();
}

/**********************
 * A1 helpers
 **********************/
function colToA1_(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/**********************
 * TOKENS: setup + cached getter (ONE TOKEN)
 **********************/
function getTokenCached_(key) {
  const unifiedKey = TOKEN_KEYS.UNIFIED || 'WB_API_TOKEN';

  if (RUNTIME.tokens[unifiedKey]) {
    const t = RUNTIME.tokens[unifiedKey];
    if (key) RUNTIME.tokens[key] = t;
    return t;
  }

  const props = PropertiesService.getScriptProperties();

  // 1) основной единый ключ
  let v = props.getProperty(unifiedKey);

  // 2) совместимость: если вдруг сохранён старый
  if (!v) v = props.getProperty('WB_RETURNS_TOKEN') || props.getProperty('WB_FEEDBACKS_TOKEN') || props.getProperty('WB_CONTENT_TOKEN');

  if (!v) {
    throw new Error(
      `Не найден токен в Script Properties: ${unifiedKey}\n` +
      `Открой меню "WB · Возвраты" → "Настроить токен (1 раз)".`
    );
  }

  v = String(v).trim();

  RUNTIME.tokens[unifiedKey] = v;
  if (key) RUNTIME.tokens[key] = v;

  return v;
}

function setupTokens_() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const r = ui.prompt('WB_API_TOKEN', 'Вставь ЕДИНЫЙ токен WB API (для всех методов)', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;

  const t = (r.getResponseText() || '').trim();
  if (!t) {
    ui.alert('Токен пустой — отмена.');
    return;
  }

  // единый ключ
  props.setProperty('WB_API_TOKEN', t);

  // совместимость со старым (можно оставить, не мешает)
  props.setProperty('WB_RETURNS_TOKEN', t);
  props.setProperty('WB_FEEDBACKS_TOKEN', t);
  props.setProperty('WB_CONTENT_TOKEN', t);

  // runtime сброс
  RUNTIME.tokens = {};

  ui.alert('Единый токен сохранён в Script Properties.');
}
