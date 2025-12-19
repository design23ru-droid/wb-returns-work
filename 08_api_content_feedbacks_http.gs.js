/**********************
 * FEEDBACKS API helpers
 **********************/
function fetchFeedbacksForNmIdWindow_(nmId, fromDate, toDate) {
  const take = 5000;
  const all = [];

  const dateFrom = Math.floor(fromDate.getTime() / 1000);
  const dateTo = Math.floor(toDate.getTime() / 1000);

  ['true', 'false'].forEach(isAnsweredStr => {
    let skip = 0;

    while (true) {
      throttle_('feedbacks');

      const url =
        `https://feedbacks-api.wildberries.ru/api/v1/feedbacks` +
        `?nmId=${encodeURIComponent(nmId)}` +
        `&isAnswered=${isAnsweredStr}` +
        `&take=${take}` +
        `&skip=${skip}` +
        `&order=dateDesc` +
        `&dateFrom=${dateFrom}` +
        `&dateTo=${dateTo}`;

      const resp = fetchJsonWithRetry_(url, getTokenCached_(TOKEN_KEYS.FEEDBACKS));
      const list = resp?.data?.feedbacks || [];
      if (!list.length) break;

      all.push(...list);

      if (list.length < take) break;
      skip += list.length;

      if (skip > 50000) {
        log_(`Feedbacks API: достигнут лимит пагинации skip=${skip} (nmId=${nmId}, dateFrom=${dateFrom}, dateTo=${dateTo})`);
        break;
      }
    }
  });

  return all;
}

function getRatingStatsForNmId_(nmId) {
  const take = 5000;
  let sum = 0, count = 0;

  ['true', 'false'].forEach(isAnsweredStr => {
    let skip = 0;

    while (true) {
      if (skip > 199990) {
        log_(`Feedbacks API (rating): достигнут лимит пагинации skip=${skip} (nmId=${nmId})`);
        break;
      }

      throttle_('feedbacks');

      const url =
        `https://feedbacks-api.wildberries.ru/api/v1/feedbacks` +
        `?nmId=${encodeURIComponent(nmId)}` +
        `&isAnswered=${isAnsweredStr}` +
        `&take=${take}` +
        `&skip=${skip}`;

      const resp = fetchJsonWithRetry_(url, getTokenCached_(TOKEN_KEYS.FEEDBACKS));
      const list = resp?.data?.feedbacks || [];
      if (!list.length) break;

      list.forEach(f => {
        const v = Number(f.productValuation);
        if (v >= 1 && v <= 5) { sum += v; count++; }
      });

      if (list.length < take) break;
      skip += list.length;
    }
  });

  if (!count) return { avg: '', count: 0 };
  return { avg: Math.round(sum / count * 10) / 10, count };
}


/**********************
 * CONTENT API helpers
 **********************/
function getCardMiniByNmId_(nmId) {
  const key = String(nmId || '').trim();
  if (!key) return null;

  if (Object.prototype.hasOwnProperty.call(RUNTIME.cardMiniByNm, key)) {
    return RUNTIME.cardMiniByNm[key];
  }

  throttle_('content');

  const url = 'https://content-api.wildberries.ru/content/v2/get/cards/list';

  // Берём пачку результатов и выбираем строгое совпадение по nmID,
  // чтобы textSearch не вернул "похожую" карточку.
  const payload = {
    settings: {
      filter: { textSearch: key, withPhoto: -1 },
      cursor: { limit: 100 }
    }
  };

  const resp = fetchJsonPostWithRetryNonThrow_(url, getTokenCached_(TOKEN_KEYS.CONTENT), payload);
  const cards = resp?.cards || [];

  let card = null;
  if (Array.isArray(cards) && cards.length) {
    for (let i = 0; i < cards.length; i++) {
      const c = cards[i];
      const cNm = (c && (c.nmID ?? c.nmId)) || '';
      if (String(cNm).trim() === key) {
        card = c;
        break;
      }
    }
  }

  let warrantyText = '';
  let warrantyMonths = null;

  if (card && Array.isArray(card.characteristics)) {
    for (let i = 0; i < card.characteristics.length; i++) {
      const ch = card.characteristics[i];
      const name = String(ch?.name || '').trim().toLowerCase();
      if (name === 'гарантийный срок' || name.includes('гарантийн')) {
        let val = ch?.value;
        if (Array.isArray(val)) val = val.join(' ');
        warrantyText = String(val || '').trim();
        warrantyMonths = parseWarrantyMonths_(warrantyText);
        break;
      }
    }
  }

  const mini = card ? {
    brand: String(card.brand || '').trim(),
    subjectName: String(card.subjectName || '').trim(),
    warrantyText,
    warrantyMonths
  } : null;

  RUNTIME.cardMiniByNm[key] = mini;
  return mini;
}

function normalizeReturnRule_(text) {
  const t = String(text || '').toLowerCase();

  if (t.includes('заяв')) return 'Заявка';

  // 14 как отдельное число (не 214/140), но допускаем "14д", "14 дней", "14дн" и т.п.
  if (/(^|[^\d])14([^\d]|$)/.test(t)) return '14 дней';

  return '';
}


/**********************
 * HTTP helpers (retry)
 **********************/

function isTransientFetchError_(e) {
  const msg = String((e && e.message) ? e.message : e || '');
  // GAS может локализовать сообщение
  if (/Address unavailable/i.test(msg)) return true;
  if (/Адрес недоступен/i.test(msg)) return true;

  // частые сетевые/транспортные фейлы
  if (/timed out/i.test(msg)) return true;
  if (/Timeout/i.test(msg)) return true;
  if (/DNS/i.test(msg)) return true;
  if (/Socket/i.test(msg)) return true;
  if (/Service invoked too many times/i.test(msg)) return false; // квоты — не ретраим

  return false;
}

function sleepJitter_(ms) {
  const wait = Math.max(0, Math.round(ms * (0.85 + Math.random() * 0.3)));
  Utilities.sleep(wait);
}

function fetchJsonWithRetry_(url, token) {
  let delayMs = 600;

  for (let i = 0; i < 10; i++) {
    let resp;

    try {
      resp = UrlFetchApp.fetch(url, {
        method: 'get',
        headers: { Authorization: token },
        muteHttpExceptions: true
      });
    } catch (e) {
      if (isTransientFetchError_(e)) {
        sleepJitter_(delayMs);
        delayMs = Math.min(15000, Math.round(delayMs * 1.8));
        continue;
      }
      throw new Error('WB API fetch failed: ' + e + '\nURL: ' + url);
    }

    const code = resp.getResponseCode();
    const text = resp.getContentText();

    if (code === 200) return JSON.parse(text);

    if (code === 429 || (code >= 500 && code <= 599)) {
      const headers = resp.getAllHeaders ? resp.getAllHeaders() : {};
      const retryAfter = headers && (headers['Retry-After'] || headers['retry-after']);

      let wait = delayMs;
      if (retryAfter) {
        const sec = Number(retryAfter);
        if (!isNaN(sec) && sec > 0) wait = Math.max(wait, sec * 1000);
      }

      sleepJitter_(wait);
      delayMs = Math.min(15000, Math.round(delayMs * 1.8));
      continue;
    }

    throw new Error('WB API error ' + code + ': ' + text);
  }

  throw new Error(
    'WB API: превышен лимит ретраев (в т.ч. Address unavailable).\n' +
    'Если ошибка повторяется — часто это блокировка IP Google на стороне сервиса.'
  );
}

function fetchJsonPostWithRetryNonThrow_(url, token, payloadObj) {
  let delayMs = 700;

  for (let i = 0; i < 10; i++) {
    let resp;

    try {
      resp = UrlFetchApp.fetch(url, {
        method: 'post',
        headers: {
          Authorization: token,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify(payloadObj),
        muteHttpExceptions: true
      });
    } catch (e) {
      if (isTransientFetchError_(e)) {
        sleepJitter_(delayMs);
        delayMs = Math.min(20000, Math.round(delayMs * 1.8));
        continue;
      }
      return null;
    }

    const code = resp.getResponseCode();

    if (code === 200) {
      try { return JSON.parse(resp.getContentText()); } catch (e) { return null; }
    }

    if (code === 429 || (code >= 500 && code <= 599)) {
      const headers = resp.getAllHeaders ? resp.getAllHeaders() : {};
      const retryAfter = headers && (headers['Retry-After'] || headers['retry-after']);

      let wait = delayMs;
      if (retryAfter) {
        const sec = Number(retryAfter);
        if (!isNaN(sec) && sec > 0) wait = Math.max(wait, sec * 1000);
      }

      sleepJitter_(wait);
      delayMs = Math.min(20000, Math.round(delayMs * 1.8));
      continue;
    }

    return null;
  }

  return null;
}

