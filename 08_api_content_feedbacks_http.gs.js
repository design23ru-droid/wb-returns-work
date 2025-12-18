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
  const payload = {
    settings: {
      filter: { textSearch: String(key), withPhoto: -1 },
      cursor: { limit: 1 }
    }
  };

  const resp = fetchJsonPostWithRetryNonThrow_(url, getTokenCached_(TOKEN_KEYS.CONTENT), payload);
  const card = resp?.cards?.[0] || null;

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
  if (t.includes('14')) return '14 дней';
  return '';
}

/**********************
 * HTTP helpers (retry)
 **********************/
function fetchJsonWithRetry_(url, token) {
  let delayMs = 600;

  for (let i = 0; i < 10; i++) {
    const resp = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: token },
      muteHttpExceptions: true
    });

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

      wait = Math.round(wait * (0.85 + Math.random() * 0.3));
      Utilities.sleep(wait);
      delayMs = Math.min(15000, Math.round(delayMs * 1.8));
      continue;
    }

    throw new Error('WB API error ' + code + ': ' + text);
  }

  throw new Error('Retry limit exceeded (too many 429/5xx)');
}

function fetchJsonPostWithRetryNonThrow_(url, token, payloadObj) {
  let delayMs = 700;

  for (let i = 0; i < 10; i++) {
    const resp = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: {
        Authorization: token,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payloadObj),
      muteHttpExceptions: true
    });

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

      wait = Math.round(wait * (0.85 + Math.random() * 0.3));
      Utilities.sleep(wait);
      delayMs = Math.min(20000, Math.round(delayMs * 1.8));
      continue;
    }

    return null;
  }

  return null;
}
