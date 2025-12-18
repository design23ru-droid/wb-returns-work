/************************************************************
 * WB · Возвраты (v6.7 decisions + messages) — ONE TOKEN
 *
 * A  Дата создания
 * B  Бренд
 * C  ID заявки
 * D  NM ID
 * E  Оценка (⭐ средняя)
 * F  Кол-во (кол-во отзывов)
 * G  Возврат (правило: Заявка / 14 дней)
 * H  Отзыв (строгая склейка -> ⭐N)
 * I  Покупка (дней) — НА ДАТУ ЗАЯВКИ (dt), день покупки не считаем
 * J  Гарантия — "На гарантии" / "Гарантия вышла"
 * K  Дедлайн (A + N дней -> 09:00)
 * L  Решение
 * M  Сообщение
 * N  _foreignBrand (служебный флаг TRUE/FALSE, скрыт)
 ************************************************************/

/**********************
 * НАСТРОЙКИ
 **********************/
const ENABLE_TOASTS = true;
const AUTO_CLEAR_CACHE_AFTER_LOAD = false;

// ✅ TTL кэшей в Script Properties (чтобы данные не устаревали навсегда)
const CACHE_TTL = {
  BRAND_MS:          7  * 24 * 3600 * 1000,   // brand:<nmId>
  RULE_MS:           7  * 24 * 3600 * 1000,   // rule:<nmId>
  WARRANTY_MS:       30 * 24 * 3600 * 1000,   // warrantyM:<nmId>
  CLAIMFB_FOUND_MS:  30 * 24 * 3600 * 1000,   // claimfb:<claimId> (⭐N)
  CLAIMFB_MISS_MS:   12 * 3600 * 1000         // claimfb:<claimId> (не найдено)
};

// ✅ Поиск отзывов (строгий + fallback)
const FEEDBACK_WINDOW_DAYS_BEFORE_ORDER = 1;          // order_dt - N дней
const FEEDBACK_WINDOW_DAYS_AFTER_CLAIM_STRICT = 1;    // claim_dt + N дней (строго)
const FEEDBACK_WINDOW_DAYS_AFTER_CLAIM_FALLBACK = 7;  // claim_dt + N дней (fallback)

const FEEDBACK_STRICT_MINUTES = 120;
const FEEDBACK_FALLBACK_MINUTES = 24 * 60;

// ✅ Дедлайн: +N дней, выставить время 09:00
const DEADLINE_ADD_DAYS = 4;
const DEADLINE_SET_HOUR = 9;
const DEADLINE_SET_MIN  = 0;
const DEADLINE_SET_SEC  = 0;

// Sheet names
const SHEET_NAME = 'Возвраты в работе';
const BRANDS_SHEET_NAME = 'Бренды';
const CATEGORIES_SHEET_NAME = 'Категории';
const MESSAGES_SHEET_NAME = 'Сообщения';


/**********************
 * Решения (фиксированный список)
 **********************/
const DECISIONS = [
  'Возврат',
  'Компенсация',
  'Возврат за отзыв',
  'Возврат 14д',
  'Нужно видео',
  'Арбитраж WB',
  '14д вышло',
  'Нет ответа',
  'Ожидание'
];

/**********************
 * TOKENS (Script Properties) — ЕДИНЫЙ ТОКЕН
 **********************/
const TOKEN_KEYS = {
  UNIFIED:  'WB_API_TOKEN',
  RETURNS:  'WB_API_TOKEN',
  FEEDBACKS:'WB_API_TOKEN',
  CONTENT:  'WB_API_TOKEN'
};

/**********************
 * RATE LIMITER
 **********************/
const RATE = {
  returns:   { minMs: 120, last: 0 },
  content:   { minMs: 650, last: 0 },
  feedbacks: { minMs: 380, last: 0 }
};

/**********************
 * RUNTIME CACHE
 **********************/
const RUNTIME = {
  cardMiniByNm: {},   // nmId -> { brand, subjectName, warrantyText, warrantyMonths }
  ratingByNm: {},     // nmId -> { avg, count }
  tokens: {}          // key -> token
};

/**********************
 * ЛИСТЫ / КОЛОНКИ
 **********************/
const HEADERS_MAIN = [
  'Дата создания',   // A
  'Бренд',           // B
  'ID заявки',       // C
  'NM ID',           // D
  'Оценка',          // E
  'Кол-во',          // F
  'Возврат',         // G
  'Отзыв',           // H
  'Покупка (дней)',  // I
  'Гарантия',        // J
  'Дедлайн',         // K
  'Решение',         // L
  'Сообщение',       // M
  '_foreignBrand'    // N
];

const COL = {
  DT: 1,
  BRAND: 2,
  CLAIM_ID: 3,
  NM_ID: 4,
  RATING: 5,
  RATING_COUNT: 6,
  RETURN_RULE: 7,
  RETURN_FB: 8,
  PURCHASE_DAYS: 9,
  WARRANTY: 10,
  DEADLINE: 11,
  DECISION: 12,
  MESSAGE: 13,
  FOREIGN_BRAND: 14
};
