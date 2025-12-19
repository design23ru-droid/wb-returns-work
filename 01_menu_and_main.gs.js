/**********************
 * МЕНЮ
 **********************/
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  const adminMenu = ui
    .createMenu('Администрирование')
    .addItem('Токены', 'setupTokens_')
    .addItem('Сбросить кэш', 'resetCacheMenu_')
    .addSeparator()
    .addItem('Миграция', 'migrateReturnsSheetSchema_');

  ui.createMenu('WB · Возвраты')
    .addItem('Загрузить', 'loadReturnsAndRatings')
    .addItem('Обновить', 'refreshReturnsSheet_')
    .addSeparator()
    .addSubMenu(adminMenu)
    .addToUi();
}

/**********************
 * КЭШ: единый пункт (выбор)
 **********************/
function resetCacheMenu_() {
  const ui = SpreadsheetApp.getUi();

  const r = ui.prompt(
    'Сброс кэша',
    'Введите вариант:\n' +
      '1 — только кэш брендов\n' +
      '2 — весь кэш\n' +
      '0 — отмена',
    ui.ButtonSet.OK_CANCEL
  );

  if (r.getSelectedButton() !== ui.Button.OK) return;

  const v = String(r.getResponseText() || '').trim();

  if (v === '1') {
    resetBrandCache_();
    ui.alert('Готово: сброшен кэш брендов.');
    return;
  }

  if (v === '2') {
    const confirm = ui.alert(
      'Подтверждение',
      'Точно сбросить ВЕСЬ кэш?\nСледующая загрузка может быть заметно дольше.',
      ui.ButtonSet.OK_CANCEL
    );
    if (confirm !== ui.Button.OK) return;

    resetAllCache_();
    ui.alert('Готово: сброшен весь кэш.');
    return;
  }

  // 0 или любое другое значение — ничего не делаем
}



/**********************
 * ГЛАВНАЯ: Загрузка + всё
 **********************/
function loadReturnsAndRatings() {
  withLock_(() => {
    const ss = SpreadsheetApp.getActive();
    let stage = 'init';

    try {
      stage = 'ensureMessagesSheet_';
      toast_(ss, 'WB: загрузка возвратов…', 5);
      ensureMessagesSheet_();

      stage = 'loadReturns_';
      const ctx = loadReturns_();
      const sh = ctx.sheet;

      stage = 'fillBrands_ / fillForeignBrandFlags_';
      toast_(ss, 'WB: бренды + чужие…', 5);
      fillBrands_(sh);
      fillForeignBrandFlags_(sh);

      stage = 'loadRatings_';
      toast_(ss, 'WB: рейтинги…', 5);
      loadRatings_(sh);

      stage = 'fillReturnConditions_';
      toast_(ss, 'WB: условия возврата…', 5);
      fillReturnConditions_(sh);

      stage = 'fillReturnFeedbacks_';
      toast_(ss, 'WB: отзыв (строгая склейка)…', 5);
      fillReturnFeedbacks_(sh, ctx.claimsMeta);

      stage = 'fillPurchaseDays_ / fillWarrantyStatus_ / fillDeadlines_';
      toast_(ss, 'WB: покупка + гарантия + дедлайн…', 5);
      fillPurchaseDays_(sh, ctx.claimsMeta);
      fillWarrantyStatus_(sh, ctx.claimsMeta);
      fillDeadlines_(sh);

      stage = 'applyDecisionDropdown_ / fillDecisionMessages_';
      toast_(ss, 'WB: решения + сообщения…', 5);
      applyDecisionDropdown_(sh);
      fillDecisionMessages_(sh);

      stage = 'applyConditionalRules_ / autoSortByDate_';
      toast_(ss, 'WB: подсветка + сортировка…', 5);
      applyConditionalRules_(sh);
      autoSortByDate_(sh);

      SpreadsheetApp.flush();
      clearToast_(ss);

      const newCount = (ctx && typeof ctx.newCount === 'number') ? ctx.newCount : 0;
      const removedCount = (ctx && typeof ctx.removedCount === 'number') ? ctx.removedCount : 0;

      SpreadsheetApp.getUi().alert(
        `Загрузка завершена.\nНовых заявок: ${newCount}\nУдалено из активных: ${removedCount}`
      );
    } catch (e) {
      clearToast_(ss);

      const msg =
        'loadReturnsAndRatings: ошибка\n' +
        `Шаг: ${stage}\n\n` +
        (e && e.message ? e.message : String(e)) +
        (e && e.stack ? '\n\nSTACK:\n' + e.stack : '');

      try { console.error(msg); } catch (x) {}
      try { Logger.log(msg); } catch (x) {}

      SpreadsheetApp.getUi().alert(msg);
      throw e;
    }
  });
}




/**********************
 * Загрузка с очисткой кэша
 **********************/
function loadReturnsAndRatingsFresh_() {
  // ВАЖНО: не оборачиваем в withLock_ — loadReturnsAndRatings() уже берёт lock.
  // Иначе возможен дедлок/timeout при вложенном lock.waitLock().
  resetAllCache_(true);
  loadReturnsAndRatings();
}


/**********************
 * СЕРВИС: Обновить (без новых)
 **********************/
function refreshReturnsSheet_() {
  withLock_(() => {
    const ss = SpreadsheetApp.getActive();

    try {
      toast_(ss, 'WB: обновление…', 5);

      ensureMessagesSheet_();

      const sh = getOrCreateSheet_(SHEET_NAME);

      toast_(ss, 'WB: бренды + чужие…', 5);
      fillBrands_(sh);
      fillForeignBrandFlags_(sh);

      toast_(ss, 'WB: рейтинги…', 5);
      loadRatings_(sh);

      toast_(ss, 'WB: условия…', 5);
      fillReturnConditions_(sh);

      toast_(ss, 'WB: отзывы…', 5);
      const claimsMeta = fetchClaimsMeta_();

      // ✅ НОВОЕ: удалить строки, которых больше нет в активных
      const removedCount = pruneMissingClaims_(sh, claimsMeta);

      fillReturnFeedbacks_(sh, claimsMeta);

      toast_(ss, 'WB: покупка + гарантия + дедлайн…', 5);
      fillPurchaseDays_(sh, claimsMeta);
      fillWarrantyStatus_(sh, claimsMeta);
      fillDeadlines_(sh);

      toast_(ss, 'WB: решения + сообщения…', 5);
      applyDecisionDropdown_(sh);
      fillDecisionMessages_(sh);

      toast_(ss, 'WB: подсветка + сортировка…', 5);
      applyConditionalRules_(sh);
      autoSortByDate_(sh);

      SpreadsheetApp.flush();
      clearToast_(ss);
      SpreadsheetApp.getUi().alert(`Обновление завершено.\nУдалено из активных: ${removedCount || 0}`);
    } catch (e) {
      clearToast_(ss);

      const msg =
        'refreshReturnsSheet_: ошибка\n\n' +
        (e && e.message ? e.message : String(e)) +
        (e && e.stack ? '\n\nSTACK:\n' + e.stack : '');

      // и в лог тоже
      try { console.error(msg); } catch (x) {}
      try { Logger.log(msg); } catch (x) {}

      SpreadsheetApp.getUi().alert(msg);
      throw e; // чтобы ошибка фиксировалась в Executions
    }
  });
}



/**********************
 * СЕРВИС: Обновить + дозагрузить новые
 **********************/
function refreshAndLoadNew_() {
  refreshReturnsSheet_();
  loadReturnsAndRatings();
  SpreadsheetApp.getUi().alert('Готово: обновили лист и дозагрузили новые возвраты.');
}

/**********************

СЕРВИС: MIGRATE СТРУКТУРЫ (с бэкапом)
**********************/

function migrateReturnsSheetSchema_() {
  withLock_(() => {
    const ss = SpreadsheetApp.getActive();
    const ui = SpreadsheetApp.getUi();

    // 1) Найдём лист возвратов
    let sh = ss.getSheetByName(SHEET_NAME);

    // Если лист переименовали — попробуем определить по заголовкам
    if (!sh && typeof _isReturnsSheet_ === 'function') {
      const sheets = ss.getSheets();
      for (let i = 0; i < sheets.length; i++) {
        try {
          if (_isReturnsSheet_(sheets[i])) { sh = sheets[i]; break; }
        } catch (e) {}
      }
    }

    if (!sh) {
      ui.alert('WB · Возвраты', `Не найден лист "${SHEET_NAME}". Миграция остановлена.`, ui.ButtonSet.OK);
      return;
    }

    toast_(ss, 'WB: MIGRATE структуры…', 5);

    // 2) Резервная копия (на случай отката)
    const ts = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const backupName = `${sh.getName()} · backup ${ts}`;

    try {
      const backup = sh.copyTo(ss);
      backup.setName(backupName);
    } catch (e) {
      clearToast_(ss);
      ui.alert('WB · Возвраты', `Не удалось создать резервную копию.\n${e}`, ui.ButtonSet.OK);
      return;
    }

    // 3) Снимем фильтр (после перестановки колонок он может “ехать”)
    try {
      const f = sh.getFilter();
      if (f) f.remove();
    } catch (e) {}

    // 4) ЖЁСТКАЯ миграция: пересборка по заголовкам + нормализация лейаута
    try {
      if (typeof rebuildSheetByHeaders_ !== 'function') {
        throw new Error('rebuildSheetByHeaders_ is not defined (проверь 02_layout.gs.js)');
      }

      const lastCol = Math.max(sh.getLastColumn(), HEADERS_MAIN.length);
      rebuildSheetByHeaders_(sh, lastCol);

      ensureSheetLayout_(sh, { allowRebuild: true });

      // Если ранее "Сообщение" уехало в "_foreignBrand" — попробуем аккуратно вернуть
      if (typeof repairMessageMovedToForeignBrand_ === 'function') {
        repairMessageMovedToForeignBrand_(sh);
      }
    } catch (e) {
      clearToast_(ss);
      ui.alert('WB · Возвраты', `Ошибка миграции структуры:\n${e}`, ui.ButtonSet.OK);
      return;
    }

    // 5) Вернём фильтр на актуальный диапазон
    try {
      if (!sh.getFilter()) sh.getRange(1, 1, 1, HEADERS_MAIN.length).createFilter();
    } catch (e) {}

    // 6) Важно: условное форматирование и валидации пересобираем под новые индексы колонок
    try { applyDecisionDropdown_(sh); } catch (e) {}
    try { applyConditionalRules_(sh); } catch (e) {}
    try { autoSortByDate_(sh); } catch (e) {}

    SpreadsheetApp.flush();
    clearToast_(ss);

    ui.alert(
      'WB · Возвраты',
      'Структура обновлена.\nДанные сохранены.\n\n' +
      'Резервная копия листа создана:\n' + backupName,
      ui.ButtonSet.OK
    );
  });
}

