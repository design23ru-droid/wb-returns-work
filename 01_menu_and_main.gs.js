/**********************
 * МЕНЮ
 **********************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('WB · Возвраты')
    .addItem('Загрузить возвраты в работе', 'loadReturnsAndRatings')
    .addItem('Загрузить (с очисткой кэша)', 'loadReturnsAndRatingsFresh_')
    .addSeparator()
    .addItem('Обновить (без загрузки новых)', 'refreshReturnsSheet_')
    .addItem('Обновить + дозагрузить новые', 'refreshAndLoadNew_')
    .addSeparator()
    .addItem('Настроить токен (1 раз)', 'setupTokens_')
    .addSeparator()
    .addItem('Сбросить кэш брендов', 'resetBrandCache_')
    .addItem('Сбросить весь кэш', 'resetAllCache_')
    .addToUi();
}

/**********************
 * ГЛАВНАЯ: Загрузка + всё
 **********************/
function loadReturnsAndRatings() {
  withLock_(() => {
    const ss = SpreadsheetApp.getActive();
    toast_(ss, 'WB: загрузка возвратов…', 5);

    ensureMessagesSheet_();

    // 👉 loadReturns_ СЧИТАЕТ newCount, но НИЧЕГО не показывает
    const ctx = loadReturns_();
    const sh = ctx.sheet;

    toast_(ss, 'WB: бренды + чужие…', 5);
    fillBrands_(sh);
    fillForeignBrandFlags_(sh);

    toast_(ss, 'WB: рейтинги…', 5);
    loadRatings_(sh);

    toast_(ss, 'WB: условия возврата…', 5);
    fillReturnConditions_(sh);

    toast_(ss, 'WB: отзыв (строгая склейка)…', 5);
    fillReturnFeedbacks_(sh, ctx.claimsMeta);

    toast_(ss, 'WB: покупка + гарантия + дедлайн…', 5);
    fillPurchaseDays_(sh, ctx.claimsMeta);
    fillWarrantyStatus_(sh, ctx.claimsMeta);
    fillDeadlines_(sh);

    toast_(ss, 'WB: решения + сообщения…', 5);
    applyDecisionDropdown_(sh);
    fillDecisionMessages_(sh);

    toast_(ss, 'WB: подсветка + сортировка…', 5);
    applyConditionalRules_(sh);
    autoSortByDate_(sh);

    SpreadsheetApp.flush();
    clearToast_(ss);

    // ✅ ФИНАЛЬНЫЙ СЧЁТЧИК (вариант A)
    SpreadsheetApp.getUi().alert(
      `Загрузка завершена.\nНовых заявок: ${ctx.newCount}`
    );
  });
}


/**********************
 * Загрузка с очисткой кэша
 **********************/
function loadReturnsAndRatingsFresh_() {
  withLock_(() => {
    resetAllCache_(true);
    loadReturnsAndRatings();
  });
}

/**********************
 * СЕРВИС: Обновить (без новых)
 **********************/
function refreshReturnsSheet_() {
  withLock_(() => {
    const ss = SpreadsheetApp.getActive();
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
    fillReturnFeedbacks_(sh, claimsMeta);

    toast_(ss, 'WB: покупка + гарантия + дедлайн…', 5);
    fillPurchaseDays_(sh, claimsMeta);
    fillWarrantyStatus_(sh, claimsMeta);
    fillDeadlines_(sh);

    toast_(ss, 'WB: решения + сообщения…', 5);
    applyDecisionDropdown_(sh);     // ← ВАЖНО
    fillDecisionMessages_(sh);

    toast_(ss, 'WB: подсветка + сортировка…', 5);
    applyConditionalRules_(sh);
    autoSortByDate_(sh);

    SpreadsheetApp.flush();
    clearToast_(ss);
    SpreadsheetApp.getUi().alert('Обновление завершено.');
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
