/**********************
 * УСЛОВНОЕ ФОРМАТИРОВАНИЕ
 **********************/
function applyConditionalRules_(sheet) {
  const dataLastRow = getDataLastRow_(sheet);
  if (dataLastRow < 2) return;

  const rows = dataLastRow - 1;
  const rules = [];

  const nmL = colToA1_(COL.NM_ID);
  const deadlineL = colToA1_(COL.DEADLINE);
  const foreignL = colToA1_(COL.FOREIGN_BRAND);
  const warrantyL = colToA1_(COL.WARRANTY);

  // 1) Повторы рядом по NM ID
  if (dataLastRow >= 3) {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=OR(AND($${nmL}2<>"";$${nmL}2=$${nmL}1);AND($${nmL}2<>"";$${nmL}2=$${nmL}3))`)
        .setBackground('#FFF2CC')
        .setRanges([sheet.getRange(2, COL.NM_ID, rows, 1)])
        .build()
    );
  }

  // 2) Чужие бренды
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$${foreignL}2=TRUE`)
      .setBackground('#F4CCCC')
      .setRanges([sheet.getRange(2, COL.BRAND, rows, 1)])
      .build()
  );

  // 3) 14 дней истекли (там 14.0001)
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(14)
      .setBackground('#E06666')
      .setRanges([sheet.getRange(2, COL.PURCHASE_DAYS, rows, 1)])
      .build()
  );

  // 4) Гарантия вышла
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$${warrantyL}2="Гарантия вышла"`)
      .setBackground('#E06666')
      .setRanges([sheet.getRange(2, COL.WARRANTY, rows, 1)])
      .build()
  );

  // 5) Дедлайн: выходные — сразу, будни — после дедлайна
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(ISNUMBER($${deadlineL}2);OR(WEEKDAY($${deadlineL}2;2)>=6;NOW()>$${deadlineL}2))`)
      .setBackground('#FCE5CD')
      .setRanges([sheet.getRange(2, COL.DEADLINE, rows, 1)])
      .build()
  );

  sheet.setConditionalFormatRules(rules);
  try { sheet.hideColumns(COL.FOREIGN_BRAND); } catch (e) {}
}
