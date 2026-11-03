/**
 * Каркас склада: ДВИЖЕНИЕ / ОСТАТКИ / СТАРТ.
 * Цена учётная = последняя цена прихода.
 * Старые листы не удаляет.
 * Запуск: setupKarkasSklad, затем fillOstankiFromBase.
 */
function setupKarkasSklad() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var base = ss.getSheetByName('База_материалы');
  if (!base) throw new Error('Нет листа База_материалы');

  var d = ss.getSheetByName('ДВИЖЕНИЕ');
  if (!d) d = ss.insertSheet('ДВИЖЕНИЕ', 0);
  else {
    d.activate();
    ss.moveActiveSheet(1);
  }
  d.clear();
  d.getRange(1, 1, 1, 12).setValues([[
    'Дата', 'Тип', 'Источник', 'Документ', 'Наименование', 'Код',
    'Ед.', 'Кол-во', 'Цена', 'Сумма', 'Склад', 'Комментарий'
  ]]);
  d.getRange(1, 1, 1, 12).setFontWeight('bold').setBackground('#d9ead3');
  d.setFrozenRows(1);
  d.getRange('B2:B2000').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Приход', 'Расход'], true)
      .setAllowInvalid(false)
      .build()
  );
  d.getRange('E2:E2000').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(base.getRange('B2:B'), true)
      .setAllowInvalid(true)
      .build()
  );

  var fF = [];
  var fG = [];
  var fJ = [];
  for (var r = 2; r <= 300; r++) {
    fF.push(['=IF(E' + r + '="","",IFERROR(INDEX(\'База_материалы\'!C:C,MATCH(E' + r + ',\'База_материалы\'!B:B,0)),""))']);
    fG.push(['=IF(E' + r + '="","",IFERROR(INDEX(\'База_материалы\'!D:D,MATCH(E' + r + ',\'База_материалы\'!B:B,0)),""))']);
    fJ.push(['=IF(OR(H' + r + '="",I' + r + '=""),"",H' + r + '*I' + r + ')']);
  }
  d.getRange(2, 6, 299, 1).setFormulas(fF);
  d.getRange(2, 7, 299, 1).setFormulas(fG);
  d.getRange(2, 10, 299, 1).setFormulas(fJ);

  var o = ss.getSheetByName('ОСТАТКИ');
  if (!o) o = ss.insertSheet('ОСТАТКИ', 1);
  else {
    o.activate();
    ss.moveActiveSheet(2);
  }
  o.clear();
  o.getRange(1, 1, 2, 7).setValues([
    ['Сводка. Цена учётная = последняя цена прихода (ДВИЖЕНИЕ), иначе СТАРТ.', '', '', '', '', '', ''],
    ['Код', 'Наименование', 'Ед.', 'Склад', 'Кол-во', 'Цена учётная', 'Сумма']
  ]);
  o.getRange(1, 1, 1, 7).merge().setFontWeight('bold');
  o.getRange(2, 1, 1, 7).setFontWeight('bold').setBackground('#cfe2f3');
  o.setFrozenRows(2);

  var s = ss.getSheetByName('СТАРТ');
  if (!s) s = ss.insertSheet('СТАРТ', 2);
  else {
    s.activate();
    ss.moveActiveSheet(3);
  }
  s.clear();
  s.getRange(1, 1, 2, 7).setValues([
    ['Инвентаризация / стартовые остатки.', '', '', '', '', '', ''],
    ['Дата', 'Наименование', 'Код', 'Ед.', 'Кол-во', 'Цена', 'Сумма']
  ]);
  s.getRange(1, 1, 1, 7).merge().setFontWeight('bold');
  s.getRange(2, 1, 1, 7).setFontWeight('bold').setBackground('#fff2cc');
  s.setFrozenRows(2);
  s.getRange('B3:B2000').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(base.getRange('B2:B'), true)
      .setAllowInvalid(true)
      .build()
  );

  // Код/Ед. — через onEdit при выборе наименования (формулы setFormulas в этой таблице дают #ERROR!).
  // Сумма — вручную или =E*F после ввода; см. fixStartSheet.

  try {
    SpreadsheetApp.getUi().alert('Каркас готов: ДВИЖЕНИЕ, ОСТАТКИ, СТАРТ. Дальше fillOstankiFromBase');
  } catch (err) {
    Logger.log('Каркас готов: ДВИЖЕНИЕ, ОСТАТКИ, СТАРТ. Дальше fillOstankiFromBase');
  }
}

/** Убирает #ERROR! с листа СТАРТ (сломанные формулы локали). */
function fixStartSheet() {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('СТАРТ');
  if (!s) throw new Error('Нет листа СТАРТ');
  var last = Math.max(s.getLastRow(), 302);
  if (last < 3) last = 302;
  s.getRange(3, 3, last - 2, 1).clearContent(); // C Код
  s.getRange(3, 4, last - 2, 1).clearContent(); // D Ед.
  s.getRange(3, 7, last - 2, 1).clearContent(); // G Сумма
  Logger.log('СТАРТ: очищены C/D/G (ошибки формул). Код/Ед. подставятся при выборе наименования.');
}

/** Сумма на СТАРТ: G = E * F. Считает скриптом (формулы локали ломаются). */
function fixStartSum() {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('СТАРТ');
  if (!s) throw new Error('Нет листа СТАРТ');
  var last = Math.max(s.getLastRow(), 3);
  if (last < 3) return;
  var n = last - 2;
  var ef = s.getRange(3, 5, n, 2).getValues();
  var out = [];
  for (var i = 0; i < ef.length; i++) {
    var q = ef[i][0];
    var p = ef[i][1];
    if (q === '' || p === '' || q === null || p === null) {
      out.push(['']);
    } else {
      out.push([Number(q) * Number(p)]);
    }
  }
  s.getRange(3, 7, n, 1).setValues(out);
  Logger.log('СТАРТ: суммы пересчитаны, строк ' + n);
}

/** Пересчёт одной строки СТАРТ (сумма). */
function updateStartSum_(sheet, row) {
  var q = sheet.getRange(row, 5).getValue();
  var p = sheet.getRange(row, 6).getValue();
  if (q === '' || p === '' || q === null || p === null) {
    sheet.getRange(row, 7).clearContent();
    return;
  }
  sheet.getRange(row, 7).setValue(Number(q) * Number(p));
}

function fillOstankiFromBase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var base = ss.getSheetByName('База_материалы');
  var ost = ss.getSheetByName('ОСТАТКИ');
  if (!base || !ost) throw new Error('Нужны База_материалы и ОСТАТКИ');
  var last = base.getLastRow();
  if (last < 2) throw new Error('База пуста');
  // B=название, C=код, D=ед., E=склад (как в База_материалы)
  var data = base.getRange(2, 2, last - 1, 4).getValues();
  var values = [];
  for (var i = 0; i < data.length; i++) {
    var name = data[i][0];
    if (!name) continue;
    // Кол-во/цена/сумма = 0 до инвентаризации (СТАРТ) и живого ДВИЖЕНИЯ.
    // Формулы SUMIFS по ДВИЖЕНИЕ подключим отдельно (локаль таблицы ломает setFormulas).
    values.push([data[i][1] || '', name, data[i][2] || '', data[i][3] || '', 0, '', '']);
  }
  var header = (typeof ostHeaderRow_ === 'function') ? ostHeaderRow_(ost) : 2;
  var start = header + 1;
  if (ost.getMaxRows() >= start) {
    ost.getRange(start, 1, ost.getMaxRows() - start + 1, 7).clearContent();
  }
  if (!values.length) return;
  ost.getRange(start, 1, values.length, 7).setValues(values);
  Logger.log('ОСТАТКИ: строк ' + values.length + ' (qty=0 до инвентаризации)');
}
