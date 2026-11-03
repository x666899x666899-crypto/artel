/**
 * Форма приход/расход — лист ДВИЖЕНИЕ.
 * Меню: Склад → Добавить приход / расход.
 */
var DVIZH_HEADER_ROW = 4;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Склад')
    .addItem('➕ Добавить приход', 'showPrihodForm')
    .addItem('➖ Добавить расход', 'showRashodForm')
    .addItem('Обновить шапку ДВИЖЕНИЕ', 'setupDvizhenieBanner')
    .addItem('Пересчитать остатки', 'rebuildOstanki')
    .addItem('Обновить поиск на ОСТАТКИ', 'setupOstankiSearch')
    .addItem('Обновить лист К ЗАКАЗУ', 'setupOrderSheet')
    .addItem('Обновить лист РЕЗЕРВ', 'setupReserveSheet')
    .addToUi();
}

function showPrihodForm() { openDvizhenieSidebar_('Приход'); }
function showRashodForm() { openDvizhenieSidebar_('Расход'); }

function openDvizhenieSidebar_(tip) {
  ensureDvizhenieLayout_();
  var html = HtmlService.createHtmlOutput(buildFormHtml_(tip, getMaterialNames_()))
    .setTitle(tip === 'Приход' ? 'Приход на склад' : 'Расход со склада')
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

function setupDvizhenieBanner() {
  ensureDvizhenieLayout_();
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ДВИЖЕНИЕ').activate();
  try {
    SpreadsheetApp.getUi().alert('Шапка готова. Меню «Склад» → Добавить приход / расход.');
  } catch (e) {}
}

function ensureDvizhenieLayout_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var d = ss.getSheetByName('ДВИЖЕНИЕ');
  if (!d) throw new Error('Нет листа ДВИЖЕНИЕ');
  var a1 = String(d.getRange(1, 1).getValue() || '');
  if (a1.indexOf('Склад') >= 0 && a1.indexOf('Добавить') >= 0) {
    d.setFrozenRows(DVIZH_HEADER_ROW);
    return;
  }
  if (a1 === 'Дата') d.insertRowsBefore(1, 3);
  d.getRange(1, 1, 3, 12).clear();
  d.getRange(1, 1, 1, 12).merge()
    .setValue('СКЛАД · ДВИЖЕНИЕ  |  Меню сверху: Склад → ➕ Добавить приход  или  ➖ Добавить расход')
    .setBackground('#1a3a2a').setFontColor('#fff').setFontWeight('bold')
    .setVerticalAlignment('middle');
  d.setRowHeight(1, 36);
  d.getRange(2, 1, 1, 12).merge()
    .setValue('Справа откроется форма с подсказками. Заполни → «Записать». Строка появится в журнале ниже. Код, ед. и сумма — сами.')
    .setBackground('#e8f5e9').setFontColor('#1b5e20').setWrap(true).setVerticalAlignment('middle');
  d.setRowHeight(2, 48);
  d.getRange(3, 1, 1, 12).merge()
    .setValue('Журнал операций ↓').setBackground('#f5f5f5').setFontColor('#666').setFontWeight('bold');
  var h = d.getRange(DVIZH_HEADER_ROW, 1, 1, 12).getValues()[0];
  if (String(h[0] || '') !== 'Дата') {
    d.getRange(DVIZH_HEADER_ROW, 1, 1, 12).setValues([[
      'Дата', 'Тип', 'Источник', 'Документ', 'Наименование', 'Код',
      'Ед.', 'Кол-во', 'Цена', 'Сумма', 'Склад', 'Комментарий'
    ]]);
  }
  d.getRange(DVIZH_HEADER_ROW, 1, 1, 12).setFontWeight('bold').setBackground('#d9ead3');
  d.setFrozenRows(DVIZH_HEADER_ROW);
  d.getRange('B5:B2000').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Приход', 'Расход'], true).setAllowInvalid(false).build()
  );
  var base = ss.getSheetByName('База_материалы');
  if (base) {
    d.getRange('E5:E2000').setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInRange(base.getRange('B2:B'), true).setAllowInvalid(true).build()
    );
  }
}

function getMaterialNames_() {
  var base = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('База_материалы');
  if (!base) return [];
  var last = base.getLastRow();
  if (last < 2) return [];
  var vals = base.getRange(2, 2, last - 1, 1).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) if (vals[i][0]) out.push(String(vals[i][0]));
  return out;
}

function lookupMaterial_(name) {
  var base = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('База_материалы');
  if (!base || !name) return { code: '', unit: '', sklad: '' };
  var last = base.getLastRow();
  var data = base.getRange(2, 2, last - 1, 4).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(name)) {
      return { code: data[i][1] || '', unit: data[i][2] || '', sklad: data[i][3] || '' };
    }
  }
  return { code: '', unit: '', sklad: '' };
}

function submitDvizhenieForm(payload) {
  ensureDvizhenieLayout_();
  var tip = payload.tip === 'Расход' ? 'Расход' : 'Приход';
  var name = String(payload.name || '').trim();
  if (!name) throw new Error('Выбери наименование из списка');
  var qty = Number(payload.qty);
  if (!(qty > 0)) throw new Error('Укажи количество больше нуля');
  var priceRaw = payload.price;
  var price = (priceRaw === '' || priceRaw == null) ? '' : Number(priceRaw);
  if (tip === 'Приход' && !(price > 0)) throw new Error('На приходе укажи цену за 1 единицу');
  if (price !== '' && !(price >= 0)) throw new Error('Цена должна быть числом');
  var mat = lookupMaterial_(name);
  var sum = price === '' ? '' : Math.round(qty * price * 100) / 100;
  var dateVal = payload.date ? new Date(payload.date + 'T12:00:00') : new Date();
  var sklad = String(payload.sklad || mat.sklad || '').trim();
  var d = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ДВИЖЕНИЕ');
  var row = nextDvizhenieRow_(d);
  d.getRange(row, 1, 1, 12).setValues([[
    dateVal, tip, String(payload.source || '').trim(), String(payload.doc || '').trim(),
    name, mat.code, mat.unit, qty, price === '' ? '' : price, sum, sklad,
    String(payload.comment || '').trim()
  ]]);
  d.getRange(row, 2).setBackground(tip === 'Приход' ? '#c8e6c9' : '#ffcdd2');
  var ost = applyOstankiFromMove_(name, mat.code, mat.unit, sklad, tip, qty, price === '' ? '' : price);
  return { ok: true, row: row, tip: tip, name: name, ostQty: ost ? ost.qty : '' };
}

function nextDvizhenieRow_(sheet) {
  var last = sheet.getLastRow();
  var start = DVIZH_HEADER_ROW + 1;
  if (last < start) return start;
  var names = sheet.getRange(start, 5, last - start + 1, 1).getValues();
  for (var i = 0; i < names.length; i++) if (!names[i][0]) return start + i;
  return last + 1;
}

function esc_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function buildFormHtml_(tip, names) {
  var isIn = tip === 'Приход';
  var c = isIn ? '#2e7d32' : '#c62828';
  var bg = isIn ? '#e8f5e9' : '#ffebee';
  var title = isIn ? 'Приход на склад' : 'Расход со склада';
  var btn = isIn ? 'Записать приход' : 'Записать расход';
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Irkutsk', 'yyyy-MM-dd');
  var opts = names.slice(0, 2500).map(function (n) { return '<option value="' + esc_(n) + '">'; }).join('');
  var srcHint = isIn ? 'Откуда пришло: поставщик, объект, цех…' : 'Куда ушло: объект, заказ, цех…';
  var srcPh = isIn ? 'Поставщик или объект' : 'Объект или заказ';
  var priceReq = isIn ? ' <span style="color:' + c + '">*</span>' : '';
  var priceHint = isIn ? 'Обязательно на приходе.' : 'Можно пусто на расходе.';
  return '<!DOCTYPE html><html><head><base target="_top"><meta charset="utf-8"><style>' +
    '*{box-sizing:border-box;margin:0;padding:0}' +
    'body{font:14px/1.45 system-ui,Segoe UI,sans-serif;color:#1a1a1a;background:#fafafa;padding:14px}' +
    '.h{background:' + bg + ';border-left:4px solid ' + c + ';padding:10px 12px;margin-bottom:14px;border-radius:0 8px 8px 0}' +
    '.h b{display:block;font-size:16px;color:' + c + ';margin-bottom:2px}' +
    '.h span{font-size:12px;color:#444}' +
    '.f{margin-bottom:12px}label{display:block;font-weight:600;font-size:13px;margin-bottom:3px}' +
    '.hint{font-size:11px;color:#666;margin-bottom:4px}' +
    'input,textarea{width:100%;padding:9px 10px;border:1px solid #ccc;border-radius:8px;font:inherit}' +
    'input:focus,textarea:focus{outline:2px solid ' + c + ';border-color:' + c + '}' +
    'textarea{min-height:56px;resize:vertical}' +
    '.r{display:grid;grid-template-columns:1fr 1fr;gap:10px}' +
    'button{width:100%;margin-top:8px;padding:12px;border:0;border-radius:10px;background:' + c +
    ';color:#fff;font-weight:700;font-size:15px;cursor:pointer}button:disabled{opacity:.55}' +
    '.m{margin-top:10px;padding:8px 10px;border-radius:8px;font-size:12px;display:none}' +
    '.ok{display:block;background:#e8f5e9;color:#1b5e20}.err{display:block;background:#ffebee;color:#b71c1c}' +
    '.req{color:' + c + '}' +
    '</style></head><body>' +
    '<div class="h"><b>' + title + '</b><span>Заполни поля и нажми кнопку. Строка появится в журнале ДВИЖЕНИЕ.</span></div>' +
    '<form id="f"><input type="hidden" id="tip" value="' + tip + '">' +
    '<div class="f"><label>Дата <span class="req">*</span></label><div class="hint">Когда операция. По умолчанию — сегодня.</div>' +
    '<input type="date" id="date" value="' + today + '" required></div>' +
    '<div class="f"><label>Наименование <span class="req">*</span></label>' +
    '<div class="hint">Выбери из базы (начни печатать — появятся варианты).</div>' +
    '<input list="ml" id="matName" placeholder="Например: краска нитра серая" required autocomplete="off">' +
    '<datalist id="ml">' + opts + '</datalist></div>' +
    '<div class="r"><div class="f"><label>Кол-во <span class="req">*</span></label>' +
    '<div class="hint">Сколько единиц (шт, кг, м…).</div>' +
    '<input type="number" id="qty" min="0" step="any" placeholder="0" required></div>' +
    '<div class="f"><label>Цена за 1 ед.' + priceReq + '</label><div class="hint">' + priceHint + '</div>' +
    '<input type="number" id="price" min="0" step="any" placeholder="0"></div></div>' +
    '<div class="f"><label>Источник / куда</label><div class="hint">' + srcHint + '</div>' +
    '<input type="text" id="source" placeholder="' + srcPh + '"></div>' +
    '<div class="f"><label>Документ</label><div class="hint">Номер накладной, УПД, заявки — если есть.</div>' +
    '<input type="text" id="doc" placeholder="№ …"></div>' +
    '<div class="f"><label>Склад</label><div class="hint">Необязательно — подтянется из базы.</div>' +
    '<input type="text" id="sklad" placeholder="Необязательно"></div>' +
    '<div class="f"><label>Комментарий</label><div class="hint">Что важно запомнить кладовщику.</div>' +
    '<textarea id="comment" placeholder="Коротко своими словами"></textarea></div>' +
    '<button type="submit" id="btn">' + btn + '</button><div class="m" id="msg"></div></form>' +
    '<script>document.getElementById("f").onsubmit=function(ev){ev.preventDefault();' +
    'var g=function(id){return document.getElementById(id)};' +
    'var b=g("btn"),m=g("msg"),nm=g("matName");' +
    'm.className="m";m.textContent="";b.disabled=true;' +
    'var p={tip:g("tip").value,date:g("date").value,name:nm.value,qty:g("qty").value,price:g("price").value,' +
    'source:g("source").value,doc:g("doc").value,sklad:g("sklad").value,comment:g("comment").value};' +
    'google.script.run.withSuccessHandler(function(r){b.disabled=false;m.className="m ok";' +
    'm.textContent="Записано · "+r.name+" · остаток "+r.ostQty;' +
    'g("qty").value="";g("price").value="";g("comment").value="";nm.focus();})' +
    '.withFailureHandler(function(e){b.disabled=false;m.className="m err";m.textContent=e.message||String(e);})' +
    '.submitDvizhenieForm(p);};</script></body></html>';
}

/** Приход плюсует остаток и ставит последнюю цену. Расход минусует. */
function applyOstankiFromMove_(name, code, unit, sklad, tip, qty, price) {
  var ost = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ОСТАТКИ');
  if (!ost || !name) return null;
  var key = String(name).trim();
  var header = ostHeaderRow_(ost);
  var start = header + 1;
  var last = Math.max(ost.getLastRow(), header);
  var names = last >= start ? ost.getRange(start, 2, last - header, 1).getValues() : [];
  var row = 0;
  for (var i = 0; i < names.length; i++) {
    if (String(names[i][0]).trim() === key) { row = start + i; break; }
  }
  var curQty = 0;
  var curPrice = '';
  if (!row) {
    row = last + 1;
    ost.getRange(row, 1, 1, 4).setValues([[code || '', key, unit || '', sklad || '']]);
  } else {
    var cur = ost.getRange(row, 5, 1, 2).getValues()[0];
    curQty = Number(cur[0]) || 0;
    curPrice = cur[1];
    if (code) ost.getRange(row, 1).setValue(code);
    if (unit) ost.getRange(row, 3).setValue(unit);
    if (sklad) ost.getRange(row, 4).setValue(sklad);
  }
  var newQty = tip === 'Расход' ? curQty - Number(qty) : curQty + Number(qty);
  newQty = Math.round(newQty * 1000) / 1000;
  var newPrice = curPrice;
  if (tip !== 'Расход' && price !== '' && Number(price) > 0) newPrice = Number(price);
  var newSum = (newPrice === '' || newPrice === null) ? '' : Math.round(newQty * Number(newPrice) * 100) / 100;
  ost.getRange(row, 5, 1, 3).setValues([[newQty, newPrice === '' ? '' : newPrice, newSum]]);
  ost.getRange(row, 5).setBackground(newQty < 0 ? '#ffcdd2' : '#ffffff');
  var catNow = String(ost.getRange(row, 8).getValue() || '');
  if (!catNow) catNow = ostCategoryForCode_(code);
  ost.getRange(row, 8).setValue(catNow || '');
  ost.getRange(row, 9).setValue(String(code || '') + ' ' + key);
  try { refreshOrderList_(); } catch (ignoreOrder) {}
  return { row: row, qty: newQty };
}

/** Полный пересчёт ОСТАТКИ: ноль + СТАРТ + все строки ДВИЖЕНИЕ. */
function rebuildOstanki() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ost = ss.getSheetByName('ОСТАТКИ');
  if (!ost) throw new Error('Нет листа ОСТАТКИ');
  var ostHeader = ostHeaderRow_(ost);
  var dataStart = ostHeader + 1;
  var last = ost.getLastRow();
  var rows = last >= dataStart ? ost.getRange(dataStart, 1, last - ostHeader, 7).getValues() : [];
  var index = {};
  var list = [];
  for (var i = 0; i < rows.length; i++) {
    var nm = String(rows[i][1] || '').trim();
    if (!nm) continue;
    index[nm] = list.length;
    list.push({ code: rows[i][0] || '', name: nm, unit: rows[i][2] || '', sklad: rows[i][3] || '', qty: 0, price: '' });
  }
  function ensure(name, code, unit, sklad) {
    name = String(name || '').trim();
    if (!name) return null;
    if (index[name] == null) {
      index[name] = list.length;
      list.push({ code: code || '', name: name, unit: unit || '', sklad: sklad || '', qty: 0, price: '' });
    } else {
      var it = list[index[name]];
      if (code) it.code = code;
      if (unit) it.unit = unit;
      if (sklad) it.sklad = sklad;
    }
    return list[index[name]];
  }
  var startSheet = ss.getSheetByName('СТАРТ');
  if (startSheet && startSheet.getLastRow() >= 3) {
    var sd = startSheet.getRange(3, 2, startSheet.getLastRow() - 2, 5).getValues();
    for (var s = 0; s < sd.length; s++) {
      var it = ensure(sd[s][0], sd[s][1], sd[s][2], '');
      if (!it) continue;
      var q = Number(sd[s][3]) || 0;
      it.qty += q;
      if (sd[s][4] !== '' && sd[s][4] !== null && Number(sd[s][4]) > 0) it.price = Number(sd[s][4]);
    }
  }
  var d = ss.getSheetByName('ДВИЖЕНИЕ');
  var dvizhHeader = typeof DVIZH_HEADER_ROW !== 'undefined' ? DVIZH_HEADER_ROW : 4;
  if (d && d.getLastRow() > dvizhHeader) {
    var dd = d.getRange(dvizhHeader + 1, 2, d.getLastRow() - dvizhHeader, 10).getValues();
    for (var j = 0; j < dd.length; j++) {
      var tip = String(dd[j][0] || '').trim();
      var it2 = ensure(dd[j][3], dd[j][4], dd[j][5], dd[j][9]);
      if (!it2) continue;
      var q2 = Number(dd[j][6]) || 0;
      if (tip === 'Расход') it2.qty -= q2;
      else if (tip === 'Приход') {
        it2.qty += q2;
        if (dd[j][7] !== '' && dd[j][7] !== null && Number(dd[j][7]) > 0) it2.price = Number(dd[j][7]);
      }
    }
  }
  var out = [];
  var colors = [];
  for (var k = 0; k < list.length; k++) {
    var x = list[k];
    x.qty = Math.round(x.qty * 1000) / 1000;
    var sum = (x.price === '' || x.price === null) ? '' : Math.round(x.qty * Number(x.price) * 100) / 100;
    out.push([x.code, x.name, x.unit, x.sklad, x.qty, x.price === '' ? '' : x.price, sum]);
    colors.push([x.qty < 0 ? '#ffcdd2' : '#ffffff']);
  }
  if (last >= dataStart) ost.getRange(dataStart, 1, last - ostHeader, 9).clearContent();
  if (out.length) {
    ost.getRange(dataStart, 1, out.length, 7).setValues(out);
    ost.getRange(dataStart, 5, out.length, 1).setBackgrounds(colors);
  }
  fillOstankiSearchCols_(ost);
  try { refreshOrderList_(); } catch (ignoreOrder) {}
  Logger.log('ОСТАТКИ пересчитаны: ' + out.length);
  try {
    SpreadsheetApp.getUi().alert('Остатки пересчитаны: ' + out.length + ' позиций. Кол-во = старт + приходы − расходы.');
  } catch (e2) {}
}

/** Лист категории → короткое имя для поиска. Код материала лежит в колонке B. */
var OST_CATS = [
  ['1. Расх_мат', 'Расходные материалы'],
  ['2. Лист_мат', 'Листовые'],
  ['3. Обрезки_лист', 'Обрезки листовых'],
  ['4. Рулон_мат', 'Рулонные'],
  ['5. Погонаж', 'Погонаж'],
  ['6. Хоз_мат', 'Хозматериалы'],
  ['7. Борта_алюм', 'Борта алюминиевые'],
  ['8. Заготовки_из_обрезк.', 'Заготовки из обрезков'],
  ['9. Расходники_станки', 'Расходники станков'],
  ['10. Световые_рамки', 'Световые рамки'],
  ['11. Сувенирка', 'Сувенирка']
];

function ostHeaderRow_(ost) {
  var n = Math.min(15, ost.getMaxRows());
  var vals = ost.getRange(1, 1, n, 2).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === 'Код' && String(vals[i][1]).trim() === 'Наименование') return i + 1;
  }
  return 2;
}

function ostCategoryMap_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var map = {};
  for (var i = 0; i < OST_CATS.length; i++) {
    var sh = ss.getSheetByName(OST_CATS[i][0]);
    if (!sh) continue;
    var last = sh.getLastRow();
    if (last < 1) continue;
    var col = sh.getRange(1, 2, Math.min(last, 4000), 1).getValues();
    for (var r = 0; r < col.length; r++) {
      var key = String(col[r][0] == null ? '' : col[r][0]).trim();
      if (!/^\d+$/.test(key)) continue;
      if (!map[key]) map[key] = OST_CATS[i][1];
    }
  }
  return map;
}

function ostCategoryForCode_(code) {
  var key = String(code == null ? '' : code).trim();
  if (!/^\d+$/.test(key)) return '';
  var map = ostCategoryMap_();
  return map[key] || '';
}

/** Колонка H — категория, I — «код + название» для поиска. I скрыта. */
function fillOstankiSearchCols_(ost) {
  var header = ostHeaderRow_(ost);
  var start = header + 1;
  var last = ost.getLastRow();
  if (last < start) return;
  var map = ostCategoryMap_();
  var n = last - header;
  var ab = ost.getRange(start, 1, n, 2).getValues();
  var cats = [];
  var blob = [];
  for (var i = 0; i < ab.length; i++) {
    var code = String(ab[i][0] == null ? '' : ab[i][0]).trim();
    var name = String(ab[i][1] || '').trim();
    cats.push([map[code] || '']);
    blob.push([code + ' ' + name]);
  }
  ost.getRange(header, 8).setValue('Категория');
  ost.getRange(start, 8, n, 1).setValues(cats);
  ost.getRange(start, 9, n, 1).setValues(blob);
  ost.hideColumns(9);
}

function setupOstankiSearch() {
  var ost = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ОСТАТКИ');
  if (!ost) throw new Error('Нет листа ОСТАТКИ');
  var fresh = String(ost.getRange(1, 1).getValue() || '').indexOf('ПОИСК ПО СКЛАДУ') < 0;
  if (fresh) ost.insertRowsBefore(1, 3);
  paintOstankiSearch_(ost, fresh);
  fillOstankiSearchCols_(ost);
  var filter = ost.getFilter();
  if (filter) filter.remove();
  var header = ostHeaderRow_(ost);
  var last = Math.max(ost.getLastRow(), header + 1);
  ost.getRange(header, 1, last - header + 1, 9).createFilter();
  ostApplyFilter_(ost);
  ost.setFrozenRows(header);
  ost.activate();
  try {
    SpreadsheetApp.getUi().alert('Поиск закреплён сверху листа ОСТАТКИ. Выбери категорию или впиши слово и нажми Enter.');
  } catch (e) {}
}

function paintOstankiSearch_(ost, fresh) {
  ost.getRange(1, 1, 4, 9).breakApart();
  ost.getRange(1, 1, 1, 8).merge()
    .setValue('ПОИСК ПО СКЛАДУ')
    .setBackground('#1a3a2a').setFontColor('#ffffff').setFontWeight('bold').setFontSize(14)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  ost.getRange(2, 1, 1, 8).setBackground('#e8f5e9').setFontWeight('bold').setFontSize(10).setFontColor('#1b5e20');
  ost.getRange(2, 1).setValue('Категория');
  ost.getRange(2, 2, 1, 3).merge().setValue('Название или код');
  ost.getRange(2, 5).setValue('Наличие');
  ost.getRange(2, 6).setValue('Найти').setHorizontalAlignment('center').setBackground('#2e7d32').setFontColor('#ffffff');
  ost.getRange(2, 7).setValue('Весь список').setHorizontalAlignment('center').setBackground('#546e7a').setFontColor('#ffffff');
  ost.getRange(2, 8).setValue('Нашлось').setHorizontalAlignment('center');
  var cats = ['Все категории'];
  for (var i = 0; i < OST_CATS.length; i++) cats.push(OST_CATS[i][1]);
  ost.getRange(3, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(cats, true).setAllowInvalid(false).build());
  ost.getRange(3, 5).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Все', 'Только в наличии', 'Нет на складе'], true).setAllowInvalid(false).build());
  if (!ost.getRange(3, 1).getValue()) ost.getRange(3, 1).setValue('Все категории');
  if (!ost.getRange(3, 5).getValue()) ost.getRange(3, 5).setValue('Все');
  ost.getRange(3, 2, 1, 3).merge();
  ost.getRange(3, 1, 1, 5).setBackground('#fff8e1').setFontSize(12).setVerticalAlignment('middle').setBorder(true, true, true, true, false, false, '#f9a825', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  ost.getRange(3, 6).insertCheckboxes().setBackground('#e8f5e9').setHorizontalAlignment('center');
  ost.getRange(3, 7).insertCheckboxes().setBackground('#eceff1').setHorizontalAlignment('center');
  ost.getRange(3, 8).setBackground('#ffffff').setFontSize(16).setFontWeight('bold').setFontColor('#1a3a2a').setHorizontalAlignment('center');
  ost.getRange(4, 1, 1, 8).merge()
    .setValue('Впиши часть названия или код и нажми Enter. Категория и «Наличие» срабатывают сразу. Галочка «Найти» — обновить. «Весь список» — показать всё. «Только в наличии» — где количество больше нуля.')
    .setBackground('#f1f8e9').setFontSize(10).setFontColor('#33691e').setWrap(true).setVerticalAlignment('middle');
  ost.setRowHeight(1, 32);
  ost.setRowHeight(2, 22);
  ost.setRowHeight(3, 36);
  ost.setRowHeight(4, 36);
  if (fresh) {
    ost.setColumnWidth(1, 220);
    ost.setColumnWidth(2, 280);
    ost.setColumnWidth(5, 170);
    ost.setColumnWidth(6, 90);
    ost.setColumnWidth(7, 110);
    ost.setColumnWidth(8, 200);
  }
  var header = ostHeaderRow_(ost);
  ost.getRange(header, 1, 1, 8).setBackground('#d9ead3').setFontWeight('bold');
  ost.setFrozenRows(header);
}

function ostSearchOnEdit_(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    if (sh.getName() !== 'ОСТАТКИ') return;
    if (String(sh.getRange(1, 1).getValue() || '').indexOf('ПОИСК ПО СКЛАДУ') < 0) return;
    if (e.range.getRow() !== 3) return;
    var col = e.range.getColumn();
    if (col !== 1 && col !== 2 && col !== 5 && col !== 6 && col !== 7) return;
    if (col === 6 || col === 7) {
      if (e.range.getValue() !== true) return;
      if (col === 7) {
        sh.getRange(3, 1).setValue('Все категории');
        sh.getRange(3, 2).clearContent();
        sh.getRange(3, 5).setValue('Все');
      }
      sh.getRange(3, 6).setValue(false);
      sh.getRange(3, 7).setValue(false);
    }
    ostApplyFilter_(sh);
  } catch (err) {
    Logger.log(err);
  }
}

function ostApplyFilter_(ost) {
  var cat = String(ost.getRange(3, 1).getValue() || '').trim();
  var q = String(ost.getRange(3, 2).getValue() || '').trim();
  var stock = String(ost.getRange(3, 5).getValue() || '').trim();
  var header = ostHeaderRow_(ost);
  var last = ost.getLastRow();
  if (last <= header) {
    ost.getRange(3, 8).setValue(0);
    return;
  }
  var filter = ost.getFilter();
  if (!filter || filter.getRange().getRow() !== header || filter.getRange().getLastRow() < last) {
    if (filter) filter.remove();
    filter = ost.getRange(header, 1, last - header + 1, 9).createFilter();
  }
  if (!cat || cat === 'Все категории') filter.setColumnFilterCriteria(8, null);
  else filter.setColumnFilterCriteria(8, SpreadsheetApp.newFilterCriteria().whenTextEqualTo(cat).build());
  if (!q) filter.setColumnFilterCriteria(9, null);
  else filter.setColumnFilterCriteria(9, SpreadsheetApp.newFilterCriteria().whenTextContains(q).build());
  if (stock === 'Только в наличии') filter.setColumnFilterCriteria(5, SpreadsheetApp.newFilterCriteria().whenNumberGreaterThan(0).build());
  else if (stock === 'Нет на складе') filter.setColumnFilterCriteria(5, SpreadsheetApp.newFilterCriteria().whenNumberLessThanOrEqualTo(0).build());
  else filter.setColumnFilterCriteria(5, null);
  var data = ost.getRange(header + 1, 5, last - header, 5).getValues();
  var n = 0;
  var ql = q.toLowerCase();
  for (var i = 0; i < data.length; i++) {
    var nameBlob = String(data[i][4] || '');
    if (!nameBlob.trim()) continue;
    if (cat && cat !== 'Все категории' && String(data[i][3] || '') !== cat) continue;
    if (ql && nameBlob.toLowerCase().indexOf(ql) < 0) continue;
    var qty = Number(data[i][0]) || 0;
    if (stock === 'Только в наличии' && !(qty > 0)) continue;
    if (stock === 'Нет на складе' && !(qty <= 0)) continue;
    n++;
  }
  ost.getRange(3, 8).setValue(n);
}

/** Лист ходовых. Имя содержит «К ЗАКАЗУ», закладка красная. */
function orderSheet_() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (String(sheets[i].getName()).indexOf('К ЗАКАЗУ') >= 0) return sheets[i];
  }
  return null;
}

function setupOrderSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = orderSheet_();
  if (!sh) sh = ss.insertSheet('🔴 К ЗАКАЗУ');
  if (sh.getName() !== '🔴 К ЗАКАЗУ') sh.setName('🔴 К ЗАКАЗУ');
  sh.setTabColor('#d50000');
  var ost = ss.getSheetByName('ОСТАТКИ');
  if (ost) {
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(ost.getIndex() + 1);
  }
  if (String(sh.getRange(2, 1).getValue() || '') !== 'Код') {
    sh.getRange(1, 1, 2, 7).breakApart();
    sh.getRange(1, 1, 1, 7).merge()
      .setValue('Ходовые. Выбери название из списка и впиши минимум. Красная строка — пора заказать. Жёлтая — уже близко.')
      .setBackground('#b71c1c').setFontColor('#ffffff').setFontWeight('bold').setFontSize(12)
      .setVerticalAlignment('middle').setWrap(true);
    sh.setRowHeight(1, 36);
    sh.getRange(2, 1, 1, 7).setValues([['Код', 'Наименование', 'Ед.', 'Минимум', 'Сейчас', 'Докупить', 'Статус']])
      .setBackground('#3e2723').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
    sh.setFrozenRows(2);
    sh.setColumnWidth(1, 90);
    sh.setColumnWidth(2, 340);
    sh.setColumnWidth(3, 70);
    sh.setColumnWidth(4, 110);
    sh.setColumnWidth(5, 100);
    sh.setColumnWidth(6, 110);
    sh.setColumnWidth(7, 150);
    sh.getRange(3, 1, 20, 7).setBackground('#ffffff').setVerticalAlignment('middle');
    sh.getRange(3, 4, 20, 1).setBackground('#fff8e1');
    sh.getRange('A3:G22').setBorder(true, true, true, true, true, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
  }
  var base = ss.getSheetByName('База_материалы');
  if (base && base.getLastRow() >= 2) {
    sh.getRange('B3:B80').setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInRange(base.getRange(2, 2, base.getLastRow() - 1, 1), true)
        .setAllowInvalid(true)
        .build()
    );
  }
  var filter = sh.getFilter();
  if (filter) filter.remove();
  sh.getRange(2, 1, Math.max(sh.getLastRow(), 22) - 1, 7).createFilter();
  refreshOrderList_();
  sh.activate();
  try {
    SpreadsheetApp.getUi().alert('Лист «К ЗАКАЗУ» готов, закладка красная. Выбери название и впиши минимум. Красная строка — пора заказать.');
  } catch (e) {}
}

function orderOnEdit_(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    if (String(sh.getName()).indexOf('К ЗАКАЗУ') < 0) return;
    var col = e.range.getColumn();
    if (col !== 2 && col !== 4) return;
    var startRow = e.range.getRow();
    var num = e.range.getNumRows();
    if (col === 2) {
      for (var r = startRow; r < startRow + num; r++) {
        if (r < 3) continue;
        var name = String(sh.getRange(r, 2).getValue() || '').trim();
        if (!name) {
          sh.getRange(r, 1, 1, 7).clearContent().setBackground('#ffffff');
          sh.getRange(r, 4).setBackground('#fff8e1');
          continue;
        }
        var mat = lookupMaterial_(name);
        sh.getRange(r, 1).setValue(mat.code || '');
        sh.getRange(r, 3).setValue(mat.unit || '');
      }
    }
    refreshOrderList_();
  } catch (err) {
    Logger.log(err);
  }
}

function refreshOrderList_() {
  var sh = orderSheet_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ost = ss.getSheetByName('ОСТАТКИ');
  if (!ost) return;
  var watch = {};
  var redN = 0;
  if (sh) {
    sh.setTabColor('#d50000');
    var last = Math.max(sh.getLastRow(), 2);
    if (last >= 3) {
      var stockIndex = ostStockIndex_(ost);
      var rows = sh.getRange(3, 1, last - 2, 7).getValues();
      for (var i = 0; i < rows.length; i++) {
        var name = String(rows[i][1] || '').trim();
        if (!name) continue;
        var stock = stockIndex[name] || null;
        var minRaw = rows[i][3];
        var minEmpty = (minRaw === '' || minRaw === null);
        var min = Number(minRaw);
        var qty = stock ? stock.qty : 0;
        var status = '';
        var buy = '';
        var bg = '#ffffff';
        var level = '';
        if (!stock) {
          status = 'Нет в остатках';
          bg = '#eeeeee';
        } else if (minEmpty || isNaN(min)) {
          status = 'Укажи минимум';
        } else if (qty <= min) {
          status = 'ЗАКАЗАТЬ';
          buy = Math.round((min - qty) * 1000) / 1000;
          bg = '#ffcdd2';
          level = 'red';
          redN++;
        } else if (qty <= min * 1.5) {
          status = 'Скоро';
          bg = '#fff9c4';
          level = 'yellow';
        } else {
          status = 'Норма';
          bg = '#c8e6c9';
          level = 'ok';
        }
        var now = stock ? qty : '';
        sh.getRange(3 + i, 5, 1, 3).setValues([[now, buy, status]]);
        sh.getRange(3 + i, 1, 1, 7).setBackground(bg);
        if (!stock || minEmpty) sh.getRange(3 + i, 4).setBackground('#fff8e1');
        if (level === 'red') sh.getRange(3 + i, 7).setFontColor('#b71c1c').setFontWeight('bold');
        else sh.getRange(3 + i, 7).setFontColor('#333333').setFontWeight('bold');
        if (stock) watch[name] = level;
        if (stock && !rows[i][0]) sh.getRange(3 + i, 1).setValue(stock.code);
        if (stock && !rows[i][2]) sh.getRange(3 + i, 3).setValue(stock.unit);
      }
    }
  }
  paintOstankiOrderMarks_(ost, watch, redN);
}

function ostStockIndex_(ost) {
  var header = ostHeaderRow_(ost);
  var last = ost.getLastRow();
  var map = {};
  if (last <= header) return map;
  var data = ost.getRange(header + 1, 1, last - header, 5).getValues();
  for (var i = 0; i < data.length; i++) {
    var nm = String(data[i][1] || '').trim();
    if (!nm || map[nm]) continue;
    map[nm] = { code: data[i][0] || '', unit: data[i][2] || '', qty: Number(data[i][4]) || 0 };
  }
  return map;
}

function paintOstankiOrderMarks_(ost, watch, redN) {
  var header = ostHeaderRow_(ost);
  var last = ost.getLastRow();
  if (last > header) {
    var names = ost.getRange(header + 1, 2, last - header, 1).getValues();
    var qtys = ost.getRange(header + 1, 5, last - header, 1).getValues();
    var bgs = [];
    for (var i = 0; i < names.length; i++) {
      var nm = String(names[i][0] || '').trim();
      var level = watch[nm] || '';
      var qty = Number(qtys[i][0]) || 0;
      var color = '#ffffff';
      if (level === 'red') color = '#ffcdd2';
      else if (level === 'yellow') color = '#fff9c4';
      else if (qty < 0) color = '#ffcdd2';
      bgs.push([color]);
    }
    ost.getRange(header + 1, 5, bgs.length, 1).setBackgrounds(bgs);
  }
  ost.getRange(2, 10).setValue('К заказу')
    .setBackground('#b71c1c').setFontColor('#ffffff').setFontWeight('bold')
    .setFontSize(10).setHorizontalAlignment('center');
  ost.getRange(3, 10).setValue(redN)
    .setBackground(redN > 0 ? '#ffcdd2' : '#ffffff')
    .setFontColor('#b71c1c').setFontWeight('bold').setFontSize(16)
    .setHorizontalAlignment('center');
  ost.setColumnWidth(10, 110);
}

/** Лист РЕЗЕРВ: постановка снимает количество с ОСТАТКИ, снятие возвращает. */
function reserveSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName('РЕЗЕРВ');
}

function setupReserveSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = reserveSheet_();
  if (!sh) sh = ss.insertSheet('РЕЗЕРВ');
  sh.setTabColor('#e65100');
  var after = orderSheet_() || ss.getSheetByName('ОСТАТКИ');
  if (after) {
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(after.getIndex() + 1);
  }
  widenManagersColumn_();
  if (String(sh.getRange(5, 1).getValue() || '') !== 'Дата') {
    sh.getRange(1, 1, 5, 8).breakApart();
    sh.getRange(1, 1, 1, 8).merge()
      .setValue('РЕЗЕРВ. Выбери материал и сотрудника, укажи количество и поставь галочку. Это количество уйдёт из остатков.')
      .setBackground('#e65100').setFontColor('#ffffff').setFontWeight('bold').setFontSize(13)
      .setVerticalAlignment('middle').setWrap(true);
    sh.setRowHeight(1, 36);
    sh.getRange(2, 1, 1, 8).setBackground('#fff3e0').setFontWeight('bold').setFontSize(10).setFontColor('#e65100');
    sh.getRange(2, 1, 1, 3).merge().setValue('Материал');
    sh.getRange(2, 4).setValue('Количество');
    sh.getRange(2, 5).setValue('Сотрудник');
    sh.getRange(2, 6).setValue('В резерв').setHorizontalAlignment('center').setBackground('#e65100').setFontColor('#ffffff');
    sh.getRange(2, 7, 1, 2).merge().setValue('На складе');
    sh.getRange(3, 1, 1, 3).merge();
    sh.getRange(3, 1, 1, 5).setBackground('#fff8e1').setFontSize(12).setVerticalAlignment('middle')
      .setBorder(true, true, true, true, false, false, '#f9a825', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sh.getRange(3, 6).insertCheckboxes().setBackground('#fff3e0').setHorizontalAlignment('center');
    sh.getRange(3, 7, 1, 2).merge().setBackground('#ffffff').setFontWeight('bold').setFontSize(12).setVerticalAlignment('middle');
    sh.setRowHeight(3, 36);
    sh.getRange(4, 1, 1, 8).merge()
      .setValue('Если сотрудника нет в списке — напиши фамилию и имя. Он сам добавится в лист База, колонка «Менеджеры». Галочка «Снять» в строке внизу возвращает количество на склад.')
      .setBackground('#fff8e1').setFontSize(10).setFontColor('#e65100').setWrap(true).setVerticalAlignment('middle');
    sh.setRowHeight(4, 36);
    sh.getRange(5, 1, 1, 8).setValues([['Дата', 'Материал', 'Код', 'Ед.', 'Кол-во', 'Сотрудник', 'Статус', 'Снять']])
      .setBackground('#3e2723').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
    sh.setFrozenRows(5);
    sh.setColumnWidth(1, 280);
    sh.setColumnWidth(4, 120);
    sh.setColumnWidth(5, 220);
    sh.setColumnWidth(6, 100);
    sh.setColumnWidth(7, 180);
    sh.setColumnWidth(8, 80);
  }
  var mats = ss.getSheetByName('База_материалы');
  if (mats && mats.getLastRow() >= 2) {
    sh.getRange(3, 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInRange(mats.getRange(2, 2, mats.getLastRow() - 1, 1), true)
        .setAllowInvalid(false)
        .build()
    );
  }
  var base = ss.getSheetByName('База');
  if (base) {
    sh.getRange(3, 5).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInRange(base.getRange('J2:J200'), true)
        .setAllowInvalid(true)
        .build()
    );
  }
  var filter = sh.getFilter();
  if (filter) filter.remove();
  var last = Math.max(sh.getLastRow(), 5);
  sh.getRange(5, 1, last - 4, 8).createFilter();
  sh.activate();
  try {
    SpreadsheetApp.getUi().alert('Лист «РЕЗЕРВ» готов. Выбери материал и сотрудника, укажи количество и поставь галочку. Количество уйдёт из остатков. «Снять» — вернёт обратно.');
  } catch (e) {}
}

function widenManagersColumn_() {
  var base = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('База');
  if (!base) return;
  base.setColumnWidth(10, 260);
  if (String(base.getRange(1, 10).getValue() || '') !== 'Менеджеры') {
    base.getRange(1, 10).setValue('Менеджеры');
  }
  base.getRange(1, 10).setFontWeight('bold');
  var slots = base.getRange(11, 10, 40, 1).getValues();
  var bgs = [];
  for (var i = 0; i < slots.length; i++) bgs.push([String(slots[i][0] || '').trim() ? '#ffffff' : '#fff8e1']);
  base.getRange(11, 10, 40, 1).setBackgrounds(bgs);
}

function ensureManager_(name) {
  name = String(name || '').trim();
  if (!name) return false;
  var base = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('База');
  if (!base) return false;
  var last = Math.max(base.getLastRow(), 2);
  var vals = base.getRange(2, 10, last - 1, 1).getValues();
  var key = name.toLowerCase();
  var emptyRow = 0;
  for (var i = 0; i < vals.length; i++) {
    var cur = String(vals[i][0] || '').trim();
    if (cur.toLowerCase() === key) return false;
    if (!cur && !emptyRow) emptyRow = 2 + i;
  }
  var row = emptyRow || (last + 1);
  base.getRange(row, 10).setValue(name).setBackground('#ffffff');
  return true;
}

function reserveOnEdit_(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    if (sh.getName() !== 'РЕЗЕРВ') return;
    var row = e.range.getRow();
    var col = e.range.getColumn();
    if (row === 3 && (col === 1 || col === 4)) {
      showReserveStock_(sh);
      return;
    }
    if (row === 3 && col === 6) {
      if (e.range.getValue() !== true) return;
      placeReserve_(sh);
      return;
    }
    if (row >= 6 && col === 8) {
      if (e.range.getValue() !== true) return;
      releaseReserve_(sh, row);
    }
  } catch (err) {
    Logger.log(err);
  }
}

function showReserveStock_(sh) {
  var name = String(sh.getRange(3, 1).getValue() || '').trim();
  if (!name) {
    sh.getRange(3, 7).setValue('');
    return;
  }
  var stock = reserveStock_(name);
  var unit = stock.unit ? (' ' + stock.unit) : '';
  sh.getRange(3, 7).setValue(stock.found ? ('На складе: ' + stock.qty + unit) : 'Нет в остатках');
}

function reserveStock_(name) {
  var ost = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ОСТАТКИ');
  var mat = lookupMaterial_(name);
  if (!ost) return { found: false, qty: 0, unit: mat.unit || '', mat: mat };
  var map = ostStockIndex_(ost);
  var hit = map[String(name).trim()];
  if (!hit) return { found: false, qty: 0, unit: mat.unit || '', mat: mat };
  return { found: true, qty: Number(hit.qty) || 0, unit: hit.unit || mat.unit || '', code: hit.code || mat.code, mat: mat };
}

function placeReserve_(sh) {
  var name = String(sh.getRange(3, 1).getValue() || '').trim();
  var qty = Number(sh.getRange(3, 4).getValue());
  var who = String(sh.getRange(3, 5).getValue() || '').trim();
  sh.getRange(3, 6).setValue(false);
  if (!name) {
    sh.getRange(3, 7).setValue('Выбери материал');
    return;
  }
  if (!(qty > 0)) {
    sh.getRange(3, 7).setValue('Укажи количество больше нуля');
    return;
  }
  if (!who) {
    sh.getRange(3, 7).setValue('Выбери или напиши сотрудника');
    return;
  }
  var stock = reserveStock_(name);
  if (!stock.found) {
    sh.getRange(3, 7).setValue('Такого материала нет в остатках');
    return;
  }
  if (stock.qty < qty) {
    sh.getRange(3, 7).setValue('На складе только ' + stock.qty + '. Резерв не поставлен');
    return;
  }
  var added = ensureManager_(who);
  var mat = stock.mat || lookupMaterial_(name);
  var code = stock.code || mat.code || '';
  var unit = stock.unit || mat.unit || '';
  writeReserveMove_('Расход', name, code, unit, mat.sklad || '', qty, who);
  var row = nextReserveRow_(sh);
  sh.getRange(row, 1, 1, 8).setValues([[new Date(), name, code, unit, qty, who, 'В резерве', false]]);
  sh.getRange(row, 1).setNumberFormat('dd.MM.yyyy');
  sh.getRange(row, 8).insertCheckboxes();
  sh.getRange(row, 1, 1, 8).setBackground('#ffe0b2');
  sh.getRange(row, 7).setFontWeight('bold').setFontColor('#e65100');
  sh.getRange(3, 4).clearContent();
  var left = Math.round((stock.qty - qty) * 1000) / 1000;
  var msg = 'В резерве ' + qty + (unit ? ' ' + unit : '') + '. На складе осталось ' + left;
  if (added) msg += '. Сотрудник добавлен в Базу';
  sh.getRange(3, 7).setValue(msg);
}

function releaseReserve_(sh, row) {
  var status = String(sh.getRange(row, 7).getValue() || '');
  if (status !== 'В резерве') {
    sh.getRange(row, 8).setValue(false);
    return;
  }
  var name = String(sh.getRange(row, 2).getValue() || '').trim();
  var code = sh.getRange(row, 3).getValue();
  var unit = sh.getRange(row, 4).getValue();
  var qty = Number(sh.getRange(row, 5).getValue());
  var who = String(sh.getRange(row, 6).getValue() || '').trim();
  if (!name || !(qty > 0)) {
    sh.getRange(row, 8).setValue(false);
    return;
  }
  var mat = lookupMaterial_(name);
  writeReserveMove_('Приход', name, code || mat.code, unit || mat.unit, mat.sklad || '', qty, who);
  sh.getRange(row, 7).setValue('Снят').setFontColor('#616161');
  sh.getRange(row, 8).setValue(false);
  sh.getRange(row, 1, 1, 8).setBackground('#eeeeee');
  sh.getRange(3, 7).setValue('Резерв снят, ' + qty + ' вернулось на склад');
}

function nextReserveRow_(sh) {
  var last = sh.getLastRow();
  if (last < 6) return 6;
  var names = sh.getRange(6, 2, last - 5, 1).getValues();
  for (var i = 0; i < names.length; i++) if (!names[i][0]) return 6 + i;
  return last + 1;
}

function writeReserveMove_(tip, name, code, unit, sklad, qty, who) {
  var d = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ДВИЖЕНИЕ');
  if (!d) throw new Error('Нет листа ДВИЖЕНИЕ');
  var row = nextDvizhenieRow_(d);
  var comment = (tip === 'Расход' ? 'Резерв' : 'Снятие резерва') + ' · ' + who;
  d.getRange(row, 1, 1, 12).setValues([[
    new Date(), tip, 'Резерв', '', name, code || '', unit || '', qty, '', '', sklad || '', comment
  ]]);
  d.getRange(row, 2).setBackground(tip === 'Приход' ? '#c8e6c9' : '#ffe0b2');
  applyOstankiFromMove_(name, code || '', unit || '', sklad || '', tip, qty, '');
}
