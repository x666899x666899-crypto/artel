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
  var last = Math.max(ost.getLastRow(), 2);
  var names = last >= 3 ? ost.getRange(3, 2, last - 2, 1).getValues() : [];
  var row = 0;
  for (var i = 0; i < names.length; i++) {
    if (String(names[i][0]).trim() === key) { row = 3 + i; break; }
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
  return { row: row, qty: newQty };
}

/** Полный пересчёт ОСТАТКИ: ноль + СТАРТ + все строки ДВИЖЕНИЕ. */
function rebuildOstanki() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ost = ss.getSheetByName('ОСТАТКИ');
  if (!ost) throw new Error('Нет листа ОСТАТКИ');
  var last = ost.getLastRow();
  var rows = last >= 3 ? ost.getRange(3, 1, last - 2, 7).getValues() : [];
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
  var start = ss.getSheetByName('СТАРТ');
  if (start && start.getLastRow() >= 3) {
    var sd = start.getRange(3, 2, start.getLastRow() - 2, 5).getValues();
    for (var s = 0; s < sd.length; s++) {
      var it = ensure(sd[s][0], sd[s][1], sd[s][2], '');
      if (!it) continue;
      var q = Number(sd[s][3]) || 0;
      it.qty += q;
      if (sd[s][4] !== '' && sd[s][4] !== null && Number(sd[s][4]) > 0) it.price = Number(sd[s][4]);
    }
  }
  var d = ss.getSheetByName('ДВИЖЕНИЕ');
  var header = typeof DVIZH_HEADER_ROW !== 'undefined' ? DVIZH_HEADER_ROW : 4;
  if (d && d.getLastRow() > header) {
    var dd = d.getRange(header + 1, 2, d.getLastRow() - header, 10).getValues();
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
  if (last >= 3) ost.getRange(3, 1, last - 2, 7).clearContent();
  if (out.length) {
    ost.getRange(3, 1, out.length, 7).setValues(out);
    ost.getRange(3, 5, out.length, 1).setBackgrounds(colors);
  }
  Logger.log('ОСТАТКИ пересчитаны: ' + out.length);
  try {
    SpreadsheetApp.getUi().alert('Остатки пересчитаны: ' + out.length + ' позиций. Кол-во = старт + приходы − расходы.');
  } catch (e2) {}
}
