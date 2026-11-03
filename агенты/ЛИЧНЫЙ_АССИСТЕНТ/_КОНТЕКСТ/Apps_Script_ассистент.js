/**
 * Личный ассистент Сергея — Telegram → задачи и напоминания.
 *
 * 1) Впиши SPREADSHEET_ID своей таблицы «Личный ассистент».
 * 2) setupAssistant() → installPolling() → installReminders().
 * 3) Токен: свойства скрипта TELEGRAM_TOKEN.
 * Не хранить токены в этом файле.
 *
 * Отдельный бот от «Личные финансы» — один бот не тянет два опроса сразу.
 */

var SPREADSHEET_ID = '1pHvJ3Y-dh4endoZK2gxKj499Os36G54AOxb96SeRJyE';
var TZ = 'Asia/Irkutsk';

function getSS_() {
  if (!SPREADSHEET_ID || SPREADSHEET_ID.indexOf('ВСТАВЬ') === 0) {
    throw new Error('Впиши SPREADSHEET_ID в начале скрипта (из URL таблицы)');
  }
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

var HELP_TEXT =
  'Задача: + позвонить Василию\n' +
  'Список: список\n' +
  'Закрыть: готово 3\n' +
  'Напоминание: напомни завтра 9:00 текст\n' +
  'напомни 05.09 14:30 текст\n' +
  'напомни через 20 минут текст\n' +
  'План / утро — сводка\n' +
  '/help — это меню';

function setupAssistant() {
  var ss = getSS_();
  ss.rename('Личный ассистент');

  ensureSheet_(ss, 'Входящее', [
    'Дата_время', 'Тип', 'Текст', 'Ссылка_файл', 'Telegram_id', 'Chat_id', 'Статус', 'Комментарий'
  ]);
  ensureSheet_(ss, 'Задачи', [
    'id', 'Создано', 'Текст', 'Срок', 'Приоритет', 'Статус', 'Источник', 'Chat_id', 'Комментарий'
  ]);
  ensureSheet_(ss, 'Напоминания', [
    'id', 'Когда', 'Текст', 'Статус', 'Задача_id', 'Chat_id'
  ]);
  ensureSheet_(ss, 'План_дня', [
    'Дата', 'План', 'Создано', 'Источник'
  ]);
  ensureSheet_(ss, 'Справка', ['Как пользоваться']);

  var help = ss.getSheetByName('Справка');
  help.getRange('A2').setValue(HELP_TEXT.replace(/\n/g, ' | '));
}

function ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
}

function registerWebhook() {
  var token = PropertiesService.getScriptProperties().getProperty('TELEGRAM_TOKEN');
  if (!token) throw new Error('Нет TELEGRAM_TOKEN');
  var url = ScriptApp.getService().getUrl();
  if (!url) throw new Error('Сначала разверни веб-приложение');
  var res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/setWebhook', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ url: url })
  });
  Logger.log(url);
  Logger.log(res.getContentText());
}

function stopWebhook() {
  var token = PropertiesService.getScriptProperties().getProperty('TELEGRAM_TOKEN');
  if (!token) throw new Error('Нет TELEGRAM_TOKEN');
  var res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/deleteWebhook', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ drop_pending_updates: true })
  });
  Logger.log(res.getContentText());
}

function pollTelegram() {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('TELEGRAM_TOKEN');
  if (!token) return;
  var offset = Number(props.getProperty('TG_OFFSET') || '0');
  var url = 'https://api.telegram.org/bot' + token + '/getUpdates?timeout=0&offset=' + offset;
  var data = JSON.parse(UrlFetchApp.fetch(url).getContentText());
  if (!data.ok) return;
  var i;
  for (i = 0; i < data.result.length; i++) {
    var upd = data.result[i];
    props.setProperty('TG_OFFSET', String(upd.update_id + 1));
    var msg = upd.message || upd.edited_message;
    if (!msg) continue;
    try {
      handleMsg_(token, props, msg);
    } catch (err) {
      Logger.log(err);
      try {
        reply_(token, String(msg.chat.id), 'Ошибка: ' + err.message);
      } catch (e2) {}
    }
  }
}

function installPolling() {
  stopWebhook();
  clearTriggers_('pollTelegram');
  ScriptApp.newTrigger('pollTelegram').timeBased().everyMinutes(1).create();
  pollTelegram();
}

function installReminders() {
  clearTriggers_('checkReminders');
  clearTriggers_('morningPush');
  ScriptApp.newTrigger('checkReminders').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('morningPush').timeBased().everyMinutes(15).create();
  checkReminders();
}

function clearTriggers_(fnName) {
  var triggers = ScriptApp.getProjectTriggers();
  var i;
  for (i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === fnName) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function doPost(e) {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('TELEGRAM_TOKEN');
  if (!token) return textOut_('no token');
  try {
    var update = JSON.parse(e.postData.contents);
    var msg = update.message || update.edited_message;
    if (!msg) return textOut_('ok');
    handleMsg_(token, props, msg);
  } catch (err) {
    Logger.log(err);
  }
  return textOut_('ok');
}

function handleMsg_(token, props, msg) {
  var chatId = String(msg.chat.id);
  var allowed = props.getProperty('ALLOWED_CHAT_ID');
  if (allowed && chatId !== String(allowed)) {
    reply_(token, chatId, 'Этот бот только для владельца.');
    return;
  }
  if (!allowed) props.setProperty('ALLOWED_CHAT_ID', chatId);

  var text = (msg.text || msg.caption || '').trim();
  var hasFile = !!(msg.photo || msg.voice || msg.audio || msg.document);
  logIncoming_(msg, text, hasFile);

  if (/^\/start(@[A-Za-z0-9_]+)?$/i.test(text)) {
    reply_(token, chatId, 'Ассистент на связи.\n' + HELP_TEXT);
    return;
  }
  if (/^\/help(@[A-Za-z0-9_]+)?$/i.test(text)) {
    reply_(token, chatId, HELP_TEXT);
    return;
  }

  if (!text && hasFile) {
    reply_(token, chatId, 'Файл во входящем. Напиши подпись: + задача или напомни …');
    return;
  }
  if (!text) return;

  var low = text.toLowerCase();

  if (/^(список|задачи)\b/.test(low)) {
    reply_(token, chatId, formatTaskList_());
    return;
  }
  if (/^(план)\b/.test(low)) {
    reply_(token, chatId, formatPlan_());
    return;
  }
  if (/^(утро|брифинг)\b/.test(low)) {
    reply_(token, chatId, formatMorning_());
    return;
  }

  var done = text.match(/^(готово|сделано|закрыть)\s+(\d+)\s*$/i);
  if (done) {
    reply_(token, chatId, closeTask_(Number(done[2])));
    return;
  }

  if (/^\+\s*/.test(text) || /^задача[:\s]/i.test(text)) {
    var taskText = text.replace(/^\+\s*/, '').replace(/^задача[:\s]+/i, '').trim();
    if (!taskText) {
      reply_(token, chatId, 'После + напиши задачу');
      return;
    }
    var id = addTask_(taskText, chatId, '');
    reply_(token, chatId, 'Задача #' + id + ':\n' + taskText);
    return;
  }

  if (/^напомни\b/i.test(text)) {
    reply_(token, chatId, addReminderFromText_(text, chatId));
    return;
  }

  // fallback: treat as task if starts with verb-ish short line, else help / AI
  if (props.getProperty('DEEPSEEK_API_KEY')) {
    reply_(token, chatId, askDeepSeek_(text));
    return;
  }
  reply_(token, chatId, 'Не понял. Добавь задачу так:\n+ текст\nили\n' + HELP_TEXT);
}

function logIncoming_(msg, text, hasFile) {
  var sh = getSS_().getSheetByName('Входящее');
  var now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
  var typ = hasFile ? 'файл' : 'текст';
  sh.appendRow([now, typ, text || '', '', String(msg.message_id || ''), String(msg.chat.id), 'новое', '']);
}

function nextId_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return 1;
  var vals = sh.getRange(2, 1, last - 1, 1).getValues();
  var max = 0;
  var i;
  for (i = 0; i < vals.length; i++) {
    var n = Number(vals[i][0]);
    if (n > max) max = n;
  }
  return max + 1;
}

function addTask_(text, chatId, due) {
  var sh = getSS_().getSheetByName('Задачи');
  var id = nextId_(sh);
  var now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
  sh.appendRow([id, now, text, due || '', 2, 'открыто', 'telegram', chatId, '']);
  return id;
}

function closeTask_(id) {
  var sh = getSS_().getSheetByName('Задачи');
  var last = sh.getLastRow();
  if (last < 2) return 'Нет задач';
  var data = sh.getRange(2, 1, last - 1, 6).getValues();
  var i;
  for (i = 0; i < data.length; i++) {
    if (Number(data[i][0]) === id) {
      sh.getRange(i + 2, 6).setValue('сделано');
      return 'Закрыл #' + id + ': ' + data[i][2];
    }
  }
  return 'Не нашёл задачу #' + id;
}

function openTasks_() {
  var sh = getSS_().getSheetByName('Задачи');
  var last = sh.getLastRow();
  if (last < 2) return [];
  var data = sh.getRange(2, 1, last - 1, 6).getValues();
  var out = [];
  var i;
  for (i = 0; i < data.length; i++) {
    if (String(data[i][5]).toLowerCase() === 'открыто') {
      out.push({ id: data[i][0], text: data[i][2], due: data[i][3], pri: data[i][4] });
    }
  }
  return out;
}

function formatTaskList_() {
  var tasks = openTasks_();
  if (!tasks.length) return 'Открытых задач нет. Добавь: + текст';
  var lines = ['Открыто (' + tasks.length + '):'];
  var i;
  for (i = 0; i < tasks.length; i++) {
    var t = tasks[i];
    var due = t.due ? ' · ' + t.due : '';
    lines.push('#' + t.id + ' ' + t.text + due);
  }
  return lines.join('\n');
}

function formatPlan_() {
  var tasks = formatTaskList_();
  var rem = upcomingReminders_(5);
  var lines = [tasks, '', 'Ближайшие напоминания:'];
  if (!rem.length) lines.push('(пусто)');
  else {
    var i;
    for (i = 0; i < rem.length; i++) {
      lines.push(rem[i].when + ' — ' + rem[i].text);
    }
  }
  return lines.join('\n');
}

function formatMorning_() {
  var tasks = openTasks_();
  var lines = ['Утро · ' + Utilities.formatDate(new Date(), TZ, 'dd.MM')];
  if (!tasks.length) {
    lines.push('Открытых задач нет.');
  } else {
    lines.push('Главное:');
    var n = Math.min(5, tasks.length);
    var i;
    for (i = 0; i < n; i++) {
      lines.push((i + 1) + '. ' + tasks[i].text);
    }
    if (tasks.length > 5) lines.push('… ещё ' + (tasks.length - 5));
  }
  var rem = upcomingReminders_(3);
  if (rem.length) {
    lines.push('');
    lines.push('Напоминания:');
    for (i = 0; i < rem.length; i++) {
      lines.push('· ' + rem[i].when + ' — ' + rem[i].text);
    }
  }
  return lines.join('\n');
}

function addReminderFromText_(text, chatId) {
  var parsed = parseReminder_(text);
  if (!parsed) {
    return 'Формат:\nнапомни завтра 9:00 текст\nнапомни 05.09 14:30 текст\nнапомни через 20 минут текст';
  }
  var sh = getSS_().getSheetByName('Напоминания');
  var id = nextId_(sh);
  var whenStr = Utilities.formatDate(parsed.when, TZ, 'yyyy-MM-dd HH:mm');
  sh.appendRow([id, whenStr, parsed.text, 'ждать', '', chatId]);
  return 'Напоминание #' + id + '\n' + whenStr + '\n' + parsed.text;
}

function parseReminder_(text) {
  var body = text.replace(/^напомни\s+/i, '').trim();
  var now = new Date();
  var when = null;
  var msg = '';

  var mRel = body.match(/^через\s+(\d+)\s*(минут[уыа]?|мин|час(?:а|ов)?|ч)\s+(.+)$/i);
  if (mRel) {
    var n = Number(mRel[1]);
    var unit = mRel[2].toLowerCase();
    var ms = /ч/.test(unit) ? n * 3600000 : n * 60000;
    when = new Date(now.getTime() + ms);
    msg = mRel[3].trim();
    return { when: when, text: msg };
  }

  var mTom = body.match(/^завтра\s+(\d{1,2})(?::(\d{2}))?\s+(.+)$/i);
  if (mTom) {
    when = dateInTz_(now, 1, Number(mTom[1]), Number(mTom[2] || 0));
    msg = mTom[3].trim();
    return { when: when, text: msg };
  }

  var mToday = body.match(/^сегодня\s+(\d{1,2})(?::(\d{2}))?\s+(.+)$/i);
  if (mToday) {
    when = dateInTz_(now, 0, Number(mToday[1]), Number(mToday[2] || 0));
    msg = mToday[3].trim();
    return { when: when, text: msg };
  }

  var mDate = body.match(/^(\d{1,2})[./](\d{1,2})(?:[./](\d{2,4}))?\s+(\d{1,2})(?::(\d{2}))?\s+(.+)$/);
  if (mDate) {
    var day = Number(mDate[1]);
    var month = Number(mDate[2]) - 1;
    var year = mDate[3] ? Number(mDate[3]) : Number(Utilities.formatDate(now, TZ, 'yyyy'));
    if (year < 100) year += 2000;
    var hour = Number(mDate[4]);
    var min = Number(mDate[5] || 0);
    when = makeTzDate_(year, month, day, hour, min);
    msg = mDate[6].trim();
    return { when: when, text: msg };
  }

  return null;
}

function dateInTz_(base, dayOffset, hour, minute) {
  var y = Number(Utilities.formatDate(base, TZ, 'yyyy'));
  var m = Number(Utilities.formatDate(base, TZ, 'MM')) - 1;
  var d = Number(Utilities.formatDate(base, TZ, 'dd')) + dayOffset;
  return makeTzDate_(y, m, d, hour, minute);
}

/** Приблизительно: собираем строку и парсим как локаль скрипта; TZ задан проектом. */
function makeTzDate_(year, month, day, hour, minute) {
  var s =
    year +
    '-' +
    pad2_(month + 1) +
    '-' +
    pad2_(day) +
    'T' +
    pad2_(hour) +
    ':' +
    pad2_(minute) +
    ':00';
  // Apps Script Date в проекте с TZ Asia/Irkutsk обычно ок для триггеров
  return new Date(s);
}

function pad2_(n) {
  return (n < 10 ? '0' : '') + n;
}

function upcomingReminders_(limit) {
  var sh = getSS_().getSheetByName('Напоминания');
  var last = sh.getLastRow();
  if (last < 2) return [];
  var data = sh.getRange(2, 1, last - 1, 4).getValues();
  var out = [];
  var i;
  for (i = 0; i < data.length; i++) {
    if (String(data[i][3]).toLowerCase() !== 'ждать') continue;
    out.push({ when: data[i][1], text: data[i][2] });
  }
  out.sort(function (a, b) {
    return new Date(a.when) - new Date(b.when);
  });
  return out.slice(0, limit || 5);
}

function checkReminders() {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('TELEGRAM_TOKEN');
  if (!token) return;
  var chatDefault = props.getProperty('ALLOWED_CHAT_ID');
  var sh = getSS_().getSheetByName('Напоминания');
  var last = sh.getLastRow();
  if (last < 2) return;
  var data = sh.getRange(2, 1, last - 1, 6).getValues();
  var now = new Date();
  var i;
  for (i = 0; i < data.length; i++) {
    if (String(data[i][3]).toLowerCase() !== 'ждать') continue;
    var when = data[i][1];
    var whenDate = when instanceof Date ? when : new Date(when);
    if (isNaN(whenDate.getTime())) continue;
    if (whenDate > now) continue;
    var chatId = String(data[i][5] || chatDefault || '');
    if (!chatId) continue;
    var text = '⏰ ' + String(data[i][2] || 'Напоминание');
    try {
      reply_(token, chatId, text);
      sh.getRange(i + 2, 4).setValue('отправлено');
    } catch (err) {
      Logger.log(err);
    }
  }
}

function morningPush() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('MORNING_PUSH') !== '1') return;
  var token = props.getProperty('TELEGRAM_TOKEN');
  var chatId = props.getProperty('ALLOWED_CHAT_ID');
  if (!token || !chatId) return;

  var hourWant = Number(props.getProperty('MORNING_HOUR') || '8');
  var hourNow = Number(Utilities.formatDate(new Date(), TZ, 'H'));
  var dayKey = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  if (hourNow !== hourWant) return;
  if (props.getProperty('MORNING_SENT') === dayKey) return;

  reply_(token, chatId, formatMorning_());
  props.setProperty('MORNING_SENT', dayKey);
}

function askDeepSeek_(userText) {
  var key = PropertiesService.getScriptProperties().getProperty('DEEPSEEK_API_KEY');
  if (!key) return 'Нет DEEPSEEK_API_KEY';
  var tasks = formatTaskList_();
  var system =
    'Ты личный ассистент Сергея Михалёва (Улан-Удэ, UTC+8). Помогаешь планировать и формулировать задачи. ' +
    'Не выдумывай факты о бизнесе. Коротко, для Telegram. Деньги и сметы не считай. ' +
    'Если просят добавить задачу — скажи написать боту: + текст\n\nТЕКУЩИЕ ЗАДАЧИ:\n' +
    tasks;
  var payload = {
    model: 'deepseek-chat',
    messages: [
      { role: 'system', content: system },
      { role: 'user', content: userText }
    ],
    temperature: 0.3
  };
  var res = UrlFetchApp.fetch('https://api.deepseek.com/chat/completions', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + key },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  var data = JSON.parse(res.getContentText());
  if (!data.choices || !data.choices[0]) return 'ИИ не ответил. Используй команды: /help';
  return String(data.choices[0].message.content || '').slice(0, 3500);
}

function reply_(token, chatId, text) {
  UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      chat_id: chatId,
      text: String(text).slice(0, 4000)
    })
  });
}

function textOut_(s) {
  return ContentService.createTextOutput(s);
}

/** Ручная проверка: sendTestReminder() */
function sendTestReminder() {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('TELEGRAM_TOKEN');
  var chatId = props.getProperty('ALLOWED_CHAT_ID');
  if (!token || !chatId) throw new Error('Нужны TELEGRAM_TOKEN и /start боту');
  reply_(token, chatId, 'Тест ассистента: связь ок ✓');
}
