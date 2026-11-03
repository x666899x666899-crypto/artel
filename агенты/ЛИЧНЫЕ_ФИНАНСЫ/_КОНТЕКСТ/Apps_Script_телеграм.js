/**
 * Личные финансы Сергея — мост Telegram → таблица.
 * Таблица: 1hhnqgvsQ9maSWCbDrZztNNER797ljfQLY4hIQB81acg
 *
 * Спам шпаргалкой: не запускать registerWebhook (старая версия шлёт шпаргалку).
 * Опрос: installPolling() — раз в минуту (это минимум Google).
 * Счета в таблицу: setupAccounts() — один раз.
 * Напоминания 12:00 и 21:00: installReminders() — один раз.
 *
 * ИИ-чат: DEEPSEEK_API_KEY (из РФ) или YANDEX_API_KEY + YANDEX_FOLDER_ID.
 * Gemini из России ключ обычно не взять.
 * Не хранить токены в этом файле.
 */

var SPREADSHEET_ID = '1hhnqgvsQ9maSWCbDrZztNNER797ljfQLY4hIQB81acg';
var TZ = 'Asia/Irkutsk';
var DRIVE_FOLDER = 'Личные_финансы_входящее';

function getSS_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

var HELP_TEXT =
  'Запись: продукты 890 · кафе 1200 · гсм 2500\n' +
  'пополнил кредитку 5000 — не расход\n\n' +
  'Спросить цифры: остатки · сводка · сколько продукты\n' +
  'Поговорить: как с помощником, обычным текстом\n' +
  '/сброс — забыть разговор · /help — это меню';

var GEMINI_SYSTEM =
  'Ты личный финансовый помощник Сергея Михалёва (Улан-Удэ, UTC+8). Только личные деньги, не завод АРТЕЛЬ.\n' +
  'Цифры бери ТОЛЬКО из блока ДАННЫЕ. Нет цифры — скажи «нет в таблице», не выдумывай и не оценивай «на глаз».\n' +
  'Перевод между своими картами и пополнение кредитки — не расход жизни.\n' +
  'Тинькофф кредитка: остаток в таблице = сколько ещё можно потратить (доступный лимит), не долг.\n' +
  'Живые рубли = Альфа + Тинькофф дебет + наличка ₽. Кредитку в «есть деньги» не складывай.\n' +
  'Отвечай по-русски, коротко, для телефона: суть, цифры, что решить. Без канцелярита.\n' +
  'Не проси пароли, не предлагай заходить в банк. Журнал сам не меняй — только советуй, что написать боту (например: продукты 890).';

function setupLichnoe() {
  var ss = getSS_();
  ss.rename('Личные финансы');

  var specs = [
    { name: 'Входящее', headers: ['Дата_время', 'Тип', 'Текст', 'Сумма_черновик', 'Статья_черновик', 'Направление_черновик', 'Ссылка_файл', 'Telegram_id', 'Chat_id', 'Статус', 'Комментарий_агента'] },
    { name: 'Журнал', headers: ['Дата', 'Счёт', 'Направление', 'Статья', 'Сумма', 'Комментарий', 'Источник', 'Telegram_id', 'Статус', 'Ссылка_файл'] },
    { name: 'Счета', headers: ['Счёт', 'Банк', 'Чей', 'Остаток', 'Дата_остатка', 'Примечание'] },
    { name: 'Бюджет', headers: ['Статья', 'План', 'Факт', 'Комментарий'] },
    { name: 'Категории', headers: ['Код', 'Название', 'Тип'] },
    { name: 'Справка', headers: ['Как пользоваться'] }
  ];

  var i;
  for (i = 0; i < specs.length; i++) {
    ensureSheet_(ss, specs[i].name, specs[i].headers);
  }

  setupCategories();

  var help = ss.getSheetByName('Справка');
  help.getRange('A2').setValue(HELP_TEXT.replace(/\n/g, ' '));
}

function categoryRows_() {
  return [
    ['дивиденды', 'Вывод из АРТЕЛЬ', 'приход'],
    ['зарплата', 'Личная зарплата', 'приход'],
    ['прочий_приход', 'Прочий приход', 'приход'],
    ['продукты', 'Продукты', 'расход'],
    ['кафе', 'Кафе и рестораны', 'расход'],
    ['гсм', 'ГСМ', 'расход'],
    ['авто', 'Авто', 'расход'],
    ['ипотека', 'Ипотека', 'расход'],
    ['сотовая', 'Сотовая связь', 'расход'],
    ['одежда', 'Вещи и одежда', 'расход'],
    ['дом', 'Для дома', 'расход'],
    ['детям', 'Детям', 'расход'],
    ['настя_алина', 'Насте, Алине, Новосибирск', 'расход'],
    ['родители', 'Маме, папе', 'расход'],
    ['подарки', 'Подарки', 'расход'],
    ['развлечения', 'Развлечения', 'расход'],
    ['спорт', 'Спорт', 'расход'],
    ['прочее', 'Прочее', 'расход'],
    ['пополнение_кредитки', 'Пополнил кредитку', 'служебное'],
    ['перевод', 'Между своими счетами', 'служебное'],
    ['накопления', 'Накопления', 'служебное']
  ];
}

function expenseBudgetNames_() {
  return ['продукты', 'кафе', 'гсм', 'авто', 'ипотека', 'сотовая', 'одежда', 'дом', 'детям', 'настя_алина', 'родители', 'подарки', 'развлечения', 'спорт', 'прочее'];
}

function setupCategories() {
  var ss = getSS_();
  var cats = ss.getSheetByName('Категории');
  if (!cats) throw new Error('Нет листа Категории');
  var headers = ['Код', 'Название', 'Тип'];
  cats.getRange(1, 1, 1, 3).setValues([headers]);
  cats.setFrozenRows(1);
  cats.getRange(1, 1, 1, 3).setFontWeight('bold');
  var rows = categoryRows_();
  var last = cats.getLastRow();
  if (last > 1) cats.getRange(2, 1, last - 1, 3).clearContent();
  cats.getRange(2, 1, rows.length, 3).setValues(rows);

  var bud = ss.getSheetByName('Бюджет');
  if (!bud) throw new Error('Нет листа Бюджет');
  bud.getRange(1, 1, 1, 4).setValues([['Статья', 'План', 'Факт', 'Комментарий']]);
  bud.setFrozenRows(1);
  bud.getRange(1, 1, 1, 4).setFontWeight('bold');
  var names = expenseBudgetNames_();
  var bLast = bud.getLastRow();
  if (bLast > 1) bud.getRange(2, 1, bLast - 1, 4).clearContent();
  var brow = [];
  var i;
  for (i = 0; i < names.length; i++) {
    var r = i + 2;
    brow.push([names[i], '', '=SUMIFS(Журнал!E:E;Журнал!D:D;A' + r + ';Журнал!C:C;"расход";Журнал!I:I;"ок")', '']);
  }
  bud.getRange(2, 1, brow.length, 4).setValues(brow);

  var help = ss.getSheetByName('Справка');
  if (help) help.getRange('A2').setValue(HELP_TEXT.replace(/\n/g, ' '));
}

function setupAccounts() {
  var sh = getSS_().getSheetByName('Счета');
  if (!sh) throw new Error('Нет листа Счета');
  var headers = ['Счёт', 'Банк', 'Чей', 'Остаток', 'Дата_остатка', 'Примечание'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  sh.getRange(2, 1, 9, 6).setValues([
    ['Тинькофф дебетовая', 'Тинькофф', 'Сергей', 13130, today, '₽'],
    ['Тинькофф кредитная', 'Тинькофф', 'Сергей', 114249, today, '₽; лимит 120000; остаток=доступно; занято 5751'],
    ['Альфа счёт', 'Альфа-Банк', 'Сергей', 178816, today, '₽'],
    ['Наличка рубли', 'наличные', 'Сергей', '', '', '₽'],
    ['Наличка доллары', 'наличные', 'Сергей', '', '', '$'],
    ['Наличка юани', 'наличные', 'Сергей', '', '', '¥'],
    ['Alipay юани', 'Alipay', 'Сергей', 2553.28, today, '¥; Китай'],
    ['WeChat Pay юани', 'WeChat Pay', 'Сергей', 2859.3, today, '¥; Китай'],
    ['Плати по миру', 'Плати по миру', 'Сергей', 18.47, today, '$']
  ]);
}

function registerWebhook() {
  var token = PropertiesService.getScriptProperties().getProperty('TELEGRAM_TOKEN');
  if (!token) throw new Error('Нет TELEGRAM_TOKEN в свойствах скрипта');
  var url = ScriptApp.getService().getUrl();
  if (!url) throw new Error('Сначала: Развернуть → Новое развёртывание → Веб-приложение');
  var res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/setWebhook', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ url: url })
  });
  Logger.log(url);
  Logger.log(res.getContentText());
}

function checkWebhook() {
  var token = PropertiesService.getScriptProperties().getProperty('TELEGRAM_TOKEN');
  if (!token) throw new Error('Нет TELEGRAM_TOKEN');
  var res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/getWebhookInfo');
  Logger.log(ScriptApp.getService().getUrl() || 'веб-приложение ещё не развёрнуто');
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
  var triggers = ScriptApp.getProjectTriggers();
  var i;
  for (i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'pollTelegram') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('pollTelegram').timeBased().everyMinutes(1).create();
  pollTelegram();
}

var REMINDER_TEXT =
  '<b>Напоминание</b>\n\n' +
  'Не забудь отправить все движения по счетам: траты, приходы, переводы, новые остатки.\n\n' +
  'Примеры:\n' +
  '<code>продукты 890</code>\n' +
  '<code>зарплата 50000</code>\n' +
  '<code>альфа 178816</code>\n' +
  '<code>пополнил кредитку 15000</code>';

function sendDailyReminder() {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('TELEGRAM_TOKEN');
  var chatId = props.getProperty('ALLOWED_CHAT_ID');
  if (!token || !chatId) {
    Logger.log('нет TELEGRAM_TOKEN или ALLOWED_CHAT_ID');
    return;
  }
  reply_(token, chatId, REMINDER_TEXT, true);
}

function installReminders() {
  var triggers = ScriptApp.getProjectTriggers();
  var i;
  for (i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendDailyReminder') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('sendDailyReminder').timeBased().atHour(12).nearMinute(0).everyDays(1).create();
  ScriptApp.newTrigger('sendDailyReminder').timeBased().atHour(21).nearMinute(0).everyDays(1).create();
  sendDailyReminder();
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
    try {
      var u = JSON.parse(e.postData.contents);
      var m = u.message || u.edited_message;
      if (m) reply_(token, String(m.chat.id), 'Не записал: ' + err.message);
    } catch (e2) {}
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
  Logger.log('chat=' + chatId + ' text=' + text);

  if (/^\/start(@[A-Za-z0-9_]+)?$/i.test(text)) {
    reply_(token, chatId, 'Трата: продукты 890\nВопрос: обычным текстом, как помощнику\nЦифры: остатки / сводка\n/help — меню');
    return;
  }
  if (/^\/help(@[A-Za-z0-9_]+)?$/i.test(text)) {
    reply_(token, chatId, HELP_TEXT);
    return;
  }
  if (/^\/(сброс|reset)(@[A-Za-z0-9_]+)?$/i.test(text)) {
    props.deleteProperty('TG_HISTORY');
    reply_(token, chatId, 'Разговор сбросил. Спрашивай заново.');
    return;
  }

  var hasFile = !!(msg.photo || msg.voice || msg.audio || msg.document || msg.video);
  var writeText = isWriteText_(text);

  if (writeText || (hasFile && !isQuery_(text))) {
    var parsed = parseMessage_(token, msg);
    var g0 = guessArticle_(text);
    parsed.article = g0.article;
    parsed.direction = g0.direction;
    appendInbox_(parsed);

    var acc = matchAccount_(text);
    if (acc && parsed.amount !== '') {
      updateAccount_(acc, parsed.amount);
      reply_(token, chatId, h1_('Остаток обновлён') + '\n' + preBlock_([preRow_(acc, money_(parsed.amount))]), true);
      return;
    }

    if (parsed.amount !== '' && text && !isNoise_(text)) {
      var g = g0;
      appendJournalDraft_(parsed, g.article, g.direction);
      ensureCategory_(g.article, g.direction);
      var recTitle = h1_('Записал');
      if (g.created) recTitle += ' · ' + muted_('новая статья');
      reply_(token, chatId, recTitle + '\n' + preBlock_([preRow_(artTitle_(g.article), money_(parsed.amount) + ' ₽')]), true);
      return;
    }

    var sum = parsed.amount !== '' ? parsed.amount + '' : 'сумму разберём в Cursor';
    var label = text || parsed.type;
    reply_(token, chatId, 'Записал во входящее: ' + label + ' — ' + sum);
    return;
  }

  if (text && !isNoise_(text)) {
    var canAi = hasAiKey_(props);
    if (canAi && isTalk_(text) && !isShortCommand_(text)) {
      var aiTalk = askGemini_(props, text);
      if (aiTalk && aiTalk.ok) {
        reply_(token, chatId, aiTalk.text);
        return;
      }
      if (aiTalk && aiTalk.text) {
        reply_(token, chatId, aiTalk.text);
        return;
      }
    }
    if (isShortCommand_(text) || (isQuery_(text) && !isTalk_(text))) {
      reply_(token, chatId, buildReport_(normQuery_(text)), true);
      return;
    }
    var ai = askGemini_(props, text);
    if (ai && ai.ok) {
      reply_(token, chatId, ai.text);
      return;
    }
    if (ai && ai.noKey) {
      reply_(token, chatId, 'Чтобы я разговаривал, нужен ключ ИИ.\nGemini из России не открывается.\nВариант: DeepSeek — свойство DEEPSEEK_API_KEY.\nПока цифры: остатки, сводка.');
      return;
    }
    if (ai && ai.text) {
      reply_(token, chatId, ai.text);
      return;
    }
    reply_(token, chatId, 'Не понял. Трата: продукты 890. Или вопрос своими словами.');
    return;
  }

  reply_(token, chatId, 'Напиши трату (продукты 890) или вопрос своими словами.');
}

function doGet() {
  return textOut_('lichnoe ok');
}

function parseMessage_(token, msg) {
  var type = 'текст';
  var text = msg.text || msg.caption || '';
  var fileId = '';
  var fileUrl = '';

  if (msg.photo && msg.photo.length) {
    type = 'фото';
    fileId = msg.photo[msg.photo.length - 1].file_id;
  } else if (msg.voice) {
    type = 'голос';
    fileId = msg.voice.file_id;
  } else if (msg.audio) {
    type = 'голос';
    fileId = msg.audio.file_id;
  } else if (msg.document) {
    type = 'файл';
    fileId = msg.document.file_id;
  } else if (msg.video) {
    type = 'фото';
    fileId = msg.video.file_id;
  }

  if (fileId) fileUrl = saveTelegramFile_(token, fileId);

  var amount = extractAmount_(text);
  return {
    type: type,
    text: text,
    amount: amount,
    fileUrl: fileUrl,
    telegramId: msg.message_id,
    chatId: msg.chat.id
  };
}

function extractAmount_(text) {
  if (!text) return '';
  var t = text.replace(/\s/g, '');
  var m = t.match(/(\d+(?:[.,]\d{1,2})?)/);
  if (!m) return '';
  return Number(m[1].replace(',', '.'));
}

function isNoise_(text) {
  var t = (text || '').trim().toLowerCase();
  return /^(start|stop|остановить|тест)$/i.test(t);
}

function guessArticle_(text) {
  var raw = (text || '').trim();
  var t = raw.toLowerCase().replace(/ё/g, 'е');
  if (/пополнил?\s*кредит|пополнение\s*кредит/.test(t)) {
    return { article: 'пополнение_кредитки', direction: 'перевод', created: false };
  }
  if (/(поступлен|приход|зарплат|дивиденд)/.test(t) && !/расход|трата/.test(t)) {
    var inc = 'прочий_приход';
    if (/зарплат/.test(t)) inc = 'зарплата';
    if (/дивиденд/.test(t)) inc = 'дивиденды';
    return { article: inc, direction: 'приход', created: false };
  }
  t = t.replace(/^(расход|трата|траты|потратил|купил|оплата|платеж)\s+/, '');
  var first = (t.split(/\s+/)[0] || '').replace(/[^а-яa-z_]/g, '');
  var firstMap = {
    продукты: 'продукты', еда: 'продукты',
    кафе: 'кафе', ресторан: 'кафе', рестораны: 'кафе',
    гсм: 'гсм', бензин: 'гсм', заправка: 'гсм',
    авто: 'авто', машина: 'авто', такси: 'авто',
    ипотека: 'ипотека',
    сотовая: 'сотовая', связь: 'сотовая',
    одежда: 'одежда', вещи: 'одежда',
    дом: 'дом',
    детям: 'детям', дети: 'детям',
    насте: 'настя_алина', алине: 'настя_алина', новосиб: 'настя_алина', нвосиб: 'настя_алина',
    маме: 'родители', папе: 'родители', мама: 'родители', папа: 'родители',
    подарки: 'подарки', подарок: 'подарки',
    развлечения: 'развлечения',
    спорт: 'спорт', зал: 'спорт', фитнес: 'спорт'
  };
  if (firstMap[first]) {
    return { article: firstMap[first], direction: 'расход', created: false };
  }
  var rules = [
    { re: /кафе|ресторан/, article: 'кафе' },
    { re: /еда|продукт|магазин|пятероч|лента|магнит/, article: 'продукты' },
    { re: /гсм|бензин|заправк/, article: 'гсм' },
    { re: /ипотек/, article: 'ипотека' },
    { re: /сотов|мтс|билайн|мегафон|теле2/, article: 'сотовая' },
    { re: /одежд|вещи/, article: 'одежда' },
    { re: /для\s*дома|хозтовар/, article: 'дом' },
    { re: /детям|дети /, article: 'детям' },
    { re: /насте|алине|новосиб|нвосиб/, article: 'настя_алина' },
    { re: /маме |папе |родител/, article: 'родители' },
    { re: /подарок|подарки/, article: 'подарки' },
    { re: /развлеч/, article: 'развлечения' },
    { re: /спорт|фитнес|спортзал/, article: 'спорт' },
    { re: /осаго|каско/, article: 'авто' }
  ];
  var i;
  for (i = 0; i < rules.length; i++) {
    if (rules[i].re.test(t)) {
      return { article: rules[i].article, direction: 'расход', created: false };
    }
  }
  if (first.length >= 3 && first !== 'расход' && first !== 'приход') {
    return { article: first, direction: 'расход', created: true };
  }
  return { article: 'прочее', direction: 'расход', created: false };
}

function ensureCategory_(code, direction) {
  if (!code || code === 'прочее') return;
  var sh = getSS_().getSheetByName('Категории');
  if (!sh) return;
  var last = sh.getLastRow();
  var type = direction === 'перевод' ? 'служебное' : (direction === 'приход' ? 'приход' : 'расход');
  if (last >= 2) {
    var vals = sh.getRange(2, 1, last - 1, 1).getValues();
    var i;
    for (i = 0; i < vals.length; i++) {
      if (String(vals[i][0]) === code) return;
    }
  }
  sh.appendRow([code, code, type]);
}

function matchAccount_(text) {
  var t = (text || '').toLowerCase().replace(/\s+/g, ' ').trim();
  if (/пополнил|пополнение/.test(t)) return '';
  if (/^плати\s*по\s*миру/.test(t)) return 'Плати по миру';
  if (/^(алипэй|alipay)/.test(t)) return 'Alipay юани';
  if (/^(вичат|wechat)/.test(t)) return 'WeChat Pay юани';
  if (/^наличк.*долл|^доллар.*налич/.test(t)) return 'Наличка доллары';
  if (/^наличк.*юан|^юан.*налич/.test(t)) return 'Наличка юани';
  if (/^наличк/.test(t)) return 'Наличка рубли';
  if (/^альфа/.test(t)) return 'Альфа счёт';
  if (/^тиньк\S*\s*(кредит|кредитк)/.test(t)) return 'Тинькофф кредитная';
  if (/^тиньк\S*\s*(дебет|дебетов)|^тиньков/.test(t)) return 'Тинькофф дебетовая';
  return '';
}

function updateAccount_(name, amount) {
  var sh = getSS_().getSheetByName('Счета');
  var last = sh.getLastRow();
  if (last < 2) return;
  var vals = sh.getRange(2, 1, last - 1, 6).getValues();
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var i;
  for (i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) === name) {
      sh.getRange(i + 2, 4).setValue(amount);
      sh.getRange(i + 2, 5).setValue(today);
      return;
    }
  }
}

function appendJournalDraft_(p, article, direction) {
  var sh = getSS_().getSheetByName('Журнал');
  if (!sh) throw new Error('Нет листа Журнал');
  sh.appendRow([
    Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'),
    '',
    direction || 'расход',
    article || 'прочее',
    p.amount,
    p.text,
    'telegram',
    p.telegramId,
    'черновик',
    p.fileUrl || ''
  ]);
}

function saveTelegramFile_(token, fileId) {
  var meta = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/getFile?file_id=' + fileId);
  var path = JSON.parse(meta.getContentText()).result.file_path;
  var blob = UrlFetchApp.fetch('https://api.telegram.org/file/bot' + token + '/' + path).getBlob();
  var folder = getFolder_();
  var file = folder.createFile(blob);
  file.setName(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd_HH-mm-ss') + '_' + file.getName());
  return file.getUrl();
}

function getFolder_() {
  var it = DriveApp.getFoldersByName(DRIVE_FOLDER);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(DRIVE_FOLDER);
}

function appendInbox_(p) {
  var sh = getSS_().getSheetByName('Входящее');
  if (!sh) throw new Error('Нет листа Входящее');
  sh.appendRow([
    Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),
    p.type,
    p.text,
    p.amount,
    p.article || '',
    p.direction || '',
    p.fileUrl,
    p.telegramId,
    p.chatId,
    'новое',
    ''
  ]);
}

function reply_(token, chatId, text, asHtml) {
  var s = String(text || '');
  if (s.length > 4000) s = s.substring(0, 3990) + '\n…';
  var payload = { chat_id: chatId, text: s, disable_web_page_preview: true };
  if (asHtml) payload.parse_mode = 'HTML';
  var res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  if (asHtml && res.getResponseCode() !== 200) {
    payload.parse_mode = undefined;
    payload.text = stripHtml_(s);
    UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  }
}

function escHtml_(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function stripHtml_(s) {
  return String(s || '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(b|strong|i|em|code|pre|u)>/gi, '')
    .replace(/<(b|strong|i|em|code|pre|u)(\s[^>]*)?>/gi, '')
    .replace(/<[^>]+>/g, '')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>');
}

function moneyHtml_(n) {
  return '<code>' + escHtml_(money_(n)) + '</code>';
}

function preBlock_(rows) {
  if (!rows.length) return '';
  return '<pre>' + rows.join('\n') + '</pre>';
}

function preRow_(name, value) {
  var n = String(name || '');
  var v = String(value || '');
  if (n.length > 16) n = n.substring(0, 15) + '…';
  var gap = 28 - n.length - v.length;
  if (gap < 1) gap = 1;
  var sp = '';
  var i;
  for (i = 0; i < gap; i++) sp += ' ';
  return escHtml_(n + sp + v);
}

function artTitle_(code) {
  var rows = categoryRows_();
  var i;
  for (i = 0; i < rows.length; i++) {
    if (rows[i][0] === code) return rows[i][1];
  }
  return code;
}

function h1_(t) {
  return '<b>' + escHtml_(t) + '</b>';
}

function muted_(t) {
  return '<i>' + escHtml_(t) + '</i>';
}

function normQuery_(text) {
  return (text || '').replace(/^\/([^\s@]+)(@[A-Za-z0-9_]+)?/, '$1').trim();
}

function isWriteText_(text) {
  var t = (text || '').trim();
  if (!t) return false;
  if (isQuery_(t)) return false;
  return extractAmount_(t) !== '';
}

function pingGemini() {
  var props = PropertiesService.getScriptProperties();
  var ans = askGemini_(props, 'Коротко: ты на связи? Назови живые рубли из данных.');
  Logger.log(ans && ans.ok ? ans.text : JSON.stringify(ans));
}

function sheetSnapshot_() {
  var period = parsePeriod_('месяц');
  var s = [formatAccounts_(''), formatJournal_(period, ''), formatBudget_(period)].join('\n\n');
  s = stripHtml_(s);
  if (s.length > 3500) s = s.substring(0, 3490) + '\n…';
  return s;
}

function historyGet_(props) {
  try {
    var h = JSON.parse(props.getProperty('TG_HISTORY') || '[]');
    return h instanceof Array ? h : [];
  } catch (e) {
    return [];
  }
}

function historyPush_(props, question, answer) {
  var h = historyGet_(props);
  h.push({ role: 'user', text: question });
  h.push({ role: 'assistant', text: answer });
  while (h.length > 8) h.shift();
  var raw = JSON.stringify(h);
  if (raw.length > 7000) {
    h = h.slice(-4);
    raw = JSON.stringify(h);
  }
  props.setProperty('TG_HISTORY', raw);
}

function historyToChat_(props) {
  var h = historyGet_(props);
  var out = [];
  var i;
  for (i = 0; i < h.length; i++) {
    var role = h[i].role === 'model' ? 'assistant' : h[i].role;
    var text = h[i].text || '';
    if (!text && h[i].parts && h[i].parts[0]) text = h[i].parts[0].text || '';
    if (role === 'user' || role === 'assistant') {
      if (text) out.push({ role: role, content: text });
    }
  }
  return out;
}

function promptUser_(question) {
  return 'ДАННЫЕ ИЗ ТАБЛИЦЫ (единственный источник цифр):\n' + sheetSnapshot_() + '\n\nВОПРОС СЕРГЕЯ:\n' + question;
}

function geminiModels_(props) {
  var preferred = props.getProperty('GEMINI_MODEL') || '';
  var list = [preferred, 'gemini-2.5-flash', 'gemini-flash-latest', 'gemini-2.0-flash', 'gemini-3.5-flash', 'gemini-1.5-flash'];
  var out = [];
  var i;
  for (i = 0; i < list.length; i++) {
    if (list[i] && out.indexOf(list[i]) === -1) out.push(list[i]);
  }
  return out;
}

function geminiText_(body) {
  var j = JSON.parse(body);
  if (j.error) {
    throw new Error((j.error.message || 'Gemini error') + '');
  }
  var c = j.candidates && j.candidates[0];
  if (!c) return '';
  if (c.finishReason === 'SAFETY') return 'Не могу так ответить. Спроси иначе.';
  var parts = c.content && c.content.parts;
  if (!parts) return '';
  var t = '';
  var i;
  for (i = 0; i < parts.length; i++) {
    if (parts[i].text) t += parts[i].text;
  }
  return t.trim();
}

function askDeepSeek_(props, key, userText, question) {
  var messages = [{ role: 'system', content: GEMINI_SYSTEM }].concat(historyToChat_(props));
  messages.push({ role: 'user', content: userText });
  var res = UrlFetchApp.fetch('https://api.deepseek.com/chat/completions', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + key },
    payload: JSON.stringify({
      model: props.getProperty('DEEPSEEK_MODEL') || 'deepseek-chat',
      messages: messages,
      temperature: 0.3,
      max_tokens: 800
    }),
    muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  var body = res.getContentText();
  if (code < 200 || code >= 300) {
    Logger.log('DeepSeek HTTP ' + code + ' ' + body.substring(0, 300));
    return { ok: false, text: 'DeepSeek не ответил (HTTP ' + code + '). Пока: остатки / сводка.' };
  }
  var j = JSON.parse(body);
  var text = j.choices && j.choices[0] && j.choices[0].message && j.choices[0].message.content;
  if (!text) return { ok: false, text: 'DeepSeek пустой ответ. Пока: остатки / сводка.' };
  historyPush_(props, question, text);
  return { ok: true, text: String(text).trim() };
}

function askYandex_(props, key, folderId, userText, question) {
  var hist = historyToChat_(props);
  var messages = [{ role: 'system', text: GEMINI_SYSTEM }];
  var i;
  for (i = 0; i < hist.length; i++) {
    messages.push({ role: hist[i].role === 'assistant' ? 'assistant' : 'user', text: hist[i].content });
  }
  messages.push({ role: 'user', text: userText });
  var model = props.getProperty('YANDEX_MODEL') || 'yandexgpt-lite';
  var res = UrlFetchApp.fetch('https://llm.api.cloud.yandex.net/foundationModels/v1/completion', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Api-Key ' + key },
    payload: JSON.stringify({
      modelUri: 'gpt://' + folderId + '/' + model + '/latest',
      completionOptions: { stream: false, temperature: 0.3, maxTokens: 800 },
      messages: messages
    }),
    muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  var body = res.getContentText();
  if (code < 200 || code >= 300) {
    Logger.log('Yandex HTTP ' + code + ' ' + body.substring(0, 300));
    return { ok: false, text: 'YandexGPT не ответил (HTTP ' + code + '). Пока: остатки / сводка.' };
  }
  var j = JSON.parse(body);
  var text = j.result && j.result.alternatives && j.result.alternatives[0] && j.result.alternatives[0].message && j.result.alternatives[0].message.text;
  if (!text) return { ok: false, text: 'YandexGPT пустой ответ. Пока: остатки / сводка.' };
  historyPush_(props, question, text);
  return { ok: true, text: String(text).trim() };
}

function askGeminiApi_(props, key, question, userText) {
  var hist = historyGet_(props);
  var contents = [];
  var i;
  for (i = 0; i < hist.length; i++) {
    var role = hist[i].role === 'assistant' || hist[i].role === 'model' ? 'model' : 'user';
    var t = hist[i].text || (hist[i].parts && hist[i].parts[0] && hist[i].parts[0].text) || '';
    if (t) contents.push({ role: role, parts: [{ text: t }] });
  }
  contents.push({ role: 'user', parts: [{ text: userText }] });
  var payload = {
    systemInstruction: { parts: [{ text: GEMINI_SYSTEM }] },
    contents: contents,
    generationConfig: { temperature: 0.3, maxOutputTokens: 800 }
  };
  var models = geminiModels_(props);
  var lastErr = '';
  for (i = 0; i < models.length; i++) {
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + models[i] + ':generateContent?key=' + encodeURIComponent(key);
    var res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    var body = res.getContentText();
    if (code >= 200 && code < 300) {
      var text = geminiText_(body);
      if (text) {
        historyPush_(props, question, text);
        props.setProperty('GEMINI_MODEL', models[i]);
        return { ok: true, text: text };
      }
      lastErr = 'пустой ответ ' + models[i];
      continue;
    }
    lastErr = 'HTTP ' + code + ' ' + models[i];
    if (code === 404 || code === 400) continue;
    Logger.log(lastErr + ' ' + body.substring(0, 300));
    return { ok: false, text: 'ИИ сейчас не ответил (' + lastErr + '). Цифры: напиши «остатки» или «сводка».' };
  }
  Logger.log(lastErr);
  return { ok: false, text: 'Gemini не подключился. Из России ключ обычно не взять. Пока: остатки / сводка.' };
}

function askGemini_(props, question) {
  var userText = promptUser_(question);
  var ds = props.getProperty('DEEPSEEK_API_KEY');
  if (ds) return askDeepSeek_(props, ds, userText, question);
  var yk = props.getProperty('YANDEX_API_KEY');
  var yf = props.getProperty('YANDEX_FOLDER_ID');
  if (yk && yf) return askYandex_(props, yk, yf, userText, question);
  var gk = props.getProperty('GEMINI_API_KEY');
  if (gk) return askGeminiApi_(props, gk, question, userText);
  return { ok: false, noKey: true };
}

function hasAiKey_(props) {
  if (props.getProperty('DEEPSEEK_API_KEY')) return true;
  if (props.getProperty('YANDEX_API_KEY') && props.getProperty('YANDEX_FOLDER_ID')) return true;
  if (props.getProperty('GEMINI_API_KEY')) return true;
  return false;
}

function isTalk_(text) {
  var t = (text || '').trim();
  if (!t) return false;
  if (/\?/.test(t)) return true;
  var low = t.toLowerCase().replace(/ё/g, 'е');
  if (/(^|\s)ли(\s|\?|$)|скажи|какие|какой|какое|привет|почему|нормально|хватит ли|можно ли|что с |как дела|поступлен/.test(low)) return true;
  if (low.replace(/[^\sа-яa-z]+/g, ' ').trim().split(/\s+/).length >= 5) return true;
  return false;
}

function isShortCommand_(text) {
  var t = normQuery_(text).toLowerCase().replace(/ё/g, 'е').replace(/[?!.,:;]+/g, '').trim();
  if (!t) return false;
  if (/^(остатк\S*|сводка|отчет|отчёт|бюджет|сегодня|вчера|месяц|неделя|недел[яи]|живые|итог)$/.test(t)) return true;
  if (/^можно\s+потрат/.test(t) && extractAmount_(text) !== '') return true;
  if (/^сколько(\s+\S+){1,3}$/.test(t)) return true;
  var words = t.split(/\s+/);
  if (words.length <= 3 && findAccountInText_(t) && extractAmount_(t) === '') return true;
  if (words.length === 1) {
    var g = guessArticle_(t);
    if (g && !g.created && g.article !== 'прочее') return true;
  }
  return false;
}

function isQuery_(text) {
  var t = (text || '').trim();
  if (!t) return false;
  if (/^\/(start|help|сброс|reset)(@[A-Za-z0-9_]+)?$/i.test(t)) return false;
  if (/^\//.test(t)) return true;
  var low = t.toLowerCase().replace(/ё/g, 'е');
  if (/(сколько|сводка|отчет|отчёт|остатк|баланс|тратил|потратил|расход|приход|бюджет|хватит|можно\s+потрат|за месяц|за неделю|сегодня|вчера|итог|живые)/.test(low)) return true;
  if (extractAmount_(t) !== '') return false;
  if (matchAccount_(low) || findAccountInText_(low)) return true;
  var g = guessArticle_(t);
  if (g && !g.created && g.article !== 'прочее') return true;
  return false;
}

function articleFromQuery_(text) {
  var t = normQuery_(text).toLowerCase().replace(/ё/g, 'е');
  t = t.replace(/сколько|сводка|отчет|отчёт|трат[аы]?|расход|приход|бюджет|остатк\S*|баланс|живые|хватит|можно|потратить|потратил|тратил|за|на|по|сегодня|вчера|недел\S*|месяц|прошл\S*|итог/g, ' ');
  t = t.replace(/\s+/g, ' ').trim();
  if (!t) return '';
  var g = guessArticle_(t);
  if (!g || g.created || g.article === 'прочее') return '';
  return g.article;
}

function queryKind_(text) {
  var low = normQuery_(text).toLowerCase().replace(/ё/g, 'е');
  var kind = { accounts: false, journal: false, budget: false, canSpend: false, incomeOnly: false, article: articleFromQuery_(text), accountName: '' };
  if (/можно\s*потрат/.test(low) && extractAmount_(text) !== '') kind.canSpend = true;
  if (/бюджет|план vs|план-факт/.test(low) && !isTalk_(text)) kind.budget = true;
  if (/^бюджет$/.test(low.replace(/[?!]/g, '').trim())) kind.budget = true;
  if (/сводка|отчет|отчёт/.test(low)) {
    kind.accounts = true;
    kind.journal = true;
  }
  if (/остатк|баланс|на счетах|сколько денег|живые/.test(low)) kind.accounts = true;
  if (/сводка|отчет|отчёт|трат|расход|приход|итог|месяц|недел|сегодня|вчера|сколько/.test(low) || kind.article) kind.journal = true;
  if (/^остатк/.test(low) && !kind.article) {
    kind.accounts = true;
    kind.journal = false;
  }
  if (!/сводка|отчет|отчёт/.test(low)) {
    kind.accountName = findAccountInText_(low);
    if (kind.accountName) kind.accounts = true;
  }
  if (/поступлен|приход/.test(low) && !/расход/.test(low)) {
    kind.incomeOnly = true;
    kind.journal = true;
  }
  if (kind.accountName && !kind.article && !/сводка|отчет|отчёт|трат|расход|месяц|недел/.test(low)) {
    kind.journal = false;
    kind.accounts = true;
  }
  if (kind.canSpend) kind.accounts = false;
  if (!kind.accounts && !kind.journal && !kind.budget && !kind.canSpend) {
    kind.accounts = true;
    kind.journal = true;
  }
  return kind;
}

function findAccountInText_(text) {
  var t = (text || '').toLowerCase();
  if (/плати\s*по\s*миру/.test(t)) return 'Плати по миру';
  if (/алипэй|alipay/.test(t)) return 'Alipay юани';
  if (/вичат|wechat/.test(t)) return 'WeChat Pay юани';
  if (/наличк.*долл|доллар/.test(t) && /налич/.test(t)) return 'Наличка доллары';
  if (/наличк.*юан|юан/.test(t) && /налич/.test(t)) return 'Наличка юани';
  if (/наличк/.test(t)) return 'Наличка рубли';
  if (/альфа/.test(t)) return 'Альфа счёт';
  if (/тиньк.*кредит|кредитк/.test(t) && !/пополн/.test(t)) return 'Тинькофф кредитная';
  if (/тиньк.*дебет|тиньков|тинькофф/.test(t)) return 'Тинькофф дебетовая';
  return '';
}

function asDate_(v) {
  if (!v && v !== 0) return null;
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }
  var s = String(v);
  var m = s.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return null;
}

function inPeriod_(v, from, to) {
  var dt = asDate_(v);
  if (!dt) return false;
  var a = new Date(from.getFullYear(), from.getMonth(), from.getDate());
  var b = new Date(to.getFullYear(), to.getMonth(), to.getDate());
  return dt.getTime() >= a.getTime() && dt.getTime() <= b.getTime();
}

function parsePeriod_(text) {
  var now = new Date();
  var low = (text || '').toLowerCase().replace(/ё/g, 'е');
  var y = Number(Utilities.formatDate(now, TZ, 'yyyy'));
  var mo = Number(Utilities.formatDate(now, TZ, 'M')) - 1;
  var d = Number(Utilities.formatDate(now, TZ, 'd'));
  var today = new Date(y, mo, d);
  if (/вчера/.test(low)) {
    var yest = new Date(y, mo, d - 1);
    return { from: yest, to: yest, label: Utilities.formatDate(yest, TZ, 'yyyy-MM-dd') };
  }
  if (/сегодня/.test(low)) {
    return { from: today, to: today, label: 'сегодня' };
  }
  if (/недел/.test(low)) {
    return { from: new Date(y, mo, d - 6), to: today, label: '7 дней' };
  }
  if (/прошл\S*\s*месяц/.test(low)) {
    var pf = new Date(y, mo - 1, 1);
    var pt = new Date(y, mo, 0);
    return { from: pf, to: pt, label: Utilities.formatDate(pf, TZ, 'yyyy-MM') };
  }
  var months = [['январ', 0], ['феврал', 1], ['март', 2], ['апрел', 3], ['ма[йя]', 4], ['июн', 5], ['июл', 6], ['август', 7], ['сентябр', 8], ['октябр', 9], ['ноябр', 10], ['декабр', 11]];
  var i;
  for (i = 0; i < months.length; i++) {
    if (new RegExp(months[i][0]).test(low)) {
      var yy = y;
      var ym = low.match(/20\d{2}/);
      if (ym) yy = Number(ym[0]);
      var f = new Date(yy, months[i][1], 1);
      var t = new Date(yy, months[i][1] + 1, 0);
      return { from: f, to: t, label: Utilities.formatDate(f, TZ, 'yyyy-MM') };
    }
  }
  var iso = low.match(/(20\d{2})-(\d{2})/);
  if (iso) {
    var f2 = new Date(Number(iso[1]), Number(iso[2]) - 1, 1);
    var t2 = new Date(Number(iso[1]), Number(iso[2]), 0);
    return { from: f2, to: t2, label: iso[0] };
  }
  var rng = low.match(/с\s+(\d{1,2})(?:[./](\d{1,2}))?(?:[./](\d{2,4}))?\s+по\s+(\d{1,2})(?:[./](\d{1,2}))?(?:[./](\d{2,4}))?/);
  if (rng) {
    var mo1 = rng[2] ? Number(rng[2]) - 1 : mo;
    var y1 = rng[3] ? (String(rng[3]).length === 2 ? 2000 + Number(rng[3]) : Number(rng[3])) : y;
    var mo2 = rng[5] ? Number(rng[5]) - 1 : mo1;
    var y2 = rng[6] ? (String(rng[6]).length === 2 ? 2000 + Number(rng[6]) : Number(rng[6])) : y1;
    var fa = new Date(y1, mo1, Number(rng[1]));
    var ta = new Date(y2, mo2, Number(rng[4]));
    return { from: fa, to: ta, label: Utilities.formatDate(fa, TZ, 'dd.MM') + '–' + Utilities.formatDate(ta, TZ, 'dd.MM') };
  }
  var f3 = new Date(y, mo, 1);
  return { from: f3, to: today, label: Utilities.formatDate(f3, TZ, 'yyyy-MM') };
}

function money_(n) {
  if (n === '' || n === null || typeof n === 'undefined' || isNaN(Number(n))) return '—';
  var x = Number(n);
  var neg = x < 0;
  x = Math.abs(x);
  var s = (Math.round(x * 100) / 100).toFixed(2).replace('.', ',');
  if (/,00$/.test(s)) s = s.replace(/,00$/, '');
  var parts = s.split(',');
  parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
  s = parts.join(',');
  return (neg ? '−' : '') + s;
}

function num_(v) {
  if (v === '' || v === null || typeof v === 'undefined') return null;
  if (typeof v === 'number' && !isNaN(v)) return v;
  var s = String(v).replace(/\s/g, '').replace(',', '.');
  if (!s) return null;
  var n = Number(s);
  return isNaN(n) ? null : n;
}

function isTransfer_(dir, article) {
  var d = String(dir || '').toLowerCase();
  var a = String(article || '').toLowerCase();
  return d === 'перевод' || a === 'перевод' || a === 'пополнение_кредитки';
}

function readAccounts_() {
  var sh = getSS_().getSheetByName('Счета');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
}

function formatAccounts_(onlyName) {
  var rows = readAccounts_();
  var rub = [];
  var fx = [];
  var creditLines = [];
  var live = 0;
  var found = false;
  var i;
  for (i = 0; i < rows.length; i++) {
    var name = String(rows[i][0] || '');
    var rest = num_(rows[i][3]);
    var note = String(rows[i][5] || '');
    var dt = asDate_(rows[i][4]);
    var date = dt ? Utilities.formatDate(dt, TZ, 'dd.MM') : '';
    if (!name) continue;
    if (/кредитн/i.test(name)) {
      if (onlyName && name !== onlyName) continue;
      found = true;
      var lim = note.match(/лимит\s*(\d+)/i);
      var limit = lim ? Number(lim[1]) : 0;
      var used = limit && rest !== null ? limit - rest : null;
      creditLines.push(preRow_('доступно', (rest === null ? '—' : money_(rest) + ' ₽')));
      if (used !== null) creditLines.push(preRow_('занято', money_(used) + ' ₽'));
      if (limit) creditLines.push(preRow_('лимит', money_(limit) + ' ₽'));
      continue;
    }
    var cur = /¥/.test(note) ? '¥' : (/\$/.test(note) ? '$' : '₽');
    if (cur === '₽' && rest !== null) live += rest;
    if (onlyName && name !== onlyName) continue;
    found = true;
    var val = rest === null ? 'нет цифры' : money_(rest) + ' ' + cur;
    if (cur === '₽') rub.push(preRow_(name, val));
    else fx.push(preRow_(name, val));
  }
  var out = [];
  if (onlyName) out.push(h1_('Счёт') + ' · ' + escHtml_(onlyName));
  else out.push(h1_('Остатки'));
  if (onlyName && !found) {
    out.push(muted_('Такого счёта нет в таблице.'));
    return out.join('\n');
  }
  if (rub.length) {
    if (!onlyName) out.push(muted_('рубли'));
    out.push(preBlock_(rub));
  }
  if (!onlyName || (onlyName && /кредитн/i.test(onlyName))) {
    if (creditLines.length) {
      out.push(h1_('Кредитка Тинькофф'));
      out.push(muted_('остаток = сколько ещё можно'));
      out.push(preBlock_(creditLines));
    }
  }
  if (fx.length) {
    if (!onlyName) out.push(muted_('валюта / Китай'));
    out.push(preBlock_(fx));
  }
  if (!onlyName) {
    out.push(h1_('Живые ₽') + '  ' + moneyHtml_(live));
    out.push(muted_('Альфа + дебет + наличка ₽, без кредитки'));
  }
  return out.join('\n');
}

function liveRubles_() {
  var rows = readAccounts_();
  var live = 0;
  var i;
  for (i = 0; i < rows.length; i++) {
    var name = String(rows[i][0] || '');
    var rest = num_(rows[i][3]);
    var note = String(rows[i][5] || '');
    if (rest === null) continue;
    if (/кредитн/i.test(name)) continue;
    if (/¥/.test(note) || /\$/.test(note)) continue;
    live += rest;
  }
  return live;
}

function creditAvailable_() {
  var rows = readAccounts_();
  var i;
  for (i = 0; i < rows.length; i++) {
    if (/кредитн/i.test(String(rows[i][0] || ''))) return num_(rows[i][3]);
  }
  return null;
}

function formatCanSpend_(amount) {
  if (amount === '' || amount === null) {
    return h1_('Можно потратить?') + '\n' + muted_('Напиши сумму: можно потратить 80000');
  }
  var live = liveRubles_();
  var cred = creditAvailable_();
  var n = Number(amount);
  var out = [];
  out.push(h1_('Можно ' + money_(n) + ' ₽?'));
  out.push(preBlock_([
    preRow_('живые', money_(live) + ' ₽'),
    preRow_('нужно', money_(n) + ' ₽')
  ]));
  if (n <= live) {
    out.push('<b>Да.</b> С дебета / Альфы.');
    out.push('Останется  ' + moneyHtml_(live - n) + ' ₽');
  } else if (cred !== null && n <= live + cred) {
    out.push('<b>Живых не хватает</b> на ' + moneyHtml_(n - live) + ' ₽');
    out.push('На кредитке доступно  ' + moneyHtml_(cred) + ' ₽');
  } else {
    out.push('<b>Нет.</b> Не хватает ' + moneyHtml_(n - live) + ' ₽ живыми.');
  }
  out.push(muted_('Ипотека и обязательные платежи впереди — смотри отдельно.'));
  return out.join('\n');
}

function readJournal_() {
  var sh = getSS_().getSheetByName('Журнал');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
}

function formatJournal_(period, article, incomeOnly) {
  var rows = readJournal_();
  var exp = 0;
  var inc = 0;
  var tr = 0;
  var drafts = 0;
  var by = {};
  var byInc = {};
  var nExp = 0;
  var nInc = 0;
  var i;
  for (i = 0; i < rows.length; i++) {
    if (!inPeriod_(rows[i][0], period.from, period.to)) continue;
    var dir = String(rows[i][2] || '');
    var art = String(rows[i][3] || 'прочее');
    var sum = num_(rows[i][4]);
    var st = String(rows[i][8] || '');
    if (sum === null) continue;
    if (article && art !== article) continue;
    if (st === 'черновик') drafts++;
    if (isTransfer_(dir, art)) {
      tr += sum;
      continue;
    }
    if (dir === 'приход' || art === 'дивиденды' || art === 'зарплата' || art === 'прочий_приход') {
      inc += sum;
      nInc++;
      byInc[art] = (byInc[art] || 0) + sum;
      continue;
    }
    exp += sum;
    nExp++;
    by[art] = (by[art] || 0) + sum;
  }
  if (incomeOnly) {
    var outI = [h1_('Поступления · ' + period.label)];
    if (!nInc) {
      outI.push(muted_('Прихода за период в журнале нет.'));
      return outI.join('\n');
    }
    outI.push(preBlock_([preRow_('+ приход', money_(inc) + ' ₽')]));
    var ik = [];
    for (var k2 in byInc) ik.push(k2);
    ik.sort(function (a, b) { return byInc[b] - byInc[a]; });
    var catI = [];
    for (i = 0; i < ik.length; i++) catI.push(preRow_(artTitle_(ik[i]), money_(byInc[ik[i]]) + ' ₽'));
    outI.push(preBlock_(catI));
    if (drafts) outI.push(muted_('черновиков: ' + drafts));
    return outI.join('\n');
  }
  var title = 'Расходы · ' + period.label;
  if (article) title = artTitle_(article) + ' · ' + period.label;
  var out = [h1_(title)];
  if (nExp === 0 && inc === 0 && tr === 0) {
    out.push(muted_('В журнале записей нет.'));
    out.push('Пиши трату: <code>продукты 890</code>');
    return out.join('\n');
  }
  var tot = [preRow_('− расход', money_(exp) + ' ₽')];
  if (inc) tot.push(preRow_('+ приход', money_(inc) + ' ₽'));
  if (tr) tot.push(preRow_('переводы', money_(tr) + ' ₽'));
  tot.push(preRow_('= сальдо', money_(inc - exp) + ' ₽'));
  out.push(preBlock_(tot));
  var keys = [];
  for (var k in by) keys.push(k);
  keys.sort(function (a, b) { return by[b] - by[a]; });
  var max = Math.min(keys.length, 12);
  if (max && !article) {
    out.push(muted_('по статьям'));
    var cat = [];
    for (i = 0; i < max; i++) {
      var pct = exp ? Math.round(by[keys[i]] / exp * 100) : 0;
      cat.push(preRow_(artTitle_(keys[i]), money_(by[keys[i]]) + '  ' + pct + '%'));
    }
    out.push(preBlock_(cat));
  }
  if (article && nExp) {
    out.push(muted_(nExp + ' операций · ' + money_(exp) + ' ₽'));
  }
  if (drafts) out.push(muted_('черновиков: ' + drafts + ' (ещё не подтверждены)'));
  return out.join('\n');
}

function bar_(fact, plan) {
  var n = 8;
  if (!plan || plan <= 0) return '';
  var p = fact / plan;
  if (p < 0) p = 0;
  if (p > 1) p = 1;
  var fill = Math.round(p * n);
  var s = '';
  var i;
  for (i = 0; i < n; i++) s += i < fill ? '█' : '░';
  return s;
}

function formatBudget_(period) {
  var sh = getSS_().getSheetByName('Бюджет');
  var out = [h1_('Бюджет') + ' · ' + escHtml_(period.label)];
  var facts = {};
  var rows = readJournal_();
  var i;
  for (i = 0; i < rows.length; i++) {
    if (!inPeriod_(rows[i][0], period.from, period.to)) continue;
    if (isTransfer_(rows[i][2], rows[i][3])) continue;
    if (String(rows[i][2]) === 'приход') continue;
    var art = String(rows[i][3] || 'прочее');
    var sum = num_(rows[i][4]);
    if (sum === null) continue;
    facts[art] = (facts[art] || 0) + sum;
  }
  var hasPlan = false;
  var cat = [];
  if (sh && sh.getLastRow() >= 2) {
    var b = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    for (i = 0; i < b.length; i++) {
      var name = String(b[i][0] || '');
      var plan = num_(b[i][1]);
      if (!name) continue;
      var fact = facts[name] || 0;
      if (plan === null && !fact) continue;
      if (plan !== null) hasPlan = true;
      var left = artTitle_(name);
      var right;
      if (plan === null) {
        right = money_(fact) + ' ₽';
      } else {
        var over = fact - plan;
        right = money_(fact) + '/' + money_(plan);
        if (over > 0) right += ' +' + money_(over);
      }
      cat.push(preRow_(left, right));
      if (plan !== null) {
        var btxt = bar_(fact, plan);
        if (btxt) cat.push(escHtml_('  ' + btxt + '  ' + Math.round(fact / plan * 100) + '%'));
      }
    }
  }
  if (cat.length) out.push(preBlock_(cat));
  if (!hasPlan) out.push(muted_('План пустой — цифры на лист Бюджет.'));
  return out.join('\n');
}

function buildReport_(text) {
  var kind = queryKind_(text);
  var period = parsePeriod_(text);
  var parts = [];
  if (kind.canSpend) parts.push(formatCanSpend_(extractAmount_(text)));
  if (kind.journal) parts.push(formatJournal_(period, kind.article, kind.incomeOnly));
  if (kind.budget) parts.push(formatBudget_(period));
  if (kind.accounts) parts.push(formatAccounts_(kind.accountName));
  if (!parts.length) {
    return h1_('Не понял запрос') + '\n' + muted_('остатки · сводка · сколько продукты');
  }
  return parts.join('\n\n');
}

function ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
}

function textOut_(s) {
  return ContentService.createTextOutput(s);
}
