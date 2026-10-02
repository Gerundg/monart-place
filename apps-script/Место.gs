/**
 * Место.gs — приём точек из мини-приложения «📍 Указать место».
 * Использует глобальные TELEGRAM_TOKEN и OBJECTS из «Сбор информации.gs» (здесь НЕ переобъявляются).
 * Сбор сообщений и напоминалку не трогает.
 */

// ===== Настройки этого файла =====
const MESTO_BOT_USERNAME   = 'monart_smartsite_bot';
const MESTO_APP_SHORT_NAME = 'place';            // короткое имя Mini App из BotFather
const MESTO_PEOPLE_CHAT    = '-1004315776569';   // объект, в таблице которого лежит лист TOPIC_NAMES
const MESTO_PEOPLE_SHEET   = 'TOPIC_NAMES';
const MESTO_SURFACES       = ['пол', 'потолок', 'стена'];

// Куда повесить закреп с кнопкой (функция mesto_postPinButton)
// Уже закреплено (не запускать повторно, иначе будет вторая кнопка):
//   { chatId: '-1004315776569', threadId: 2 }   // мед колледж, тема «тест 1» — 2026-10-02
const MESTO_PIN_TARGETS = [
  { chatId: '-1004448770601', threadId: 260 },   // Биолаборатория
  { chatId: '-1004448770601', threadId: 265 },   // Биолаборатория, «Замечания СК»
  { chatId: '-1004448770601', threadId: 222 }    // Биолаборатория
];
// =================================


/** Проверка в браузере: открыть URL веб-приложения — увидите эту строку. */
function doGet() {
  return ContentService.createTextOutput('Место: веб-приложение работает ✅');
}

/** Приём данных из мини-приложения. */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    // 1. Проверяем, что запрос действительно из Telegram (подпись initData)
    const init = mesto_checkInitData_(body.initData);
    const user = JSON.parse(init.user || '{}');

    // 2. Объект и тема берём из подписанного start_param (подделать нельзя)
    const target = mesto_parseStartParam_(init.start_param);
    const obj = OBJECTS[target.chatId];
    if (!obj) throw new Error('Неизвестный объект: ' + target.chatId);

    // 3. Данные точки
    const floor     = mesto_clean_(body.floor, 40);
    const elevation = mesto_clean_(body.elevation, 20);
    const axes      = mesto_clean_(body.axes, 40);
    const surface   = mesto_clean_(body.surface, 20);
    const comment   = mesto_clean_(body.comment, 500);
    if (!floor || !axes) throw new Error('Не выбрана точка на плане.');
    if (MESTO_SURFACES.indexOf(surface) < 0) throw new Error('Не выбрана поверхность.');

    // 4. Кто отправил (ищем в листе TOPIC_NAMES по @username)
    const person = mesto_findPerson_(user.username);
    const name = person.name || [user.first_name, user.last_name].filter(String).join(' ') || ('id ' + user.id);

    // 5. Текст сообщения
    let text = '📍 ' + floor + ', отм. ' + elevation + ', оси ' + axes + ', ' + surface + '.';
    if (comment) text += '\nКомментарий: ' + comment;
    text += ' (от ' + name + ')';

    // 6. Пишем в тему
    const sent = mesto_tg_('sendMessage', mesto_withThread_({
      chat_id: target.chatId,
      text: text
    }, target.threadId));

    // 7. Сразу пишем строку в Inbox (свои сообщения бот через getUpdates не видит)
    const link = mesto_messageLink_(target.chatId, target.threadId, sent.message_id);
    const sender = name + (user.username ? ' (@' + user.username + ')' : '');
    const companyRole = [person.company, person.role].filter(String).join(' / ') || '-';
    const sheet = SpreadsheetApp.openById(obj.sheetId).getSheetByName('Inbox');
    if (!sheet) throw new Error('В таблице объекта «' + obj.name + '» нет листа Inbox.');
    sheet.appendRow([new Date(), text, link, sender, companyRole, 'Место', 'Нет']);

    return mesto_json_({ ok: true });
  } catch (err) {
    console.error('Место doPost: ' + err.message);
    return mesto_json_({ ok: false, error: err.message });
  }
}


// ================= Вспомогательные функции =================

/** Проверка подписи Telegram initData. Возвращает поля initData объектом. */
function mesto_checkInitData_(initData) {
  if (!initData) throw new Error('Откройте приложение из Telegram (нет подписи).');

  const data = {};
  String(initData).split('&').forEach(function (pair) {
    const i = pair.indexOf('=');
    if (i < 0) return;
    data[mesto_decode_(pair.slice(0, i))] = mesto_decode_(pair.slice(i + 1));
  });
  if (!data.hash) throw new Error('Нет подписи Telegram.');

  const secret = Utilities.computeHmacSha256Signature(
    Utilities.newBlob(TELEGRAM_TOKEN).getBytes(),
    Utilities.newBlob('WebAppData').getBytes()
  );
  const variants = [['hash'], ['hash', 'signature']];
  const valid = variants.some(function (exclude) {
    const checkString = Object.keys(data)
      .filter(function (k) { return exclude.indexOf(k) < 0; })
      .sort()
      .map(function (k) { return k + '=' + data[k]; })
      .join('\n');
    const sig = Utilities.computeHmacSha256Signature(Utilities.newBlob(checkString).getBytes(), secret);
    return mesto_hex_(sig) === data.hash;
  });
  if (!valid) throw new Error('Подпись Telegram не прошла проверку.');

  const age = Date.now() / 1000 - Number(data.auth_date || 0);
  if (age > 86400) throw new Error('Сессия устарела. Закройте и откройте приложение заново.');
  return data;
}

/** "-1004315776569_2" → { chatId: '-1004315776569', threadId: 2 } */
function mesto_parseStartParam_(sp) {
  if (!sp) throw new Error('Нет привязки к теме. Откройте по кнопке из закрепа.');
  const i = sp.lastIndexOf('_');
  const chatId = i > 0 ? sp.slice(0, i) : sp;
  const threadId = i > 0 ? Number(sp.slice(i + 1)) : 0;
  return { chatId: chatId, threadId: threadId || 0 };
}

/** Поиск человека в листе TOPIC_NAMES по @username (колонки A, E, F, G). */
function mesto_findPerson_(username) {
  const empty = { name: '', company: '', role: '' };
  if (!username) return empty;
  try {
    const sheetId = OBJECTS[MESTO_PEOPLE_CHAT].sheetId;
    const sh = SpreadsheetApp.openById(sheetId).getSheetByName(MESTO_PEOPLE_SHEET);
    if (!sh) return empty;
    const rows = sh.getDataRange().getValues();
    const want = String(username).replace(/^@/, '').toLowerCase();
    for (let r = 1; r < rows.length; r++) {
      const u = String(rows[r][0] || '').trim().replace(/^@/, '').toLowerCase();
      if (u && u === want) {
        return {
          name: String(rows[r][4] || '').trim(),
          company: String(rows[r][5] || '').trim(),
          role: String(rows[r][6] || '').trim()
        };
      }
    }
  } catch (err) {
    console.error('mesto_findPerson_: ' + err.message);
  }
  return empty;
}

/** Вызов Telegram Bot API. */
function mesto_tg_(method, payload) {
  const resp = UrlFetchApp.fetch('https://api.telegram.org/bot' + TELEGRAM_TOKEN + '/' + method, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const res = JSON.parse(resp.getContentText());
  if (!res.ok) throw new Error('Telegram (' + method + '): ' + res.description);
  return res.result;
}

/** Добавляет message_thread_id, если тема не «Общая». */
function mesto_withThread_(payload, threadId) {
  if (threadId && threadId !== 1) payload.message_thread_id = threadId;
  return payload;
}

/** Ссылка на сообщение: https://t.me/c/4315776569/2/123 */
function mesto_messageLink_(chatId, threadId, messageId) {
  const internal = String(chatId).replace(/^-100/, '');
  return threadId && threadId !== 1
    ? 'https://t.me/c/' + internal + '/' + threadId + '/' + messageId
    : 'https://t.me/c/' + internal + '/' + messageId;
}

function mesto_clean_(v, max) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, max); }
function mesto_decode_(s) { return decodeURIComponent(String(s).replace(/\+/g, ' ')); }
function mesto_hex_(bytes) { return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join(''); }
function mesto_json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}


// ================= Запускать вручную из редактора =================

/**
 * Публикует в каждой теме из MESTO_PIN_TARGETS сообщение с кнопкой «📍 Указать место» и закрепляет его.
 * Боту нужно право «Закреплять сообщения».
 */
function mesto_postPinButton() {
  MESTO_PIN_TARGETS.forEach(function (t) {
    const link = 'https://t.me/' + MESTO_BOT_USERNAME + '/' + MESTO_APP_SHORT_NAME +
                 '?startapp=' + t.chatId + '_' + t.threadId;
    const msg = mesto_tg_('sendMessage', mesto_withThread_({
      chat_id: t.chatId,
      text: 'Нажмите кнопку, чтобы отметить место дефекта или работ на плане.',
      reply_markup: { inline_keyboard: [[{ text: '📍 Указать место', url: link }]] }
    }, t.threadId));
    mesto_tg_('pinChatMessage', { chat_id: t.chatId, message_id: msg.message_id, disable_notification: true });
    console.log('Закреплено в ' + t.chatId + ' / тема ' + t.threadId + ': ' + link);
  });
}

/**
 * Тест без телефона: имитирует отправку из мини-приложения в тему «тест 1».
 * Должно прийти сообщение в тему и появиться строка в Inbox.
 */
function mesto_testSend() {
  const fields = {
    auth_date: String(Math.floor(Date.now() / 1000)),
    query_id: 'test',
    start_param: '-1004315776569_2',
    user: JSON.stringify({ id: 1, first_name: 'Тест', username: 'test_user' })
  };
  const checkString = Object.keys(fields).sort().map(function (k) { return k + '=' + fields[k]; }).join('\n');
  const secret = Utilities.computeHmacSha256Signature(
    Utilities.newBlob(TELEGRAM_TOKEN).getBytes(), Utilities.newBlob('WebAppData').getBytes());
  fields.hash = mesto_hex_(Utilities.computeHmacSha256Signature(Utilities.newBlob(checkString).getBytes(), secret));
  const initData = Object.keys(fields).map(function (k) {
    return encodeURIComponent(k) + '=' + encodeURIComponent(fields[k]);
  }).join('&');

  const out = doPost({ postData: { contents: JSON.stringify({
    initData: initData,
    floor: '1 этаж', elevation: '+0.000', axes: '3–4 / Б–В',
    surface: 'потолок', comment: 'Тестовая отправка из редактора'
  }) } });
  console.log(out.getContent());
}
