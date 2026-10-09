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
const MESTO_MAX_PHOTOS     = 5;

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
    const photos = mesto_photos_(body.photos);
    const plan = body.plan ? mesto_photos_([body.plan])[0] : null;   // схема с меткой / обводкой
    const area = body.mode === 'area';                                 // «Обвести»: без координат
    if (area && !plan) throw new Error('Нет схемы с обведённой областью.');
    if (!floor || !(axes || plan)) throw new Error('Не выбрана точка на плане.');
    if (!plan && MESTO_SURFACES.indexOf(surface) < 0) throw new Error('Не выбрана поверхность.');
    if (surface && MESTO_SURFACES.indexOf(surface) < 0) throw new Error('Неизвестная поверхность.');

    // 4. Кто отправил (ищем в листе TOPIC_NAMES по @username)
    const person = mesto_findPerson_(user.username);
    const name = person.name || [user.first_name, user.last_name].filter(Boolean).join(' ') || ('id ' + user.id);

    // 5. Текст сообщения
    let text = '📍 ' + floor +
      (elevation ? ', отм. ' + elevation : '') +
      (axes ? ', оси ' + axes : '') +
      (surface ? ', ' + surface : '') +
      (area ? ' — область обведена на схеме' : plan ? ' — место отмечено на схеме' : '') + '.';
    if (comment) {
      const ru = mesto_translate_(comment, body.lang);   // турецкий (и любой не кириллический) → + русский
      text += ru
        ? '\nКомментарий (' + (/[çğışöüİ]/i.test(comment) || body.lang === 'tr' ? 'TR' : 'ориг.') + '): ' + comment +
          '\nКомментарий (RU): ' + ru
        : '\nКомментарий: ' + comment;
    }
    text += ' (от ' + name + ')';

    // 6. Пишем в тему: без фото — сообщение, с фото — фото/альбом с подписью
    const messageId = mesto_send_(target, text, plan ? [plan].concat(photos) : photos);

    // 7. Сразу пишем строку в Inbox (свои сообщения бот через getUpdates не видит)
    const link = mesto_messageLink_(target.chatId, target.threadId, messageId);
    const sender = name + (user.username ? ' (@' + user.username + ')' : '');
    const companyRole = [person.company, person.role].filter(String).join(' / ') || '-';
    const sheet = SpreadsheetApp.openById(obj.sheetId).getSheetByName('Inbox');
    if (!sheet) throw new Error('В таблице объекта «' + obj.name + '» нет листа Inbox.');
    const inboxText = text + (photos.length ? '\n📷 Фото: ' + photos.length : '') + (plan ? '\n🗺 Схема с меткой' : '');
    sheet.appendRow([new Date(), inboxText, link, sender, companyRole, 'Место', 'Нет']);

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

/** Вызов Telegram Bot API с файлами (multipart). Все поля — строки или Blob. */
function mesto_tgFiles_(method, fields) {
  const resp = UrlFetchApp.fetch('https://api.telegram.org/bot' + TELEGRAM_TOKEN + '/' + method, {
    method: 'post',
    payload: fields,
    muteHttpExceptions: true
  });
  const res = JSON.parse(resp.getContentText());
  if (!res.ok) throw new Error('Telegram (' + method + '): ' + res.description);
  return res.result;
}

/** Фото из мини-приложения: массив base64 JPEG → массив Blob (не больше MESTO_MAX_PHOTOS). */
function mesto_photos_(list) {
  if (!Array.isArray(list)) return [];
  return list.slice(0, MESTO_MAX_PHOTOS).map(function (b64, i) {
    if (typeof b64 !== 'string' || !/^[A-Za-z0-9+/=]+$/.test(b64) || b64.length > 8000000) {
      throw new Error('Фото ' + (i + 1) + ' повреждено или слишком большое.');
    }
    return Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', 'photo' + (i + 1) + '.jpg');
  });
}

/** Отправляет точку в тему. Возвращает message_id (для ссылки в Inbox). */
function mesto_send_(target, text, photos) {
  if (!photos.length) {
    return mesto_tg_('sendMessage', mesto_withThread_({ chat_id: target.chatId, text: text }, target.threadId)).message_id;
  }
  const caption = text.slice(0, 1024);   // лимит подписи в Telegram
  const fields = mesto_withThread_({ chat_id: target.chatId }, target.threadId);
  if (fields.message_thread_id) fields.message_thread_id = String(fields.message_thread_id);
  if (photos.length === 1) {
    fields.photo = photos[0];
    fields.caption = caption;
    return mesto_tgFiles_('sendPhoto', fields).message_id;
  }
  fields.media = JSON.stringify(photos.map(function (p, i) {
    const m = { type: 'photo', media: 'attach://p' + i };
    if (i === 0) m.caption = caption;
    return m;
  }));
  photos.forEach(function (p, i) { fields['p' + i] = p; });
  return mesto_tgFiles_('sendMediaGroup', fields)[0].message_id;
}

/**
 * Перевод комментария на русский, если он написан не кириллицей (обычно — по-турецки).
 * Возвращает русский текст или null (если перевод не нужен или не удался).
 */
function mesto_translate_(comment, lang) {
  if (!comment || /[а-яё]/i.test(comment) || !/[a-zçğışöü]/i.test(comment)) return null;
  try {
    const ru = LanguageApp.translate(comment, lang === 'tr' ? 'tr' : '', 'ru');
    return ru && ru.trim().toLowerCase() !== comment.trim().toLowerCase() ? ru.trim() : null;
  } catch (err) {
    console.error('mesto_translate_: ' + err.message);
    return null;
  }
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
 * Проверка без отправки сообщений: для каждого объекта из OBJECTS — открывается ли таблица,
 * есть ли лист Inbox, видит ли бот группу. Результат — в журнале.
 */
function mesto_diag() {
  try { console.log('Бот: @' + mesto_tg_('getMe', {}).username); }
  catch (e) { console.log('❌ Telegram: ' + e.message); }
  Object.keys(OBJECTS).forEach(function (id) {
    const o = OBJECTS[id];
    let line = o.name + ' (' + id + '): ';
    try {
      const ss = SpreadsheetApp.openById(o.sheetId);
      const sh = ss.getSheetByName('Inbox');
      line += 'таблица «' + ss.getName() + '» — ' + (sh ? 'лист Inbox есть, строк: ' + sh.getLastRow() : '❌ НЕТ листа «Inbox»');
    } catch (e) {
      line += '❌ таблица не открывается: ' + e.message;
    }
    try { line += '; группа «' + mesto_tg_('getChat', { chat_id: id }).title + '» видна'; }
    catch (e) { line += '; ❌ группа: ' + e.message; }
    console.log(line);
  });
  if (!OBJECTS['-1004448770601']) console.log('❌ В OBJECTS нет Биолаборатории (-1004448770601)');
}

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
  console.log(mesto_testPost_([]));
}

/**
 * Тест фото без телефона: отправляет в «тест 1» точку с двумя картинками (альбом).
 * Картинки берутся с сайта мини-приложения.
 */
function mesto_testSendPhoto() {
  const base = 'https://gerundg.github.io/monart-place/plans/БТЛ/монолит/';
  const photos = ['bio-18410.png', 'bio-3200.png'].map(function (f) {
    return Utilities.base64Encode(UrlFetchApp.fetch(encodeURI(base + f)).getBlob().getBytes());
  });
  console.log(mesto_testPost_(photos));
}

/**
 * Тест генплана без телефона: точка без осей + схема с меткой (кусок генплана с сайта) в «тест 1».
 */
function mesto_testSendPlan() {
  const plan = Utilities.base64Encode(
    UrlFetchApp.fetch(encodeURI('https://gerundg.github.io/monart-place/plans/мед колледж/med-genplan.png')).getBlob().getBytes());
  console.log(mesto_testPost_([], { floor: 'Ген.план', elevation: '', axes: '', surface: '', plan: plan,
                                    comment: 'Тест генплана из редактора' }));
}

/** Проверка перевода (ничего не отправляет). Первый запуск может попросить разрешение. */
function mesto_testTranslate() {
  console.log(mesto_translate_('Duvarda çatlak var, sıva dökülüyor', 'tr'));
}

/** Имитация запроса из мини-приложения с настоящей подписью (для тестов). */
function mesto_testPost_(photos, override) {
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

  const req = {
    initData: initData,
    floor: '1 этаж', elevation: '+0.000', axes: '3-4/Б-В',
    surface: 'потолок', comment: 'Тестовая отправка из редактора',
    photos: photos
  };
  Object.keys(override || {}).forEach(function (k) { req[k] = override[k]; });
  const out = doPost({ postData: { contents: JSON.stringify(req) } });
  return out.getContent();
}
