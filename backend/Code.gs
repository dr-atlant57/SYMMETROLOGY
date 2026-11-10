/**
 * SYMMETROLOGY.ORG — owned lead intake (Google Apps Script, V8 runtime).
 * Deploy as Web app: Execute as Me; Who has access: Anyone.
 * Run setup() once as the owner of nextdollmill@gmail.com.
 * Google Sheet is the durable record; email notification is secondary.
 */
const CONFIG = Object.freeze({
  RECIPIENT: 'nextdollmill@gmail.com',
  SHEET_PROPERTY: 'SYMMETROLOGY_REQUEST_SHEET_ID',
  SHEET_NAME: 'Requests',
  PUBLIC_SITE: 'https://symmetrology.org/',
  DEDUP_SECONDS: 90
});

function setup() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(CONFIG.SHEET_PROPERTY);
  if (!id) {
    const ss = SpreadsheetApp.create('SYMMETROLOGY — Requests');
    const sheet = ss.getSheets()[0];
    sheet.setName(CONFIG.SHEET_NAME);
    sheet.appendRow(['Received at', 'Receipt ID', 'Email', 'Request type', 'Page', 'Language', 'Email notification', 'Details']);
    sheet.setFrozenRows(1);
    id = ss.getId();
    props.setProperty(CONFIG.SHEET_PROPERTY, id);
  }
  // Prompt user to grant MailApp scope before public deployment.
  const remaining = MailApp.getRemainingDailyQuota();
  const url = 'https://docs.google.com/spreadsheets/d/' + id + '/edit';
  Logger.log('Request log: ' + url);
  Logger.log('Remaining email recipient quota today: ' + remaining);
  return url;
}

function doGet() {
  const configured = Boolean(PropertiesService.getScriptProperties().getProperty(CONFIG.SHEET_PROPERTY));
  return page_(configured, 'en', configured ? 'ready' : 'unavailable', '');
}

function doPost(e) {
  const p = (e && e.parameter) || {};
  const lang = /^(en|ru|de|ka)$/.test(String(p.lang || '')) ? String(p.lang) : 'en';
  const email = String(p.email || '').trim().toLowerCase();
  if (String(p.website || p._honey || '').trim()) return page_(false, lang, 'invalid', '');
  if (email.length > 254 || !/^[^\s@<>]{1,64}@[^\s@<>.]+(?:\.[^\s@<>.]+)+$/.test(email)) {
    return page_(false, lang, 'invalid', '');
  }
  const id = PropertiesService.getScriptProperties().getProperty(CONFIG.SHEET_PROPERTY);
  if (!id) return page_(false, lang, 'unavailable', '');

  const requestType = String(p.request_type || 'early_access').slice(0, 60);
  const source = String(p.source || CONFIG.PUBLIC_SITE).slice(0, 200);
  let lock;
  try {
    lock = LockService.getScriptLock();
    lock.waitLock(15000);
    const cache = CacheService.getScriptCache();
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, email + ':' + requestType);
    const cacheKey = 'req_' + Utilities.base64EncodeWebSafe(digest).replace(/=+$/, '');
    const existing = cache.get(cacheKey);
    if (existing) return page_(true, lang, 'duplicate', existing);

    const sheet = SpreadsheetApp.openById(id).getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) throw new Error('Request sheet not found');
    const receiptId = Utilities.getUuid();
    const now = new Date();
    // Prefix strings that could be interpreted as spreadsheet formulas.
    const safeEmail = sheetText_(email);
    sheet.appendRow([now, receiptId, safeEmail, sheetText_(requestType), sheetText_(source), lang, 'PENDING', '']);
    const row = sheet.getLastRow();
    SpreadsheetApp.flush(); // Confirm persistence before acknowledging success.
    cache.put(cacheKey, receiptId, CONFIG.DEDUP_SECONDS);

    // Email failure never loses a request: every successful capture remains in Sheets.
    let status = 'SENT';
    let details = '';
    try {
      if (MailApp.getRemainingDailyQuota() < 1) throw new Error('MailApp daily quota exhausted');
      MailApp.sendEmail({
        to: CONFIG.RECIPIENT,
        subject: '[SYMMETROLOGY] New request ' + receiptId.slice(0, 8),
        body: 'New Symmetrology.org request\n\nEmail: ' + email +
              '\nType: ' + requestType + '\nSource: ' + source +
              '\nLanguage: ' + lang + '\nReceipt ID: ' + receiptId +
              '\n\nAll requests are saved in the SYMMETROLOGY — Requests Google Sheet.',
        replyTo: email,
        name: 'Symmetrology — requests'
      });
    } catch (mailError) {
      status = 'EMAIL_ERROR';
      details = String(mailError && mailError.message || mailError).slice(0, 250);
      console.error('Notification failure for ' + receiptId + ': ' + details);
    }
    sheet.getRange(row, 7, 1, 2).setValues([[status, details]]);
    SpreadsheetApp.flush();
    return page_(true, lang, 'received', receiptId);
  } catch (error) {
    console.error('Lead intake failed:', error);
    return page_(false, lang, 'unavailable', '');
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (_) {} }
  }
}

/** Run manually in the Apps Script editor AFTER setup(). Sends one TEST lead. */
function testDelivery() {
  const result = doPost({parameter:{
    email: CONFIG.RECIPIENT,
    request_type: 'owner_delivery_test_' + new Date().getTime(),
    source: CONFIG.PUBLIC_SITE,
    lang: 'ru',
    website: ''
  }});
  Logger.log('Test executed. Open the Requests sheet and check Gmail Inbox/Spam.');
  return result;
}

function sheetText_(value) {
  const s = String(value);
  return /^[=+\-@\t\r]/.test(s) ? "'" + s : s;
}

function page_(success, lang, state, receipt) {
  const tr = {
    en: {title: 'Symmetrology — Request', received:'Your request has been recorded.', duplicate:'Your request was already recorded.', ready:'Request service is configured.', invalid:'Please enter a valid email address and try again.', unavailable:'We could not record your request. Please contact us directly.', home:'Return to Symmetrology.org', direct:'Write to us directly', ref:'Reference'},
    ru: {title: 'Symmetrology — Запрос', received:'Ваш запрос сохранён.', duplicate:'Ваш запрос уже сохранён.', ready:'Сервис заявок настроен.', invalid:'Проверьте email и повторите отправку.', unavailable:'Не удалось сохранить запрос. Напишите нам напрямую.', home:'Вернуться на Symmetrology.org', direct:'Написать напрямую', ref:'Номер заявки'},
    de: {title: 'Symmetrology — Anfrage', received:'Ihre Anfrage wurde gespeichert.', duplicate:'Ihre Anfrage wurde bereits gespeichert.', ready:'Anfragedienst ist eingerichtet.', invalid:'Bitte prüfen Sie Ihre E-Mail-Adresse.', unavailable:'Ihre Anfrage konnte nicht gespeichert werden. Kontaktieren Sie uns direkt.', home:'Zurück zu Symmetrology.org', direct:'E-Mail senden', ref:'Referenz'},
    ka: {title: 'Symmetrology — მოთხოვნა', received:'თქვენი მოთხოვნა შენახულია.', duplicate:'თქვენი მოთხოვნა უკვე შენახულია.', ready:'მოთხოვნების სერვისი გამართულია.', invalid:'შეამოწმეთ ელფოსტა და სცადეთ ხელახლა.', unavailable:'მოთხოვნის შენახვა ვერ მოხერხდა. მოგვწერეთ პირდაპირ.', home:'დაბრუნება Symmetrology.org-ზე', direct:'მოგვწერეთ', ref:'განაცხადის ნომერი'}
  };
  const t = tr[lang] || tr.en;
  const label = t[state] || t.unavailable;
  const html = '<!doctype html><html lang="' + lang + '"><head><meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1"><title>' + t.title + '</title>' +
    '<style>body{margin:0;background:#0b101b;color:#ecf0f6;font:17px/1.55 system-ui,Arial,sans-serif;display:grid;min-height:100vh;place-items:center;padding:22px;box-sizing:border-box}main{max-width:570px;padding:34px;border:1px solid #334154;border-radius:16px}h1{font-size:28px;line-height:1.25}p{color:#c8d3e1}a{color:#8bc5f5}small{color:#92a4ba}</style></head><body><main><h1>' +
    (success ? '✓ ' : '⚠ ') + label + '</h1>' +
    (success && receipt ? '<p><small>' + t.ref + ': ' + receipt + '</small></p>' : '') +
    '<p><a href="' + CONFIG.PUBLIC_SITE + '">' + t.home + '</a></p>' +
    (!success ? '<p><a href="mailto:' + CONFIG.RECIPIENT + '?subject=Symmetrology%20Request">' + t.direct + '</a></p>' : '') +
    '</main></body></html>';
  return HtmlService.createHtmlOutput(html).setTitle(t.title).addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
