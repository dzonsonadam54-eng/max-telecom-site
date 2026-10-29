/**
 * MAX Telecom — пересылка отчёта из 1С на сайт.
 *
 * 1С по расписанию присылает письмо с отчётом «Товары для сайта» (.xlsx) на этот Gmail.
 * Скрипт раз в 15 минут находит новые такие письма и отправляет файл на сайт.
 * Доступа к 1С у скрипта нет и не нужно.
 *
 * Установка (один раз, в Gmail-ящике для приёма):
 *  1. script.google.com → Новый проект → вставить этот код.
 *  2. Настройки проекта (шестерёнка) → Свойства скрипта → добавить:
 *       SITE_URL   = https://max-telecom-site.vercel.app
 *       SITE_TOKEN = значение STOCK_PUSH_TOKEN из Vercel
 *       FROM       = адрес, с которого 1С отправляет рассылку (необязательно, но надёжнее)
 *  3. Выполнить функцию setup один раз и разрешить доступ к Gmail.
 *     Она создаст метку «site-processed» и триггер «каждые 15 минут».
 */

const SUBJECT = 'Товары для сайта'; // часть темы письма рассылки из 1С

function setup() {
  GmailApp.createLabel('site-processed');
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('relay').timeBased().everyMinutes(15).create();
  relay();
}

function relay() {
  const p = PropertiesService.getScriptProperties();
  const url = (p.getProperty('SITE_URL') || '').replace(/\/$/, '') + '/api/stock-upload';
  const token = p.getProperty('SITE_TOKEN');
  const from = p.getProperty('FROM');
  if (!token) throw new Error('Не задан SITE_TOKEN в свойствах скрипта');

  const done = GmailApp.getUserLabelByName('site-processed');
  const query = `subject:"${SUBJECT}" has:attachment filename:xlsx -label:site-processed newer_than:3d` + (from ? ` from:${from}` : '');
  const threads = GmailApp.search(query, 0, 10);
  if (!threads.length) return;

  // Only the newest report matters: older ones in the same batch are just marked processed.
  const messages = threads.flatMap(t => t.getMessages()).sort((a, b) => b.getDate() - a.getDate());
  const newest = messages[0];
  const att = newest.getAttachments().find(a => /\.xlsx$/i.test(a.getName()));
  if (att) {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({ fileBase64: Utilities.base64Encode(att.getBytes()), filename: att.getName(), updated: newest.getDate().toISOString() }),
      muteHttpExceptions: true,
    });
    const code = res.getResponseCode();
    if (code !== 200) {
      // Leave the mail unprocessed so it is retried, and notify the owner of this mailbox.
      MailApp.sendEmail(Session.getActiveUser().getEmail(), 'Сайт MAX: каталог не обновился',
        'Сайт ответил ' + code + ': ' + res.getContentText().slice(0, 500) +
        '\n\nКаталог на сайте остался прежним. Проверьте вариант отчёта «Товары для сайта» в 1С.');
      return;
    }
  }
  threads.forEach(t => t.addLabel(done).markRead());
}
