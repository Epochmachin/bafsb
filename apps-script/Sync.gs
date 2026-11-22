/**
 * BAFSB Student Tracker → Google Calendar + Drive backup + Google Sheet copy
 * 1) Exam dates, mark-submission deadlines and the class routine go to your Google Calendar.
 * 2) The whole app data is backed up to Drive folder "BAFSB Tracker Backup" (one file per day, last 60 days).
 * 3) A readable copy (students, marks, notes) is written to the Google Sheet "BAFSB Tracker Data".
 *
 * Security: the link ends with ?k=SECRET. Without the right secret nothing happens.
 * Change KEY below to your own secret before Deploy, and put the same link in the app (Settings → Google link).
 * After changing code: Deploy → Manage deployments → ✏️ → Version: New version → Deploy (keeps the same link).
 */
const KEY = '__KEY__';
const CAL_ID = 'primary';
const TZ = 'Asia/Dhaka';
const BK_FOLDER = 'BAFSB Tracker Backup';
const BK_KEEP = 60;
const SHEET_NAME = 'BAFSB Tracker Data';

function doGet(e) { return handle_(e, {action: 'ping'}); }

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ok: false, error: 'Bad data'}); }
  return handle_(e, body);
}

function handle_(e, body) {
  if (!e || !e.parameter || e.parameter.k !== KEY) return out_({ok: false, error: 'Secret code does not match'});
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const cal = CAL_ID === 'primary' ? CalendarApp.getDefaultCalendar() : CalendarApp.getCalendarById(CAL_ID);
    if (body.action === 'ping') return out_({ok: true, calendar: cal.getName(), v: 1});
    if (body.action === 'backup') return out_(backup_(body));
    if (body.action === 'listBackups') return out_(listBackups_());
    if (body.action === 'getBackup') return out_(getBackup_(body.fileId));
    if (!body.id || !/^[\w-]+$/.test(body.id)) return out_({ok: false, error: 'Event id missing'});
    removeOld_(cal, body.id);
    if (body.action === 'delete') return out_({ok: true});
    if (body.action !== 'upsert' || !body.item) return out_({ok: false, error: 'Unknown action'});
    const ev = makeEvent_(cal, body.item);
    PropertiesService.getScriptProperties().setProperty('ev_' + body.id, JSON.stringify({id: ev.getId(), series: !!body.item.weekly}));
    return out_({ok: true});
  } catch (err) {
    return out_({ok: false, error: String(err && err.message || err).slice(0, 160)});
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Drive backup ---------- */
function folder_() {
  const it = DriveApp.getFoldersByName(BK_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(BK_FOLDER);
}
function backup_(body) {
  const data = String(body.data || '');
  if (data.length < 20 || data.charAt(0) !== '{') return {ok: false, error: 'No backup data'};
  const f = folder_(), name = 'bafsb-backup-' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd') + '.json';
  const ex = f.getFilesByName(name);
  if (ex.hasNext()) ex.next().setContent(data); else f.createFile(name, data, 'application/json');
  const files = [], it = f.getFiles();
  while (it.hasNext()) { const x = it.next(); if (/^bafsb-backup-\d{4}-\d{2}-\d{2}\.json$/.test(x.getName())) files.push(x); }
  files.sort((a, b) => b.getName() < a.getName() ? -1 : 1);
  files.slice(BK_KEEP).forEach(x => x.setTrashed(true));
  let sheetUrl = '';
  try { sheetUrl = writeSheet_(JSON.parse(data)); } catch (err) { sheetUrl = 'Sheet error: ' + err.message; }
  return {ok: true, name: name, size: data.length, sheet: sheetUrl};
}
function listBackups_() {
  const out = [], it = folder_().getFiles();
  while (it.hasNext()) { const x = it.next(); if (/^bafsb-backup-.*\.json$/.test(x.getName())) out.push({id: x.getId(), name: x.getName(), size: x.getSize(), at: x.getLastUpdated().getTime()}); }
  out.sort((a, b) => b.name < a.name ? -1 : 1);
  return {ok: true, files: out.slice(0, BK_KEEP)};
}
function getBackup_(fileId) {
  const f = folder_(), x = DriveApp.getFileById(String(fileId || ''));
  const parents = x.getParents(); let inside = false;
  while (parents.hasNext()) if (parents.next().getId() === f.getId()) inside = true;
  if (!inside) return {ok: false, error: 'This file is not in the backup folder'};
  return {ok: true, name: x.getName(), data: x.getBlob().getDataAsString('UTF-8')};
}

/* ---------- Google Sheet copy (one tab per class + Notes tab) ---------- */
function spreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('sheetId');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (x) {} }
  const ss = SpreadsheetApp.create(SHEET_NAME);
  DriveApp.getFileById(ss.getId()).moveTo(folder_());
  props.setProperty('sheetId', ss.getId());
  return ss;
}
function writeSheet_(db) {
  const ss = spreadsheet_();
  const exams = db.examDefs || [];
  const used = {};
  (db.classes || []).forEach(c => {
    const title = ((c.year || '') + ' ' + c.name + ' ' + (c.section || '') + ' ' + (c.subject || '')).replace(/[\[\]\*\?\/\\:]/g, ' ').trim().slice(0, 90) || 'Class';
    used[title] = true;
    const sh = ss.getSheetByName(title) || ss.insertSheet(title);
    const head = ['Year', 'Student ID', 'Roll', 'Name','Guardian phone', 'Phone 2', 'Father', 'Father occupation', 'Mother', 'Mother occupation', 'Address', 'Weakness', 'Status'];
    exams.forEach(x => x.cols.forEach(col => head.push(x.short + ' ' + col.label)));
    const rows = [head];
    (db.students || []).filter(s => s.classId === c.id).sort((a, b) => (Number(a.roll) || 0) - (Number(b.roll) || 0)).forEach(s => {
      const r = [c.year || '', s.sid || '', s.roll, s.name, s.phone, s.phone2, s.father, s.fOcc, s.mother, s.mOcc, s.address, s.weakness, s.archived ? 'Removed' : ''];
      exams.forEach(x => { const m = ((db.marks || {})[c.id + '|' + x.key] || {})[s.id] || {}; x.cols.forEach(col => r.push(m[col.k] == null ? '' : m[col.k])); });
      rows.push(r.map(v => v == null ? '' : v));
    });
    sh.clearContents();
    sh.getRange(1, 1, rows.length, head.length).setValues(rows);
    sh.getRange(1, 1, 1, head.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  const nTitle = 'Notes', nsh = ss.getSheetByName(nTitle) || ss.insertSheet(nTitle); used[nTitle] = true;
  const byId = {}; (db.students || []).forEach(s => byId[s.id] = s);
  const cById = {}; (db.classes || []).forEach(c => cById[c.id] = c);
  const nrows = [['Date', 'Class', 'Roll', 'Name', 'Type', 'Rating', 'Note']];
  (db.notes || []).slice().sort((a, b) => a.date < b.date ? 1 : -1).forEach(n => {
    const s = byId[n.studentId] || {}, c = cById[s.classId] || {};
    nrows.push([n.date, (c.name || '') + ' ' + (c.section || ''), s.roll || '', s.name || '', n.type, n.rating || '', n.text || '']);
  });
  nsh.clearContents();
  nsh.getRange(1, 1, nrows.length, 7).setValues(nrows);
  nsh.getRange(1, 1, 1, 7).setFontWeight('bold');
  ss.getSheets().forEach(sh => { if (!used[sh.getName()] && ss.getSheets().length > 1) ss.deleteSheet(sh); });
  return ss.getUrl();
}

/* ---------- Calendar ---------- */
function removeOld_(cal, id) {
  const props = PropertiesService.getScriptProperties(), k = 'ev_' + id;
  const old = JSON.parse(props.getProperty(k) || 'null');
  if (old && old.id) {
    try {
      if (old.series) { const s = cal.getEventSeriesById(old.id); if (s) s.deleteEventSeries(); }
      else { const ev = cal.getEventById(old.id); if (ev) ev.deleteEvent(); }
    } catch (x) {}
  }
  props.deleteProperty(k);
}

function parse_(s) { return Utilities.parseDate(s, TZ, "yyyy-MM-dd'T'HH:mm:ss"); }

/** item: {title, desc, allDay, date, start, end, weekly, until, popups:[minutes]} */
function makeEvent_(cal, it) {
  const title = String(it.title || 'School').slice(0, 200), opt = {description: String(it.desc || '').slice(0, 6000)};
  let ev;
  if (it.weekly) {
    const rec = CalendarApp.newRecurrence().addWeeklyRule();
    if (it.until) rec.until(Utilities.parseDate(it.until, TZ, 'yyyy-MM-dd'));
    ev = cal.createEventSeries(title, parse_(it.start), parse_(it.end), rec, opt);
  } else if (it.allDay) ev = cal.createAllDayEvent(title, Utilities.parseDate(it.date, TZ, 'yyyy-MM-dd'), opt);
  else ev = cal.createEvent(title, parse_(it.start), parse_(it.end), opt);
  try { if (it.color) ev.setColor(String(it.color)); } catch (x) {}
  ev.removeAllReminders();
  (it.popups || []).slice(0, 5).forEach(m => ev.addPopupReminder(Math.max(0, Math.min(40320, Number(m) || 0))));
  return ev;
}

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/** Press Run once: Google asks for permission (Calendar, Drive, Sheets), then shows a test in the log. */
function testSetup() {
  Logger.log('Calendar OK: ' + CalendarApp.getDefaultCalendar().getName());
  Logger.log('Backup folder: ' + folder_().getName());
  Logger.log('Sheet: ' + spreadsheet_().getUrl());
}
