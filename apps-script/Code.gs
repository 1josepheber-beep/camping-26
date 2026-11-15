/**
 * Thanksgiving Camping app – sheet connector.
 *
 * Paste this whole file into Extensions > Apps Script in the CAMPING sheet,
 * then Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 *
 * What it does:
 * - Reads the list from the tab named in LIST_TAB, finding columns by their header names
 *   (✓, Who, Section, Category, Brings, Store, Item, Qty, Cost, Notes). Every other header
 *   in that row is treated as a family checkbox column.
 * - Lets the app tick items, take / give back items, add items, and change Qty, Notes and Cost.
 * - Keeps family codes on a tab called "App codes" (Family | Code).
 *   Change a family's code there any time. Default code is the last name followed by #.
 * - Keeps shared reminders on a tab called "App reminders" and a history of changes on "App log".
 * - Remembers which packing items a family doesn't need, on "App not needed" (Family | Item).
 * - Emails a reminder to the person who set it, at the date and time they chose (Houston time).
 *   Turn this on once by running setupEmailReminders from the editor.
 */

var LIST_TAB = 'Thanksgiving 26 List';
var CODES_TAB = 'App codes';
var REMINDERS_TAB = 'App reminders';
var LOG_TAB = 'App log';
var NOT_NEEDED_TAB = 'App not needed';
var FIXED_HEADERS = ['✓', 'Who', 'Section', 'Category', 'Brings', 'Store', 'Item', 'Qty', 'Cost', 'Notes'];
var MAX_BATCH = 200;

// Kept only while one request runs, so each save reads the sheet as little as possible.
var LAYOUT_ = null, LOG_BUF_ = null;
function resetRun_() { LAYOUT_ = null; LOG_BUF_ = null; }

function doGet(e) {
  resetRun_();
  try {
    return json_(Object.assign({ ok: true }, readAll_()));
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  resetRun_();
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    var result = apply_(body) || {};
    // A backup save sent while the app was closing doesn't need the list back.
    if (body.quiet) return json_(Object.assign({ ok: true }, result));
    return json_(Object.assign({ ok: true }, result, readAll_()));
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function ss_() { return SpreadsheetApp.getActive(); }

function listSheet_() {
  var sh = ss_().getSheetByName(LIST_TAB);
  if (!sh) throw new Error('The tab "' + LIST_TAB + '" is missing.');
  return sh;
}

function layout_(sh) {
  var lastCol = sh.getLastColumn();
  var topRows = Math.min(20, sh.getLastRow());
  var top = sh.getRange(1, 1, topRows, lastCol).getDisplayValues();
  var hr = -1;
  for (var r = 0; r < top.length; r++) {
    var rowTrim = top[r].map(function (s) { return String(s).trim(); });
    if (rowTrim.indexOf('Item') > -1 && rowTrim.indexOf('Who') > -1) { hr = r; break; }
  }
  if (hr < 0) throw new Error('Could not find the header row (it needs "Item" and "Who").');
  var head = top[hr].map(function (s) { return String(s).trim(); });
  var col = {};
  head.forEach(function (h, i) { if (h && !(h in col)) col[h] = i; });
  var families = [];
  head.forEach(function (h, i) { if (h && FIXED_HEADERS.indexOf(h) < 0) families.push({ name: h, i: i }); });
  return { headerRow: hr + 1, head: head, col: col, families: families, top: top.slice(0, hr), lastCol: lastCol };
}

function layoutNow_() { return LAYOUT_ || (LAYOUT_ = layout_(listSheet_())); }

function need_(L, name) {
  if (!(name in L.col)) throw new Error('The column "' + name + '" is missing from the header row.');
  return L.col[name];
}

function readAll_() {
  var sh = listSheet_();
  var L = layoutNow_();
  var first = L.headerRow + 1;
  var n = sh.getLastRow() - L.headerRow;
  var disp = n > 0 ? sh.getRange(first, 1, n, L.lastCol).getDisplayValues() : [];
  var isOn = function (v) { return String(v).trim().toUpperCase() === 'TRUE'; };
  var c = L.col;
  var get = function (d, name) { return name in c ? String(d[c[name]] || '').trim() : ''; };
  var items = [];
  for (var k = 0; k < disp.length; k++) {
    var item = get(disp[k], 'Item');
    if (!item) continue;
    var packed = {};
    L.families.forEach(function (f) { if (isOn(disp[k][f.i])) packed[f.name] = true; });
    items.push({
      row: first + k,
      item: item,
      section: get(disp[k], 'Section'),
      category: get(disp[k], 'Category'),
      brings: get(disp[k], 'Brings'),
      store: get(disp[k], 'Store'),
      qty: get(disp[k], 'Qty'),
      cost: get(disp[k], 'Cost'),
      costNum: ('Cost' in c) ? toNum_(disp[k][c['Cost']]) : 0,
      notes: get(disp[k], 'Notes'),
      who: get(disp[k], 'Who'),
      done: ('✓' in c) ? isOn(disp[k][c['✓']]) : false,
      packed: packed
    });
  }
  var trip = {};
  var wanted = ['Adults', 'Kids', 'Breakfast', 'Lunch', 'Dinner'];
  L.top.forEach(function (row) {
    row.forEach(function (cell, i) {
      var key = String(cell).trim();
      if (wanted.indexOf(key) > -1 && !(key in trip)) {
        for (var j = i + 1; j < row.length; j++) {
          if (String(row[j]).trim()) { trip[key] = String(row[j]).trim(); break; }
        }
      }
    });
  });
  var title = L.top.length ? String(L.top[0][0] || '').trim() : '';
  return {
    title: title || LIST_TAB,
    families: L.families.map(function (f) { return f.name; }),
    items: items,
    trip: trip,
    reminders: readReminders_(),
    notNeeded: readNotNeeded_(),
    log: readLog_(200),
    updated: new Date().toISOString()
  };
}

/* ---------- family codes ---------- */

function codesSheet_(families) {
  var sh = ss_().getSheetByName(CODES_TAB);
  if (!sh) {
    sh = ss_().insertSheet(CODES_TAB, ss_().getSheets().length);
    sh.getRange(1, 1, 1, 2).setValues([['Family', 'Code']]);
  }
  var have = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues().map(function (r) { return String(r[0]).trim(); }) : [];
  (families || []).forEach(function (f) {
    if (have.indexOf(f) < 0) { sh.appendRow([f, f + '#']); have.push(f); }
  });
  return sh;
}

function codeOk_(family, code) {
  var L = layoutNow_();
  var fams = L.families.map(function (f) { return f.name; });
  if (fams.indexOf(family) < 0) return false;
  var sh = codesSheet_(fams);
  var rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues() : [];
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === family) {
      return String(rows[i][1]).trim().toLowerCase() === String(code || '').trim().toLowerCase() && String(rows[i][1]).trim() !== '';
    }
  }
  return false;
}

/* ---------- items a family doesn't need ---------- */

function notNeededSheet_() {
  var sh = ss_().getSheetByName(NOT_NEEDED_TAB);
  if (!sh) {
    sh = ss_().insertSheet(NOT_NEEDED_TAB, ss_().getSheets().length);
    sh.getRange(1, 1, 1, 3).setValues([['Family', 'Item', 'When']]);
  }
  return sh;
}

function readNotNeeded_() {
  var sh = ss_().getSheetByName(NOT_NEEDED_TAB);
  var out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues().forEach(function (r) {
    var f = String(r[0]).trim(), it = String(r[1]).trim();
    if (!f || !it) return;
    (out[f] = out[f] || []).push(it);
  });
  return out;
}

function setNotNeeded_(family, item, on) {
  var sh = notNeededSheet_();
  var rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues() : [];
  for (var i = rows.length - 1; i >= 0; i--) {
    if (String(rows[i][0]).trim() === family && String(rows[i][1]).trim() === item) {
      if (on) return; // already there
      sh.deleteRow(i + 2);
    }
  }
  if (on) sh.appendRow([family, item, new Date()]);
}

/* ---------- reminders + log ---------- */

var REMINDER_HEADERS = ['Id', 'Date', 'Time', 'Reminder', 'For', 'Added by', 'Email to', 'Email status'];
var TIME_ZONE = 'America/Chicago';
var APP_LINK = 'https://1josepheber-beep.github.io/camping-26/';

function remindersSheet_() {
  var sh = ss_().getSheetByName(REMINDERS_TAB);
  if (!sh) sh = ss_().insertSheet(REMINDERS_TAB, ss_().getSheets().length);
  var head = sh.getRange(1, 1, 1, REMINDER_HEADERS.length).getDisplayValues()[0];
  if (head.join('|') !== REMINDER_HEADERS.join('|')) sh.getRange(1, 1, 1, REMINDER_HEADERS.length).setValues([REMINDER_HEADERS]);
  return sh;
}

// The app only learns whether a reminder will be emailed and whether it went out. Email addresses stay in the sheet.
function readReminders_() {
  var sh = ss_().getSheetByName(REMINDERS_TAB);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, REMINDER_HEADERS.length).getDisplayValues()
    .filter(function (r) { return String(r[3]).trim(); })
    .map(function (r) {
      return {
        id: String(r[0]), date: String(r[1]), time: String(r[2]), text: String(r[3]),
        for: String(r[4]) || 'Everyone', by: String(r[5]),
        email: String(r[6]).trim() ? true : false,
        emailStatus: String(r[7] || '')
      };
    });
}

function cleanEmails_(s) {
  var list = String(s || '').split(/[,;\s]+/).map(function (x) { return x.trim(); }).filter(String);
  list.forEach(function (x) {
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x)) throw new Error('"' + x + '" doesn\'t look like an email address.');
  });
  return list.join(', ');
}

function dueDate_(date, time) {
  var t = /^\d{1,2}:\d{2}/.test(String(time || '')) ? String(time).slice(0, 5) : '09:00';
  return Utilities.parseDate(String(date).slice(0, 10) + ' ' + t, TIME_ZONE, 'yyyy-MM-dd HH:mm');
}

/**
 * Runs every 5 minutes (set up by setupEmailReminders). Emails each reminder to the address
 * the person typed, once its date and time has come, and marks it Sent.
 */
function sendDueReminders() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;
  try {
    var sh = ss_().getSheetByName(REMINDERS_TAB);
    if (!sh || sh.getLastRow() < 2) return;
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, REMINDER_HEADERS.length).getDisplayValues();
    var now = new Date();
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var to = String(r[6]).trim(), status = String(r[7]).trim();
      if (!to || !String(r[3]).trim() || /^(Sent|Skipped)/.test(status)) continue;
      var due;
      try { due = dueDate_(r[1], r[2]); } catch (e) { sh.getRange(i + 2, 8).setValue('Error: bad date'); continue; }
      if (due.getTime() > now.getTime()) continue;
      if (now.getTime() - due.getTime() > 24 * 3600 * 1000) { sh.getRange(i + 2, 8).setValue('Skipped: time had already passed'); continue; }
      var when = Utilities.formatDate(due, TIME_ZONE, "EEE, MMM d 'at' h:mm a");
      var forTxt = r[4] && r[4] !== 'Everyone' ? 'For the ' + r[4] + ' family' : 'For everyone';
      try {
        MailApp.sendEmail({
          to: to,
          name: 'Thanksgiving Camping',
          subject: 'Camping reminder: ' + r[3],
          body: r[3] + '\n\n' + when + '\n' + forTxt + '\n\nOpen the camping list: ' + APP_LINK,
          htmlBody: '<div style="font-family:Arial,sans-serif;font-size:16px;color:#1D3A2C">' +
            '<p style="font-size:20px;font-weight:bold;margin:0 0 8px">' + esc_(r[3]) + '</p>' +
            '<p style="margin:0 0 4px">' + esc_(when) + '</p><p style="margin:0 0 16px;color:#5E6E63">' + esc_(forTxt) + '</p>' +
            '<p><a href="' + APP_LINK + '" style="color:#2D6BB0">Open the camping list</a></p></div>'
        });
        sh.getRange(i + 2, 8).setValue('Sent ' + Utilities.formatDate(new Date(), TIME_ZONE, 'MMM d, h:mm a'));
      } catch (err) {
        sh.getRange(i + 2, 8).setValue('Error: ' + String(err && err.message || err).slice(0, 120));
      }
    }
  } finally {
    lock.releaseLock();
  }
}

function esc_(s) {
  return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; });
}

/** Run this once from the Apps Script editor. It makes sendDueReminders run every 5 minutes. */
function setupEmailReminders() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendDueReminders') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendDueReminders').timeBased().everyMinutes(5).create();
  remindersSheet_();
  return 'Email reminders are on. Checking every 5 minutes.';
}

function logSheet_() {
  var sh = ss_().getSheetByName(LOG_TAB);
  if (!sh) {
    sh = ss_().insertSheet(LOG_TAB, ss_().getSheets().length);
    sh.getRange(1, 1, 1, 4).setValues([['When', 'Family', 'What', 'Item']]);
  }
  return sh;
}

function log_(family, what, item) {
  var row = [new Date(), family, what, item || ''];
  if (LOG_BUF_) { LOG_BUF_.push(row); return; }
  logSheet_().appendRow(row);
}

// During a batch, history lines are written together at the end in one go.
function flushLog_() {
  var rows = LOG_BUF_; LOG_BUF_ = null;
  if (!rows || !rows.length) return;
  var sh = logSheet_();
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
}

function readLog_(max) {
  var sh = ss_().getSheetByName(LOG_TAB);
  if (!sh || sh.getLastRow() < 2) return [];
  var n = Math.min(max, sh.getLastRow() - 1);
  var start = sh.getLastRow() - n + 1;
  var rows = sh.getRange(start, 1, n, 4).getValues();
  return rows.reverse().map(function (r) {
    return { when: r[0] instanceof Date ? r[0].toISOString() : String(r[0]), family: String(r[1]), what: String(r[2]), item: String(r[3]) };
  });
}

/* ---------- changes ---------- */

function findRow_(sh, L, row, item) {
  var itemCol = need_(L, 'Item') + 1;
  if (row && row > L.headerRow && row <= sh.getLastRow()) {
    if (String(sh.getRange(row, itemCol).getDisplayValue()).trim() === item) return row;
  }
  var n = sh.getLastRow() - L.headerRow;
  if (n > 0) {
    var names = sh.getRange(L.headerRow + 1, itemCol, n, 1).getDisplayValues();
    for (var k = 0; k < names.length; k++) {
      if (String(names[k][0]).trim() === item) return L.headerRow + 1 + k;
    }
  }
  throw new Error('"' + item + '" is not on the list anymore. Pull down to refresh.');
}

function toNum_(v) {
  if (typeof v === 'number') return v;
  var n = parseFloat(String(v == null ? '' : v).replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

function asText_(v) {
  var s = String(v == null ? '' : v).trim();
  // Keep things like 1/2 or 3-4 from turning into dates.
  if (/^[\d\s]+([\/\-][\d\s]+)+$/.test(s)) return "'" + s;
  return s;
}

function apply_(b) {
  var family = String(b.family || '').trim();
  if (b.action === 'login') {
    if (!codeOk_(family, b.code)) throw new Error('That code is not right for ' + family + '.');
    return { loggedIn: family };
  }
  if (!codeOk_(family, b.code)) throw new Error('CODE: Your family code changed. Enter the new one.');
  if (b.action === 'batch') return applyBatch_(family, b.changes);
  return applyOne_(family, b);
}

/**
 * Several changes from one family in one trip (the app's Save button).
 * Each change has an id, so a change that arrives twice (a retry, or the
 * backup save sent when the app closes) is only applied once.
 */
function applyBatch_(family, list) {
  list = Array.isArray(list) ? list.slice(0, MAX_BATCH) : [];
  var cache = null;
  try { cache = CacheService.getScriptCache(); } catch (e) {}
  var keys = [];
  list.forEach(function (ch) { if (ch && ch.cid) keys.push('cid:' + ch.cid); });
  var seen = {};
  if (cache && keys.length) { try { seen = cache.getAll(keys) || {}; } catch (e) { seen = {}; } }
  var applied = {}, errors = [];
  LOG_BUF_ = [];
  try {
    list.forEach(function (ch) {
      if (!ch || typeof ch !== 'object') return;
      var key = ch.cid ? 'cid:' + ch.cid : '';
      if (key && (seen[key] || applied[key])) return;
      try { applyOne_(family, ch); } catch (err) { errors.push(String(err && err.message || err)); }
      if (key) applied[key] = '1';
    });
  } finally {
    flushLog_();
  }
  if (cache && Object.keys(applied).length) { try { cache.putAll(applied, 21600); } catch (e) {} }
  return { errors: errors };
}

function applyOne_(family, b) {
  if (b.action === 'addReminder') {
    var text = String(b.text || '').trim();
    if (!text) throw new Error('Write what to remind.');
    var emails = cleanEmails_(b.email);
    if (emails) {
      var due = dueDate_(b.date, b.time);
      if (due.getTime() < new Date().getTime() - 60 * 1000) throw new Error('That time has already passed. Pick a later time to get an email.');
    }
    var rs = remindersSheet_();
    rs.appendRow([String(new Date().getTime()), "'" + String(b.date || ''), "'" + String(b.time || ''), text, String(b.for || 'Everyone'), family, emails, emails ? 'Waiting' : '']);
    log_(family, 'Added reminder', text);
    return {};
  }
  if (b.action === 'deleteReminder') {
    var rs2 = ss_().getSheetByName(REMINDERS_TAB);
    if (rs2 && rs2.getLastRow() > 1) {
      var ids = rs2.getRange(2, 1, rs2.getLastRow() - 1, 1).getDisplayValues();
      for (var i = ids.length - 1; i >= 0; i--) {
        if (String(ids[i][0]) === String(b.id)) {
          var txt = rs2.getRange(i + 2, 4).getDisplayValue();
          rs2.deleteRow(i + 2);
          log_(family, 'Removed reminder', txt);
          break;
        }
      }
    }
    return {};
  }

  var sh = listSheet_();
  var L = layoutNow_();
  var item = String(b.item || '').trim();

  if (b.action === 'addItem') {
    if (!item) throw new Error('Write the item name.');
    var c0 = L.col;
    var itemCol = need_(L, 'Item');
    var n = sh.getLastRow() - L.headerRow;
    var disp = n > 0 ? sh.getRange(L.headerRow + 1, 1, n, L.lastCol).getDisplayValues() : [];
    var section = String(b.section || '').trim();
    var after = -1, lastItemRow = -1;
    for (var k = 0; k < disp.length; k++) {
      if (String(disp[k][itemCol]).trim()) {
        lastItemRow = L.headerRow + 1 + k;
        if (section && 'Section' in c0 && String(disp[k][c0['Section']]).trim() === section) after = lastItemRow;
      }
    }
    if (after < 0) after = lastItemRow > 0 ? lastItemRow : L.headerRow + 1;
    sh.insertRowAfter(after);
    var nr = after + 1;
    var src = sh.getRange(after, 1, 1, L.lastCol), dst = sh.getRange(nr, 1, 1, L.lastCol);
    src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    var put = function (name, val) { if (name in c0) sh.getRange(nr, c0[name] + 1).setValue(val); };
    put('Section', section);
    put('Category', String(b.category || '').trim());
    put('Brings', String(b.brings || 'Group').trim());
    put('Store', String(b.store || '').trim());
    put('Item', item);
    put('Qty', asText_(b.qty));
    put('Notes', String(b.notes || '').trim());
    if (b.mine && String(b.brings || 'Group').trim() !== 'Each family') put('Who', family);
    log_(family, 'Added', item);
    return {};
  }

  if (b.action === 'notNeeded') {
    findRow_(sh, L, Number(b.row), item); // make sure it's still on the list
    setNotNeeded_(family, item, !!b.value);
    log_(family, b.value ? "Doesn't need" : 'Needs after all', item);
    return {};
  }

  var r = findRow_(sh, L, Number(b.row), item);
  var cell = function (name) { return sh.getRange(r, need_(L, name) + 1); };

  switch (b.action) {
    case 'done':
      cell('✓').setValue(!!b.value);
      log_(family, b.value ? 'Marked bought' : 'Unmarked bought', item);
      break;
    case 'pack':
      var f = L.families.filter(function (x) { return x.name === family; })[0];
      if (!f) throw new Error('No column for ' + family + ' in the sheet.');
      sh.getRange(r, f.i + 1).setValue(!!b.value);
      log_(family, b.value ? 'Packed' : 'Unpacked', item);
      break;
    case 'take':
      var cur = String(cell('Who').getDisplayValue()).trim();
      if (cur && cur !== family) throw new Error(cur + ' already took "' + item + '".');
      cell('Who').setValue(family);
      log_(family, 'Took', item);
      break;
    case 'release':
      var cur2 = String(cell('Who').getDisplayValue()).trim();
      if (cur2 && cur2 !== family) throw new Error('Only ' + cur2 + ' can give back "' + item + '".');
      cell('Who').setValue('');
      log_(family, 'Gave back', item);
      break;
    case 'edit':
      var changes = [];
      if ('qty' in b) {
        var q = asText_(b.qty);
        if (String(cell('Qty').getDisplayValue()).trim() !== q.replace(/^'/, '')) { cell('Qty').setValue(q); changes.push('amount to "' + q.replace(/^'/, '') + '"'); }
      }
      if ('store' in b) {
        var st = String(b.store == null ? '' : b.store).trim();
        if (String(cell('Store').getDisplayValue()).trim() !== st) { cell('Store').setValue(st); changes.push(st ? 'store to ' + (st === 'N/A' ? 'no store needed' : st) : 'store (cleared)'); }
      }
      if ('notes' in b) {
        var nt = String(b.notes || '').trim();
        if (String(cell('Notes').getDisplayValue()).trim() !== nt) { cell('Notes').setValue(nt); changes.push('note'); }
      }
      if ('cost' in b) {
        var raw = String(b.cost == null ? '' : b.cost).replace(/[^0-9.\-]/g, '');
        var newCost = raw === '' ? '' : Math.round(parseFloat(raw) * 100) / 100;
        if (newCost !== '' && isNaN(newCost)) throw new Error('The cost has to be a number, like 12.50');
        var oldCost = cell('Cost').getValue();
        if (newCost === '' ? oldCost !== '' : toNum_(oldCost) !== newCost || oldCost === '') {
          cell('Cost').setValue(newCost);
          changes.push(newCost === '' ? 'cost (cleared)' : 'cost to $' + newCost.toFixed(2));
        }
      }
      if (changes.length) log_(family, 'Changed ' + changes.join(' and '), item);
      break;
    default:
      throw new Error('Unknown action.');
  }
  return {};
}
