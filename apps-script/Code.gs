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
 * - Lets the app tick items, take / give back items, and change Qty and Notes.
 * - Keeps family codes on a hidden tab called "App codes" (Family | Code).
 *   Change a family's code there any time. Default code is the last name followed by #.
 * - Keeps shared reminders on a tab called "App reminders" and a history of changes on "App log".
 */

var LIST_TAB = 'Thanksgiving 26 List';
var CODES_TAB = 'App codes';
var REMINDERS_TAB = 'App reminders';
var LOG_TAB = 'App log';
var FIXED_HEADERS = ['✓', 'Who', 'Section', 'Category', 'Brings', 'Store', 'Item', 'Qty', 'Cost', 'Notes'];

function doGet(e) {
  try {
    return json_(Object.assign({ ok: true }, readAll_()));
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    var result = apply_(body) || {};
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

function need_(L, name) {
  if (!(name in L.col)) throw new Error('The column "' + name + '" is missing from the header row.');
  return L.col[name];
}

function readAll_() {
  var sh = listSheet_();
  var L = layout_(sh);
  var first = L.headerRow + 1;
  var n = sh.getLastRow() - L.headerRow;
  var vals = n > 0 ? sh.getRange(first, 1, n, L.lastCol).getValues() : [];
  var disp = n > 0 ? sh.getRange(first, 1, n, L.lastCol).getDisplayValues() : [];
  var c = L.col;
  var get = function (d, name) { return name in c ? String(d[c[name]] || '').trim() : ''; };
  var items = [];
  for (var k = 0; k < vals.length; k++) {
    var item = get(disp[k], 'Item');
    if (!item) continue;
    var packed = {};
    L.families.forEach(function (f) { if (vals[k][f.i] === true) packed[f.name] = true; });
    items.push({
      row: first + k,
      item: item,
      section: get(disp[k], 'Section'),
      category: get(disp[k], 'Category'),
      brings: get(disp[k], 'Brings'),
      store: get(disp[k], 'Store'),
      qty: get(disp[k], 'Qty'),
      cost: get(disp[k], 'Cost'),
      notes: get(disp[k], 'Notes'),
      who: get(disp[k], 'Who'),
      done: ('✓' in c) ? vals[k][c['✓']] === true : false,
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
    log: readLog_(20),
    updated: new Date().toISOString()
  };
}

/* ---------- family codes ---------- */

function codesSheet_(families) {
  var sh = ss_().getSheetByName(CODES_TAB);
  if (!sh) {
    sh = ss_().insertSheet(CODES_TAB);
    sh.getRange(1, 1, 1, 2).setValues([['Family', 'Code']]);
    try { sh.hideSheet(); } catch (e) {}
  }
  var have = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues().map(function (r) { return String(r[0]).trim(); }) : [];
  (families || []).forEach(function (f) {
    if (have.indexOf(f) < 0) { sh.appendRow([f, f + '#']); have.push(f); }
  });
  return sh;
}

function codeOk_(family, code) {
  var L = layout_(listSheet_());
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

/* ---------- reminders + log ---------- */

function remindersSheet_() {
  var sh = ss_().getSheetByName(REMINDERS_TAB);
  if (!sh) {
    sh = ss_().insertSheet(REMINDERS_TAB);
    sh.getRange(1, 1, 1, 6).setValues([['Id', 'Date', 'Time', 'Reminder', 'For', 'Added by']]);
  }
  return sh;
}

function readReminders_() {
  var sh = ss_().getSheetByName(REMINDERS_TAB);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 6).getDisplayValues()
    .filter(function (r) { return String(r[3]).trim(); })
    .map(function (r) { return { id: String(r[0]), date: String(r[1]), time: String(r[2]), text: String(r[3]), for: String(r[4]) || 'Everyone', by: String(r[5]) }; });
}

function log_(family, what, item) {
  var sh = ss_().getSheetByName(LOG_TAB);
  if (!sh) {
    sh = ss_().insertSheet(LOG_TAB);
    sh.getRange(1, 1, 1, 4).setValues([['When', 'Family', 'What', 'Item']]);
  }
  sh.appendRow([new Date(), family, what, item || '']);
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

  if (b.action === 'addReminder') {
    var text = String(b.text || '').trim();
    if (!text) throw new Error('Write what to remind.');
    var rs = remindersSheet_();
    rs.appendRow([String(new Date().getTime()), "'" + String(b.date || ''), "'" + String(b.time || ''), text, String(b.for || 'Everyone'), family]);
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
  var L = layout_(sh);
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
      if ('notes' in b) {
        var nt = String(b.notes || '').trim();
        if (String(cell('Notes').getDisplayValue()).trim() !== nt) { cell('Notes').setValue(nt); changes.push('note'); }
      }
      if (changes.length) log_(family, 'Changed ' + changes.join(' and '), item);
      break;
    default:
      throw new Error('Unknown action.');
  }
  return {};
}
