// Diet Detectives Responses: receives worksheet answers from the GitHub Pages worksheets.
var SHEET_ID = '1I83RosqGgE0Suqm2MJQxfrSE6n5H8N6J_KJnitQ6Eis';
var TABS = { nutrients: 'Nutrients', lipids: 'Lipids' };
var FIXED = ['Student ID', 'Name', 'Class', 'Status', 'Started', 'Last update', 'Minutes on task', 'Exits', 'Exit times', 'Big inserts', 'Paste attempts'];

function setup() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  Object.keys(TABS).forEach(function (k) {
    var sh = ss.getSheetByName(TABS[k]) || ss.insertSheet(TABS[k]);
    if (sh.getLastRow() === 0) { sh.appendRow(FIXED); sh.setFrozenRows(1); sh.getRange(1, 1, 1, FIXED.length).setFontWeight('bold'); }
  });
  var def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);
  return 'ok';
}

function clean(v, max) {
  var s = String(v == null ? '' : v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '').slice(0, max || 3000);
  if (/^[=+\-@\t\r]/.test(s)) s = "'" + s;
  return s;
}

function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet(e) { return out({ ok: true, ping: 'diet-detectives' }); }

function doPost(e) {
  var d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return out({ ok: false, error: 'bad json' }); }
  if (!d || d.action !== 'save' || !TABS[d.version]) return out({ ok: false, error: 'bad request' });
  if (!/^S[a-z0-9]{6,24}$/.test(String(d.id))) return out({ ok: false, error: 'bad id' });
  var keys = Array.isArray(d.keys) ? d.keys.slice(0, 80) : [];
  var labels = Array.isArray(d.labels) ? d.labels.slice(0, 80) : [];
  var answers = d.answers || {};
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName(TABS[d.version]);
    if (!sh) { setup(); sh = ss.getSheetByName(TABS[d.version]); }
    var lastCol = Math.max(sh.getLastColumn(), FIXED.length);
    var header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    // Add answer columns the first time they are seen
    var added = [];
    keys.forEach(function (k, i) {
      var lab = clean(labels[i] || k, 120);
      if (header.indexOf(lab) === -1) { header.push(lab); added.push(lab); }
    });
    if (added.length) sh.getRange(1, header.length - added.length + 1, 1, added.length).setValues([added]).setFontWeight('bold');
    var row = header.map(function () { return ''; });
    var set = function (name, v) { var i = header.indexOf(name); if (i > -1) row[i] = v; };
    set('Student ID', clean(d.id, 40));
    set('Name', clean(d.name, 60));
    set('Class', clean(d.cls, 10));
    set('Status', clean(d.status, 20));
    set('Started', d.started ? new Date(d.started) : '');
    set('Last update', new Date());
    set('Minutes on task', Math.round((Number(d.elapsed) || 0) / 6) / 10);
    set('Exits', Number(d.exits) || 0);
    set('Exit times', clean(d.exitLog, 2000));
    set('Big inserts', Number(d.bigInserts) || 0);
    set('Paste attempts', Number(d.pasteTries) || 0);
    keys.forEach(function (k, i) { set(clean(labels[i] || k, 120), clean(answers[k], 3000)); });
    // Upsert by Student ID
    var n = sh.getLastRow(), at = -1;
    if (n > 1) {
      var ids = sh.getRange(2, 1, n - 1, 1).getValues();
      for (var r = 0; r < ids.length; r++) if (String(ids[r][0]) === String(d.id)) { at = r + 2; break; }
    }
    if (at === -1) at = n + 1;
    // Never let a late autosave downgrade a submission's exit count
    if (at <= n) {
      var exCol = header.indexOf('Exits') + 1;
      var prev = Number(sh.getRange(at, exCol).getValue()) || 0;
      if (prev > row[exCol - 1]) row[exCol - 1] = prev;
      var stCol = header.indexOf('Status') + 1;
      if (String(sh.getRange(at, stCol).getValue()) === 'Submitted' && row[stCol - 1] === 'In progress') return out({ ok: true, ignored: 'late autosave' });
    }
    sh.getRange(at, 1, 1, row.length).setValues([row]);
    if (row[header.indexOf('Exits')] >= 3) sh.getRange(at, header.indexOf('Exits') + 1).setBackground('#F9E1DE');
    return out({ ok: true });
  } finally {
    lock.releaseLock();
  }
}
