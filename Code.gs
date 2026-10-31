/**
 * TWC Disposals — Google backend for https://jen826.github.io/twc-disposals/
 *
 * The "Disposal Log" tab is the register: one row per item type per location, readable and filterable.
 * Prefer editing in the app; if you edit a row here, keep the ID column as it is.
 * "Lists" holds the dropdown choices (add a property, item or vendor there and the app picks it up).
 * "Settings" holds the access code. Photos are saved in a Drive folder next to this spreadsheet.
 *
 * Setup: Extensions > Apps Script, paste this file, run setup(), then Deploy > New deployment > Web app
 * (Execute as: Me, Who has access: Anyone).
 */

var LOG = 'Disposal Log';
var COLS = [
  ['id', 'ID'], ['requested', 'Date Requested'], ['property', 'Property'], ['location', 'Unit / Location'],
  ['item', 'Item'], ['description', 'Description'], ['qty', 'Qty'], ['condition', 'Condition'], ['reason', 'Reason'],
  ['method', 'Method'], ['vendor', 'Vendor / Contractor'], ['planned', 'Planned Date'], ['completed', 'Completed Date'],
  ['cost', 'Cost (SGD)'], ['status', 'Status'], ['photos', 'Photos'], ['remarks', 'Remarks'],
  ['requestedBy', 'Requested By'], ['updatedAt', 'Last Updated'], ['updatedBy', 'Updated By']
];
var DATE_KEYS = { requested: 1, planned: 1, completed: 1 };
var DEFAULT_LISTS = {
  'Property': ['1 KS', '89 NR', '40 BP', '219 Kallang Bahru'],
  'Item': ['Table', 'Chair', 'Cabinet', 'Pedestal', 'Sofa', 'Partition', 'Appliance', 'IT equipment', 'Others'],
  'Method': ['Dispose', 'Recycle', 'Sell', 'Donate', 'Sent to storage', 'Return to vendor'],
  'Condition': ['Damaged', 'Worn', 'Obsolete', 'Good – surplus'],
  'Status': ['Pending approval', 'Approved', 'Scheduled', 'Completed', 'Cancelled'],
  'Vendor / Contractor': ['SuperHandy']
};

/* ---------------- Sheet setup ---------------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Disposals').addItem('Set up sheets', 'setup').addToUi();
}
function setup() {
  ensureSheets_();
  try { SpreadsheetApp.getActive().toast('Ready. Next: Deploy > New deployment > Web app.', 'Disposals', 8); } catch (e) {}
}
function ss_() { return SpreadsheetApp.getActive(); }

function ensureSheets_() {
  var ss = ss_();
  if (!ss.getSheetByName(LOG)) {
    var s = ss.insertSheet(LOG, 0);
    s.getRange(1, 1, 1, COLS.length).setValues([COLS.map(function (c) { return c[1]; })])
      .setFontWeight('bold').setBackground('#1b2127').setFontColor('#ffffff');
    s.setFrozenRows(1);
    s.getRange(2, 1, 999, COLS.length).setNumberFormat('@');
    s.getRange(2, colOf_('qty'), 999, 1).setNumberFormat('0');
    s.getRange(2, colOf_('cost'), 999, 1).setNumberFormat('#,##0.00');
    [90, 110, 90, 200, 100, 180, 50, 110, 220, 120, 150, 110, 110, 90, 120, 260, 220, 120, 150, 110].forEach(function (w, i) { s.setColumnWidth(i + 1, w); });
  }
  if (!ss.getSheetByName('Lists')) {
    var l = ss.insertSheet('Lists');
    var keys = Object.keys(DEFAULT_LISTS), n = 0;
    keys.forEach(function (k) { n = Math.max(n, DEFAULT_LISTS[k].length); });
    var rows = [keys];
    for (var i = 0; i < n; i++) rows.push(keys.map(function (k) { return DEFAULT_LISTS[k][i] || ''; }));
    l.getRange(1, 1, rows.length, keys.length).setValues(rows);
    l.getRange(1, 1, 1, keys.length).setFontWeight('bold').setBackground('#1b2127').setFontColor('#ffffff');
    l.setFrozenRows(1);
  }
  if (!ss.getSheetByName('Settings')) {
    var st = ss.insertSheet('Settings');
    st.getRange(1, 1, 4, 2).setValues([
      ['Setting', 'Value'], ['Access code', ''], ['', ''],
      ['About', 'Leave Access code blank so anyone with the link can use the app, or enter a code that staff type once per device.']]);
    st.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#1b2127').setFontColor('#ffffff');
    st.setColumnWidth(1, 160); st.setColumnWidth(2, 520);
  }
  var def = ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);
}
function colOf_(key) { for (var i = 0; i < COLS.length; i++) if (COLS[i][0] === key) return i + 1; return -1; }

function accessCode_() {
  var s = ss_().getSheetByName('Settings');
  if (!s) return '';
  var v = s.getRange(1, 1, Math.max(1, s.getLastRow()), 2).getValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][0]).trim().toLowerCase() === 'access code') return String(v[i][1] || '').trim();
  return '';
}
function check_(code) {
  var need = accessCode_();
  if (need && String(code || '').trim() !== need) throw new Error('ACCESS_DENIED');
}

/* ---------------- Rows <-> records ---------------- */

function iso_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var s = String(v || '').trim();
  return s;
}
function photoIds_(v) {
  var out = [], re = /([\w-]{20,})/g, m, s = String(v || '');
  while ((m = re.exec(s))) if (out.indexOf(m[1]) < 0) out.push(m[1]);
  return out;
}
function toRec_(row) {
  var r = {};
  COLS.forEach(function (c, i) {
    var v = row[i];
    if (c[0] === 'qty' || c[0] === 'cost') r[c[0]] = v === '' || v === null ? '' : Number(v) || 0;
    else if (c[0] === 'photos') r.photos = photoIds_(v);
    else if (DATE_KEYS[c[0]]) r[c[0]] = iso_(v);
    else if (c[0] === 'updatedAt') r.updatedAt = v instanceof Date ? v.getTime() : (Number(v) || Date.parse(v) || 0);
    else r[c[0]] = v === null ? '' : String(v);
  });
  return r;
}
function toRow_(r) {
  return COLS.map(function (c) {
    var v = r[c[0]];
    if (c[0] === 'photos') return (v || []).map(function (id) { return 'https://drive.google.com/file/d/' + id + '/view'; }).join('\n');
    if (c[0] === 'updatedAt') return Utilities.formatDate(new Date(v || Date.now()), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    if (c[0] === 'qty' || c[0] === 'cost') return v === '' || v === undefined || v === null ? '' : Number(v) || 0;
    return v === undefined || v === null ? '' : String(v);
  });
}
function readLog_() {
  var s = ss_().getSheetByName(LOG), n = s.getLastRow();
  if (n < 2) return [];
  return s.getRange(2, 1, n - 1, COLS.length).getValues()
    .map(function (row, i) { var r = toRec_(row); r._row = i + 2; return r; })
    .filter(function (r) { return r.id || r.item || r.property; });
}
function readLists_() {
  var s = ss_().getSheetByName('Lists'), out = {};
  if (!s || s.getLastRow() < 1) return DEFAULT_LISTS;
  var v = s.getRange(1, 1, s.getLastRow(), Math.max(1, s.getLastColumn())).getValues();
  v[0].forEach(function (h, j) {
    h = String(h || '').trim(); if (!h) return;
    out[h] = [];
    for (var i = 1; i < v.length; i++) { var x = String(v[i][j] || '').trim(); if (x && out[h].indexOf(x) < 0) out[h].push(x); }
  });
  return out;
}
function nextId_(rows) {
  var max = 0;
  rows.forEach(function (r) { var m = /(\d+)$/.exec(r.id || ''); if (m) max = Math.max(max, Number(m[1])); });
  return function () { max++; return 'D-' + ('0000' + max).slice(-4); };
}
function clean_(r) {
  var out = {};
  COLS.forEach(function (c) { out[c[0]] = r[c[0]]; });
  out.qty = r.qty === '' || r.qty === undefined ? '' : Math.max(0, Math.round(Number(r.qty) || 0));
  out.cost = r.cost === '' || r.cost === undefined || r.cost === null ? '' : Number(r.cost) || 0;
  out.photos = (r.photos || []).filter(function (p) { return /^[\w-]{20,}$/.test(p); }).slice(0, 12);
  ['property', 'location', 'item', 'description', 'condition', 'reason', 'method', 'vendor', 'status', 'remarks', 'requestedBy']
    .forEach(function (k) { out[k] = String(r[k] || '').slice(0, 1000); });
  Object.keys(DATE_KEYS).forEach(function (k) { out[k] = /^\d{4}-\d{2}-\d{2}$/.test(String(r[k] || '')) ? r[k] : ''; });
  return out;
}

/* ---------------- API ---------------- */

function apiBoot(code) {
  check_(code); ensureSheets_();
  return { rows: readLog_().map(strip_), lists: readLists_(), serverTime: Date.now(), sheetUrl: ss_().getUrl() };
}
function apiList(code) {
  check_(code);
  return { rows: readLog_().map(strip_), lists: readLists_(), serverTime: Date.now() };
}
function strip_(r) { delete r._row; return r; }

/** Adds or updates records. Each record without an id gets the next D-0001 style id. */
function apiSave(code, recs, who) {
  check_(code);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var s = ss_().getSheetByName(LOG), rows = readLog_(), next = nextId_(rows), now = Date.now(), saved = [];
    (recs || []).slice(0, 50).forEach(function (r0) {
      var r = clean_(r0); r.updatedAt = now; r.updatedBy = String(who || '').slice(0, 80);
      var cur = r0.id ? rows.filter(function (x) { return x.id === r0.id; })[0] : null;
      if (cur) {
        r.id = cur.id; r.requestedBy = cur.requestedBy || r.requestedBy;
        s.getRange(cur._row, 1, 1, COLS.length).setValues([toRow_(r)]);
      } else {
        r.id = next(); if (!r.requestedBy) r.requestedBy = r.updatedBy;
        if (!r.status) r.status = 'Pending approval';
        s.appendRow(toRow_(r));
        rows.push({ id: r.id, _row: s.getLastRow() });
      }
      saved.push(r);
    });
    return { rows: saved, serverTime: now };
  } finally { lock.releaseLock(); }
}
function apiDelete(code, id) {
  check_(code);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var cur = readLog_().filter(function (x) { return x.id === id; })[0];
    if (!cur) return false;
    ss_().getSheetByName(LOG).deleteRow(cur._row);
    return true;
  } finally { lock.releaseLock(); }
}

function folder_() {
  var props = PropertiesService.getScriptProperties(), id = props.getProperty('PHOTOS');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var parent;
  try { parent = DriveApp.getFileById(ss_().getId()).getParents().next(); } catch (e) { parent = DriveApp.getRootFolder(); }
  var f = parent.createFolder('TWC Disposal photos');
  props.setProperty('PHOTOS', f.getId());
  return f;
}
function apiUploadPhoto(code, dataUrl) {
  check_(code);
  var m = /^data:(image\/(?:jpeg|png|webp));base64,(.+)$/.exec(String(dataUrl || ''));
  if (!m) throw new Error('That file is not a supported photo. Please use a JPG or PNG.');
  var bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > 8 * 1024 * 1024) throw new Error('That photo is too large.');
  var f = folder_().createFile(Utilities.newBlob(bytes, m[1], 'disposal-' + Date.now() + (m[1] === 'image/png' ? '.png' : '.jpg')));
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
  catch (e) { try { f.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (e2) {} }
  return f.getId();
}
/** Photo bytes for browsers that can't load Drive thumbnails directly. */
function apiPhotoData(code, id) {
  check_(code);
  if (!/^[\w-]{20,}$/.test(String(id))) return '';
  try {
    var f = DriveApp.getFileById(id), b = f.getSize() > 2 * 1024 * 1024 ? f.getThumbnail() : f.getBlob();
    return b ? 'data:' + (b.getContentType() || 'image/jpeg') + ';base64,' + Utilities.base64Encode(b.getBytes()) : '';
  } catch (e) { return ''; }
}

var API_FUNCTIONS = { apiBoot: apiBoot, apiList: apiList, apiSave: apiSave, apiDelete: apiDelete, apiUploadPhoto: apiUploadPhoto, apiPhotoData: apiPhotoData };

function doGet() {
  return HtmlService.createHtmlOutput('<p style="font-family:sans-serif">TWC Disposals backend is running. Open the app at <a href="https://jen826.github.io/twc-disposals/">jen826.github.io/twc-disposals</a>.</p>');
}
function doPost(e) {
  var out;
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var f = API_FUNCTIONS[req.fn];
    if (!f) throw new Error('Unknown request');
    out = { ok: true, result: f.apply(null, req.args || []) };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
