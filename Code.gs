// נוכחות ושכר - קוד השרת (Google Apps Script)
// הקוד יוצר טבלה בגוגל שיטס בחשבון שלך ושומר בה את הכניסות והיציאות.

var DATA = 'נתונים';
var SETTINGS = 'הגדרות';
var SUMMARY = 'סיכום';
var HMONTHS = ['ינואר','פברואר','מרץ','אפריל','מאי','יוני','יולי','אוגוסט','ספטמבר','אוקטובר','נובמבר','דצמבר'];
var WDAYS = ['ראשון','שני','שלישי','רביעי','חמישי','שישי','שבת'];
var MONEY = '#,##0.00 "₪"';

/* ---------- logic (pure) ---------- */
function minute_(ms) { return ms == null || ms === '' ? null : Math.round(Number(ms) / 60000) * 60000; }

function dayName_(key) {
  var p = key.split('-');
  return WDAYS[new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2])).getDay()];
}

// records: [{date, heb, in, out, paid}] ; days: [{date, heb, shifts:[{in,out}], paid}]
function mergeDays_(records, days) {
  var touched = {};
  days.forEach(function (d) { touched[d.date] = true; });
  var out = records.filter(function (r) { return !touched[r.date]; });
  days.forEach(function (d) {
    (d.shifts || []).forEach(function (s) {
      if (s.in == null) return;
      out.push({ date: d.date, heb: d.heb || '', in: minute_(s.in), out: minute_(s.out), paid: !!d.paid });
    });
  });
  out.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : a.in - b.in; });
  return out;
}

function monthsOf_(records) {
  var seen = {}, list = [];
  records.forEach(function (r) { var m = r.date.slice(0, 7); if (!seen[m]) { seen[m] = true; list.push(m); } });
  list.sort();
  return list;
}

function monthName_(m) { return HMONTHS[Number(m.slice(5, 7)) - 1] + ' ' + m.slice(0, 4); }

/* ---------- spreadsheet ---------- */
function getSS_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SSID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  var ss = SpreadsheetApp.create('נוכחות ושכר');
  ss.setSpreadsheetTimeZone('Asia/Jerusalem');
  props.setProperty('SSID', ss.getId());

  var data = ss.getSheets()[0];
  data.setName(DATA);
  data.setRightToLeft(true);
  var head = ['תאריך', 'תאריך עברי', 'יום', 'כניסה', 'יציאה', 'זמן עבודה', 'שכר', 'שולם?', 'חודש'];
  data.getRange(1, 1, 1, 9).setValues([head]).setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff').setHorizontalAlignment('center');
  data.getRange('A:A').setNumberFormat('@');
  data.getRange('D:E').setNumberFormat('dd/MM hh:mm');
  data.getRange('F:F').setNumberFormat('[h]:mm');
  data.getRange('G:G').setNumberFormat(MONEY);
  data.setFrozenRows(1);
  data.setColumnWidths(1, 9, 110);
  data.setColumnWidth(2, 150);

  var st = ss.insertSheet(SETTINGS);
  st.setRightToLeft(true);
  st.getRange('A1:B1').setValues([['שכר לשעה', 45]]);
  st.getRange('A1').setFontWeight('bold');
  st.getRange('B1').setBackground('#fff2cc').setFontWeight('bold');
  st.getRange('A2').setValue('אפשר לשנות את המספר הצהוב, והשכר מתעדכן בכל הטבלאות.');
  st.setColumnWidth(1, 160);

  var sm = ss.insertSheet(SUMMARY, 0);
  sm.setRightToLeft(true);
  return ss;
}

function readRecords_(ss) {
  var sh = ss.getSheetByName(DATA);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, 8).getValues();
  var out = [];
  vals.forEach(function (v) {
    if (!v[0]) return;
    var date = v[0] instanceof Date ? Utilities.formatDate(v[0], 'Asia/Jerusalem', 'yyyy-MM-dd') : String(v[0]);
    out.push({
      date: date, heb: String(v[1] || ''),
      in: v[3] instanceof Date ? minute_(v[3].getTime()) : null,
      out: v[4] instanceof Date ? minute_(v[4].getTime()) : null,
      paid: String(v[7]) === 'שולם'
    });
  });
  return out.filter(function (r) { return r.in != null; });
}

function writeRecords_(ss, records) {
  var sh = ss.getSheetByName(DATA);
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 9).clearContent();
  if (!records.length) return;
  var rows = records.map(function (r, i) {
    var n = i + 2;
    return [
      r.date, r.heb, dayName_(r.date), new Date(r.in), r.out != null ? new Date(r.out) : '',
      '=IF(AND(ISNUMBER(D' + n + '),ISNUMBER(E' + n + ')),E' + n + '-D' + n + ',"")',
      '=IF(F' + n + '="","",ROUND(ROUND(F' + n + '*1440,0)*\'' + SETTINGS + '\'!$B$1/60,2))',
      r.paid ? 'שולם' : '',
      '=LEFT(A' + n + ',7)'
    ];
  });
  var rng = sh.getRange(2, 1, rows.length, 9);
  rng.setNumberFormat('General');
  sh.getRange(2, 1, rows.length, 1).setNumberFormat('@');
  rng.setValues(rows);
  sh.getRange(2, 4, rows.length, 2).setNumberFormat('dd/MM hh:mm');
  sh.getRange(2, 6, rows.length, 1).setNumberFormat('[h]:mm');
  sh.getRange(2, 7, rows.length, 1).setNumberFormat(MONEY);
  sh.getRange(2, 1, rows.length, 9).setHorizontalAlignment('center');
}

function ref_(col) { return "'" + DATA + "'!" + col + ':' + col; }

function ensureTabs_(ss, months) {
  var want = {};
  months.forEach(function (m) { want[monthName_(m)] = m; });
  var pattern = new RegExp('^(' + HMONTHS.join('|') + ') \\d{4}$');
  ss.getSheets().forEach(function (sh) {
    if (pattern.test(sh.getName()) && !want[sh.getName()]) ss.deleteSheet(sh);
  });
  months.forEach(function (m, idx) {
    var name = monthName_(m);
    if (ss.getSheetByName(name)) return;
    var sh = ss.insertSheet(name, 1 + idx);
    sh.setRightToLeft(true);
    sh.getRange('A1').setNumberFormat('@').setValue(name).setFontWeight('bold').setFontSize(14);
    var labels = [['סה״כ זמן'], ['סה״כ לתשלום'], ['כבר שולם'], ['נשאר לתשלום']];
    sh.getRange('A2:A5').setValues(labels).setFontWeight('bold');
    sh.getRange('B2').setFormula('=SUMIFS(' + ref_('F') + ',' + ref_('I') + ',"' + m + '")').setNumberFormat('[h]:mm');
    sh.getRange('B3').setFormula('=SUMIFS(' + ref_('G') + ',' + ref_('I') + ',"' + m + '")').setNumberFormat(MONEY);
    sh.getRange('B4').setFormula('=SUMIFS(' + ref_('G') + ',' + ref_('I') + ',"' + m + '",' + ref_('H') + ',"שולם")').setNumberFormat(MONEY);
    sh.getRange('B5').setFormula('=B3-B4').setNumberFormat(MONEY);
    sh.getRange('B2:B5').setBackground('#ddebf7').setFontWeight('bold').setHorizontalAlignment('center');
    var heads = ['תאריך', 'תאריך עברי', 'יום', 'כניסה', 'יציאה', 'זמן עבודה', 'שכר', 'שולם?'];
    sh.getRange(7, 1, 1, 8).setValues([heads]).setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff').setHorizontalAlignment('center');
    sh.getRange('A8').setFormula('=IFERROR(FILTER(\'' + DATA + '\'!A2:H,\'' + DATA + '\'!I2:I="' + m + '"),"")');
    sh.getRange('D8:E').setNumberFormat('hh:mm');
    sh.getRange('F8:F').setNumberFormat('[h]:mm');
    sh.getRange('G8:G').setNumberFormat(MONEY);
    sh.getRange('A8:H').setHorizontalAlignment('center');
    sh.setColumnWidths(1, 8, 105);
    sh.setColumnWidth(2, 150);
    sh.setColumnWidth(1, 120);
  });
}

function buildSummary_(ss, months) {
  var sh = ss.getSheetByName(SUMMARY);
  sh.clear();
  sh.getRange('A1:E1').setValues([['חודש', 'סה״כ זמן', 'סה״כ לתשלום', 'שולם', 'נשאר לתשלום']]).setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff').setHorizontalAlignment('center');
  if (!months.length) return;
  var rows = months.map(function (m, i) {
    var n = i + 2;
    return [
      monthName_(m),
      '=SUMIFS(' + ref_('F') + ',' + ref_('I') + ',"' + m + '")',
      '=SUMIFS(' + ref_('G') + ',' + ref_('I') + ',"' + m + '")',
      '=SUMIFS(' + ref_('G') + ',' + ref_('I') + ',"' + m + '",' + ref_('H') + ',"שולם")',
      '=C' + n + '-D' + n
    ];
  });
  sh.getRange(2, 1, rows.length, 1).setNumberFormat('@');
  sh.getRange(2, 1, rows.length, 5).setValues(rows);
  var t = rows.length + 2;
  sh.getRange(t, 1, 1, 5).setValues([['סה״כ', '=SUM(B2:B' + (t - 1) + ')', '=SUM(C2:C' + (t - 1) + ')', '=SUM(D2:D' + (t - 1) + ')', '=SUM(E2:E' + (t - 1) + ')']]).setFontWeight('bold').setBackground('#ddebf7');
  sh.getRange(2, 2, rows.length + 1, 1).setNumberFormat('[h]:mm');
  sh.getRange(2, 3, rows.length + 1, 3).setNumberFormat(MONEY);
  sh.getRange(2, 1, rows.length + 1, 5).setHorizontalAlignment('center');
  sh.setColumnWidths(1, 5, 120);
}

function state_(ss, records) {
  var rate = Number(ss.getSheetByName(SETTINGS).getRange('B1').getValue()) || 45;
  return { ok: true, rate: rate, url: ss.getUrl(), shifts: records };
}

function handle_(req) {
  var ss = getSS_();
  if (req.action === 'setRate') {
    var v = Number(req.rate);
    if (!(v > 0)) throw new Error('bad rate');
    ss.getSheetByName(SETTINGS).getRange('B1').setValue(v);
    return state_(ss, readRecords_(ss));
  }
  var records = readRecords_(ss);
  if (req.action === 'saveDays') {
    records = mergeDays_(records, req.days || []);
    writeRecords_(ss, records);
    var months = monthsOf_(records);
    ensureTabs_(ss, months);
    buildSummary_(ss, months);
  }
  return state_(ss, records);
}

/* ---------- web endpoints ---------- */
function doPost(e) {
  var out;
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    out = handle_(JSON.parse(e.postData.contents));
  } catch (err) {
    out = { ok: false, error: String(err) };
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService.createTextOutput('נוכחות ושכר: השרת עובד');
}

// מריצים פעם אחת ידנית כדי לאשר הרשאות וליצור את הטבלה
function setup() {
  var ss = getSS_();
  Logger.log('הטבלה נוצרה: ' + ss.getUrl());
}
