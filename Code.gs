/**
 * المخبر الاقتصادي — نظام التقييم 2026
 * Google Apps Script — Server Side
 *
 * Setup:
 *   1. Open your Google Sheet
 *   2. Extensions → Apps Script
 *   3. Paste this file as Code.gs
 *   4. Create a new HTML file called "Page" and paste Page.html
 *   5. Deploy → New deployment → Web app
 *      - Execute as: Me
 *      - Who has access: Anyone
 *   6. Open the deployment URL — done!
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Page')
    .setTitle('المخبر الاقتصادي — نظام التقييم 2026')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0');
}

/* ── DATA READ/WRITE ── */

function getData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('_data');
  if (!sh) return null;
  const raw = sh.getRange('A1').getValue();
  if (!raw) return null;
  try { return JSON.parse(raw); } catch(e) { return null; }
}

function saveData(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('_data') || ss.insertSheet('_data');
  sh.getRange('A1').setValue(JSON.stringify(data));
  sh.getRange('A2').setValue('آخر تحديث: ' + new Date().toLocaleString('ar-EG'));
  try { sh.hideSheet(); } catch(e) {}
  return {success: true};
}

/* ── FORMATTED SHEET EXPORT ── */

function exportToSheets(data) {
  if (!data || !data.notes) return {success: false, error: 'لا توجد بيانات'};
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Also save raw data
  let meta = ss.getSheetByName('_data') || ss.insertSheet('_data');
  meta.getRange('A1').setValue(JSON.stringify(data));
  meta.getRange('A2').setValue('آخر تحديث: ' + new Date().toLocaleString('ar-EG'));
  try { meta.hideSheet(); } catch(e) {}

  writeTeamSummary(ss, data);
  (data.members || []).forEach(function(m) {
    var name = typeof m === 'string' ? m : m.name;
    var role = typeof m === 'string' ? 'mod' : (m.role || 'mod');
    writeMemberSheet(ss, data, name, role);
  });
  reorderSheets(ss, data.members || []);
  return {success: true, count: data.notes.length};
}

/* ── TEAM SUMMARY SHEET ── */

function writeTeamSummary(ss, data) {
  var sh = ss.getSheetByName('📊 ملخص الفريق') || ss.insertSheet('📊 ملخص الفريق');
  sh.clearContents(); sh.clearFormats();
  sh.getCharts().forEach(function(c) { sh.removeChart(c); });
  sh.setRightToLeft(true);

  sh.getRange('A1:H1').merge();
  sh.getRange('A1').setValue('📊 ملخص فريق المخبر الاقتصادي — ' + new Date().toLocaleDateString('ar-EG'));
  sh.getRange('A1').setBackground('#1a1008').setFontColor('#e8b84a').setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
  sh.setRowHeight(1, 44);

  sh.getRange('A2:H2').merge();
  sh.getRange('A2').setValue('إجمالي الملاحظات: ' + data.notes.length + '   |   الأعضاء: ' + (data.members || []).length);
  sh.getRange('A2').setBackground('#221608').setFontColor('#887858').setFontSize(11).setHorizontalAlignment('center');

  var roleLabels = {mod: 'مودريشن', design: 'جرافيك', video: 'فيديو'};
  var headers = ['العضو', 'الوظيفة', 'الملاحظات', 'متوسط التقييم', 'أعلى معيار', 'إيجابي', 'تحسين', 'آخر تقييم'];
  sh.getRange(4, 1, 1, headers.length).setValues([headers])
    .setBackground('#b85840').setFontColor('#f2e8d8').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
  sh.setRowHeight(4, 30);

  var members = data.members || [];
  var rows = members.map(function(m) {
    var name = typeof m === 'string' ? m : m.name;
    var role = typeof m === 'string' ? 'mod' : (m.role || 'mod');
    var notes = data.notes.filter(function(n) { return n.member === name; });
    var scored = notes.filter(function(n) { return n.stars > 0; });
    var avg = scored.length ? Math.round(scored.reduce(function(a, n) { return a + n.stars; }, 0) / scored.length * 10) / 10 : 0;
    var last = notes[0] ? new Date(notes[0].date).toLocaleDateString('ar-EG') : '—';
    return [name, roleLabels[role] || role, notes.length, avg, getTopCriteria(notes),
      notes.filter(function(n) { return n.type === 'pos'; }).length,
      notes.filter(function(n) { return n.type === 'neg'; }).length, last];
  }).sort(function(a, b) { return b[3] - a[3]; });

  if (rows.length) {
    sh.getRange(5, 1, rows.length, headers.length).setValues(rows).setHorizontalAlignment('center').setFontSize(12);
    rows.forEach(function(row, i) {
      var rr = sh.getRange(5 + i, 1, 1, headers.length);
      rr.setBackground(i % 2 === 0 ? '#1a1008' : '#221608').setFontColor('#e8e0d4');
      rr.setBorder(true, true, true, true, true, true, '#3a2a10', SpreadsheetApp.BorderStyle.SOLID);
      var sc = row[3]; var scC = sh.getRange(5 + i, 4);
      if (sc >= 4.5) { scC.setBackground('#1a3a1a').setFontColor('#80b08a').setFontWeight('bold'); }
      else if (sc >= 3.5) { scC.setBackground('#0a2a3a').setFontColor('#b8d8f0').setFontWeight('bold'); }
      else if (sc >= 2.5) { scC.setBackground('#3a2a0a').setFontColor('#c8a050').setFontWeight('bold'); }
      else if (sc > 0) { scC.setBackground('#3a0a0a').setFontColor('#cc7a60').setFontWeight('bold'); }
    });

    var cr = 5 + rows.length + 2;
    sh.getRange(cr, 1).setValue('العضو');
    sh.getRange(cr, 2).setValue('المتوسط');
    rows.forEach(function(r, i) {
      sh.getRange(cr + 1 + i, 1).setValue(r[0]);
      sh.getRange(cr + 1 + i, 2).setValue(r[3]);
    });
    var chart = sh.newChart().setChartType(Charts.ChartType.BAR)
      .addRange(sh.getRange(cr, 1, rows.length + 1, 2)).setPosition(5 + rows.length + 2, 4, 0, 0)
      .setOption('title', 'مقارنة متوسط التقييم').setOption('titleTextStyle', {color: '#e8b84a', fontSize: 13, bold: true})
      .setOption('backgroundColor', {fill: '#1a1008'}).setOption('legend', {position: 'none'})
      .setOption('hAxis', {minValue: 0, maxValue: 5, textStyle: {color: '#887858'}})
      .setOption('vAxis', {textStyle: {color: '#e8e0d4'}}).setOption('colors', ['#b85840'])
      .setOption('width', 460).setOption('height', 260).build();
    sh.insertChart(chart);
  }

  [140, 100, 100, 120, 180, 80, 80, 110].forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setFrozenRows(4);
}

/* ── MEMBER SHEET ── */

function writeMemberSheet(ss, data, member, role) {
  var shName = '👤 ' + member;
  var sh = ss.getSheetByName(shName) || ss.insertSheet(shName);
  sh.clearContents(); sh.clearFormats();
  sh.getCharts().forEach(function(c) { sh.removeChart(c); });
  sh.setRightToLeft(true);

  var roleColors = {mod: '#b85840', design: '#5a9068', video: '#4878a0'};
  var roleLabels = {mod: 'مودريشن / Community Manager', design: 'Graphic Designer', video: 'Video Editor'};
  var rColor = roleColors[role] || '#c8902a';
  var notes = data.notes.filter(function(n) { return n.member === member; });
  var scored = notes.filter(function(n) { return n.stars > 0; });
  var avg = scored.length ? Math.round(scored.reduce(function(a, n) { return a + n.stars; }, 0) / scored.length * 10) / 10 : 0;
  var pos = notes.filter(function(n) { return n.type === 'pos'; }).length;
  var neg = notes.filter(function(n) { return n.type === 'neg'; }).length;
  var neu = notes.filter(function(n) { return n.type === 'neu'; }).length;

  sh.getRange('A1:J1').merge();
  sh.getRange('A1').setValue('👤 ' + member + '  ·  ' + (roleLabels[role] || role));
  sh.getRange('A1').setBackground('#1a1008').setFontColor(rColor).setFontSize(15).setFontWeight('bold').setHorizontalAlignment('center');
  sh.setRowHeight(1, 42);

  sh.getRange('A2:J2').merge();
  sh.getRange('A2').setValue('إجمالي: ' + notes.length + ' ملاحظة  |  متوسط: ' + (avg || '—') + '/5  |  ' + new Date().toLocaleDateString('ar-EG'));
  sh.getRange('A2').setBackground('#221608').setFontColor('#887858').setFontSize(11).setHorizontalAlignment('center');

  // Stats row
  var sr = 4;
  var stats = [
    {v: avg ? avg + '/5' : '—', l: 'متوسط التقييم', bg: avg >= 4.5 ? '#1a3a1a' : avg >= 3.5 ? '#0a2a3a' : avg > 0 ? '#3a2a0a' : '#221608', tc: avg >= 4.5 ? '#80b08a' : avg >= 3.5 ? '#b8d8f0' : avg > 0 ? '#c8a050' : '#887858'},
    {v: notes.length, l: 'إجمالي الملاحظات', bg: '#2a1a08', tc: '#e8b84a'},
    {v: pos, l: 'إيجابي 👍', bg: '#0a2a10', tc: '#80b08a'},
    {v: neg, l: 'يحتاج تحسين ⚠️', bg: '#2a0a08', tc: '#cc7a60'},
    {v: neu, l: 'ملاحظة 📌', bg: '#08182a', tc: '#b8d8f0'}
  ];
  stats.forEach(function(s, i) {
    var col = 1 + i * 2;
    sh.getRange(sr, col, 1, 2).merge().setValue(s.v).setBackground(s.bg).setFontColor(s.tc).setFontSize(20).setFontWeight('bold').setHorizontalAlignment('center');
    sh.getRange(sr + 1, col, 1, 2).merge().setValue(s.l).setBackground(s.bg).setFontColor('#887858').setFontSize(10).setHorizontalAlignment('center');
  });
  sh.setRowHeight(sr, 42); sh.setRowHeight(sr + 1, 22);

  // Criteria breakdown
  var cr = sr + 3;
  sh.getRange(cr, 1, 1, 5).merge().setValue('📋 تفصيل المعايير');
  sh.getRange(cr, 1).setBackground('#2a1808').setFontColor('#e8b84a').setFontSize(13).setFontWeight('bold').setHorizontalAlignment('right');
  sh.setRowHeight(cr, 32);
  sh.getRange(cr + 1, 1, 1, 5).setValues([['المعيار', 'عدد التقييمات', 'متوسط النجوم', 'إيجابي', 'يحتاج تحسين']]);
  sh.getRange(cr + 1, 1, 1, 5).setBackground('#b85840').setFontColor('#f2e8d8').setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');

  var uCrits = [];
  var seen = {};
  notes.forEach(function(n) { if (n.critLabel && !seen[n.critLabel]) { seen[n.critLabel] = true; uCrits.push(n.critLabel); } });
  var cData = uCrits.map(function(c) {
    var cn = notes.filter(function(n) { return n.critLabel === c; });
    var cs = cn.filter(function(n) { return n.stars > 0; });
    var ca = cs.length ? Math.round(cs.reduce(function(a, n) { return a + n.stars; }, 0) / cs.length * 10) / 10 : 0;
    return [c, cn.length, ca, cn.filter(function(n) { return n.type === 'pos'; }).length, cn.filter(function(n) { return n.type === 'neg'; }).length];
  });

  if (cData.length) {
    sh.getRange(cr + 2, 1, cData.length, 5).setValues(cData).setHorizontalAlignment('center').setFontSize(11);
    cData.forEach(function(row, i) {
      var rr = sh.getRange(cr + 2 + i, 1, 1, 5);
      rr.setBackground(i % 2 === 0 ? '#1a1008' : '#221608').setFontColor('#e8e0d4');
      rr.setBorder(true, true, true, true, true, true, '#3a2a10', SpreadsheetApp.BorderStyle.SOLID);
      var sc = row[2]; var scC = sh.getRange(cr + 2 + i, 3);
      if (sc >= 4.5) { scC.setBackground('#1a3a1a').setFontColor('#80b08a').setFontWeight('bold'); }
      else if (sc >= 3.5) { scC.setBackground('#0a2a3a').setFontColor('#b8d8f0').setFontWeight('bold'); }
      else if (sc >= 2.5) { scC.setBackground('#3a2a0a').setFontColor('#c8a050').setFontWeight('bold'); }
      else if (sc > 0) { scC.setBackground('#3a0a0a').setFontColor('#cc7a60').setFontWeight('bold'); }
    });
  }

  // Notes log
  var nr = cr + cData.length + 4;
  sh.getRange(nr, 1, 1, 7).merge().setValue('🗒️ سجل الملاحظات');
  sh.getRange(nr, 1).setBackground('#2a1808').setFontColor('#e8b84a').setFontSize(13).setFontWeight('bold').setHorizontalAlignment('right');
  sh.setRowHeight(nr, 32);
  sh.getRange(nr + 1, 1, 1, 7).setValues([['التاريخ', 'الفترة', 'المعيار', 'النجوم', 'النوع', 'الملاحظة', 'الشهر']]);
  sh.getRange(nr + 1, 1, 1, 7).setBackground('#b85840').setFontColor('#f2e8d8').setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');

  var pm = {daily: 'يومي', weekly: 'أسبوعي', monthly: 'شهري'};
  var tm = {pos: 'إيجابي 👍', neg: 'تحسين ⚠️', neu: 'ملاحظة 📌'};
  var tc = {pos: '#0a2a10', neg: '#2a0808', neu: '#08182a'};

  if (notes.length) {
    var nR = notes.map(function(n) {
      return [n.date ? new Date(n.date).toLocaleDateString('ar-EG') : '', pm[n.period] || n.period || '',
        n.critLabel || '', n.stars ? '★'.repeat(n.stars) : '—', tm[n.type] || n.type || '', n.text || '', n.monthKey || ''];
    });
    sh.getRange(nr + 2, 1, nR.length, 7).setValues(nR).setHorizontalAlignment('center').setFontSize(11);
    notes.forEach(function(n, i) {
      var rr = sh.getRange(nr + 2 + i, 1, 1, 7);
      rr.setBackground(tc[n.type] || '#1a1008').setFontColor('#e8e0d4');
      rr.setBorder(true, true, true, true, true, true, '#3a2a10', SpreadsheetApp.BorderStyle.SOLID);
    });
    sh.getRange(nr + 2, 6, nR.length, 1).setHorizontalAlignment('right');
  }

  [120, 90, 160, 90, 110, 260, 90].forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setFrozenRows(2);
}

/* ── HELPERS ── */

function reorderSheets(ss, members) {
  try {
    var s = ss.getSheetByName('📊 ملخص الفريق');
    if (s) { ss.setActiveSheet(s); ss.moveActiveSheet(1); }
    members.forEach(function(m, i) {
      var name = typeof m === 'string' ? m : m.name;
      var sh = ss.getSheetByName('👤 ' + name);
      if (sh) { ss.setActiveSheet(sh); ss.moveActiveSheet(i + 2); }
    });
  } catch(e) {}
}

function getTopCriteria(notes) {
  if (!notes.length) return '—';
  var c = {};
  notes.forEach(function(n) { if (n.critLabel) c[n.critLabel] = (c[n.critLabel] || 0) + 1; });
  var entries = Object.keys(c).map(function(k) { return [k, c[k]]; });
  entries.sort(function(a, b) { return b[1] - a[1]; });
  return entries.length ? entries[0][0] : '—';
}
