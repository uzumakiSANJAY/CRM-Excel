/**
 * CRM Sheet Builder
 * IMPORTANT: Make sure your Google Drive is not "Almost out of storage"
 *            before running. Low storage causes script failures.
 *
 * ORDER:
 *   0. cleanTriggers  (run once to clear old triggers)
 *   1. buildCRM       (creates the sheets)
 *   2. addFormatting  (colours + dropdowns)
 *   3. buildCharts    (dashboard)
 *
 * To test if the script works at all, run: testScript
 */

// ─── DIAGNOSTIC ──────────────────────────────────────────────────────────────
function testScript() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var t  = ss.getSheetByName('__test');
  if (t) ss.deleteSheet(t);
  ss.insertSheet('__test').getRange('A1').setValue('OK');
  SpreadsheetApp.getUi().alert('Script is working. You can delete the __test sheet and proceed.');
}

// ─── CLEAR OLD TRIGGERS ───────────────────────────────────────────────────────
function cleanTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });
  SpreadsheetApp.getUi().alert('Old triggers cleared. Now run buildCRM.');
}

// ─── AUTO-FILL via onEdit (no ARRAYFORMULA needed) ───────────────────────────
function onEdit(e) {
  try {
    var sheet = e.range.getSheet();
    if (sheet.getName() !== 'CRM Data') return;
    var row = e.range.getRow();
    var col = e.range.getColumn();
    if (row < 2) return;

    // Col A: Customer selected → fill Company (B) + Phone (C)
    if (col === 1) {
      var val = e.value;
      sheet.getRange(row, 2).clearContent();
      sheet.getRange(row, 3).clearContent();
      if (!val) return;
      var master = e.source.getSheetByName('Master Data');
      if (!master) return;
      var rows = master.getRange('A2:C50').getValues();
      for (var i = 0; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === String(val).trim()) {
          sheet.getRange(row, 2).setValue(rows[i][1]);
          sheet.getRange(row, 3).setValue(rows[i][2]);
          break;
        }
      }
    }

    // Col G: Collected Amount → auto-set Status (D)
    if (col === 7) {
      var amt  = sheet.getRange(row, 6).getValue();
      var coll = Number(e.value) || 0;
      if (!amt) return;
      sheet.getRange(row, 4).setValue(coll >= amt ? 'Paid' : coll > 0 ? 'Partial' : 'Pending');
    }
  } catch (_) {}
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('CRM Tools')
    .addItem('TEST — Check script works', 'testScript')
    .addSeparator()
    .addItem('0. Clean Old Triggers',  'cleanTriggers')
    .addItem('1. Build Sheets',        'buildCRM')
    .addItem('2. Add Formatting',      'addFormatting')
    .addItem('3. Add Charts',          'buildCharts')
    .addSeparator()
    .addItem('Add Customer',           'addNewCustomer')
    .addItem('Add Collector',          'addNewCollector')
    .addToUi();
}

// ─── STEP 1 ───────────────────────────────────────────────────────────────────
function buildCRM() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  ['CRM Data', 'Master Data', 'Charts'].forEach(function(n) {
    var s = ss.getSheetByName(n); if (s) ss.deleteSheet(s);
  });

  // Master Data
  var m = ss.insertSheet('Master Data');
  m.getRange('A1:G1').setValues([['Customer Name','Company Name','Phone Number','','Collectors','','Status']]);
  m.getRange('A2:C4').setValues([
    ['John Smith',  'ABC Corp',        '9876543210'],
    ['Jane Doe',    'XYZ Ltd',         '9123456789'],
    ['Raj Kumar',   'Raj Enterprises', '9000011111']
  ]);
  m.getRange('E2:E5').setValues([['Arun'],['Sunita'],['Vikram'],['Priya']]);
  m.getRange('G2:G6').setValues([['Pending'],['Paid'],['Overdue'],['Partial'],['Cancelled']]);

  // CRM Data — headers only, no formatting yet
  var c = ss.insertSheet('CRM Data');
  c.getRange('A1:N1').setValues([[
    'Customer Name','Company Name','Phone Number','Status',
    'Bill Generate Date','Amount','Collected Amount','Collection Date',
    'Collected By','Pending Amount','Delay Days',
    'Contacted Date','Last Contacted Date','Notes'
  ]]);
  c.setFrozenRows(1);

  // Pending Amount + Delay Days — limited range ARRAYFORMULAs (avoids memory issues)
  c.getRange('J2').setFormula('=ARRAYFORMULA(IF(F2:F101="","",IFERROR(F2:F101-G2:G101,0)))');
  c.getRange('K2').setFormula('=ARRAYFORMULA(IF(E2:E101="","",IF(D2:D101="Paid","",IFERROR(TODAY()-INT(E2:E101),0))))');

  // Charts placeholder
  ss.insertSheet('Charts').getRange('A1').setValue('Run Step 2, then Step 3 from CRM Tools menu.');

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Step 1 done!\nRun: CRM Tools → Step 2 — Add Formatting');
}

// ─── STEP 2 ───────────────────────────────────────────────────────────────────
function addFormatting() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var c  = ss.getSheetByName('CRM Data');
  var m  = ss.getSheetByName('Master Data');
  if (!c || !m) { SpreadsheetApp.getUi().alert('Run Step 1 first.'); return; }

  var ROWS = 100;

  // Header colour
  c.getRange('A1:N1').setFontWeight('bold').setBackground('#1a73e8').setFontColor('white').setFontSize(11);
  m.getRange('A1:G1').setFontWeight('bold').setBackground('#1a73e8').setFontColor('white');

  // Column widths
  [155,155,130,110,145,110,145,130,135,125,100,135,150,200].forEach(function(w,i){
    c.setColumnWidth(i+1, w);
  });

  // Dropdowns — list-based (fast, no range reference)
  var custs = m.getRange('A2:A50').getValues().map(function(r){return r[0];}).filter(Boolean);
  var colls = m.getRange('E2:E20').getValues().map(function(r){return r[0];}).filter(Boolean);
  if (!custs.length) custs = ['Add via Master Data sheet'];
  if (!colls.length) colls = ['Add via Master Data sheet'];

  c.getRange(2,1,ROWS,1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(custs,true).setAllowInvalid(true).build());
  c.getRange(2,4,ROWS,1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Pending','Paid','Overdue','Partial','Cancelled'],true).build());
  c.getRange(2,9,ROWS,1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(colls,true).setAllowInvalid(true).build());

  // Number / date formats
  c.getRange(2,5,ROWS,1).setNumberFormat('dd/mm/yyyy');
  c.getRange(2,8,ROWS,1).setNumberFormat('dd/mm/yyyy');
  c.getRange(2,12,ROWS,2).setNumberFormat('dd/mm/yyyy');
  c.getRange(2,6,ROWS,2).setNumberFormat('#,##0.00');
  c.getRange(2,10,ROWS,1).setNumberFormat('#,##0.00');

  // Conditional formatting
  var dR = c.getRange('D2:D'+(ROWS+1));
  var kR = c.getRange('K2:K'+(ROWS+1));
  c.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Paid')     .setBackground('#b7e1cd').setFontColor('#0f5132').setRanges([dR]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Pending')  .setBackground('#fff3cd').setFontColor('#856404').setRanges([dR]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Overdue')  .setBackground('#f8d7da').setFontColor('#842029').setRanges([dR]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Partial')  .setBackground('#cce5ff').setFontColor('#004085').setRanges([dR]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Cancelled').setBackground('#e2e3e5').setFontColor('#383d41').setRanges([dR]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(30)   .setBackground('#f8d7da').setFontColor('#842029').setRanges([kR]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(1,30)     .setBackground('#fff3cd').setFontColor('#856404').setRanges([kR]).build()
  ]);

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Step 2 done!\nRun: CRM Tools → Step 3 — Add Charts');
}

// ─── STEP 3 ───────────────────────────────────────────────────────────────────
function buildCharts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s  = ss.getSheetByName('Charts');
  if (!s) { SpreadsheetApp.getUi().alert('Run Step 1 first.'); return; }
  s.clear();

  s.getRange('A1').setValue('CRM Dashboard').setFontSize(20).setFontWeight('bold').setFontColor('#1a73e8');

  s.getRange('A3').setValue('By Status').setFontSize(12).setFontWeight('bold');
  s.getRange('A4:E4').setValues([['Status','Count','Billed','Collected','Pending']])
   .setFontWeight('bold').setBackground('#1a73e8').setFontColor('white').setHorizontalAlignment('center');
  ['Pending','Paid','Overdue','Partial','Cancelled'].forEach(function(st,i){
    var r=i+5;
    s.getRange(r,1).setValue(st);
    s.getRange(r,2).setFormula('=COUNTIF(\'CRM Data\'!D:D,"'+st+'")');
    s.getRange(r,3).setFormula('=SUMIF(\'CRM Data\'!D:D,"'+st+'",\'CRM Data\'!F:F)').setNumberFormat('#,##0.00');
    s.getRange(r,4).setFormula('=SUMIF(\'CRM Data\'!D:D,"'+st+'",\'CRM Data\'!G:G)').setNumberFormat('#,##0.00');
    s.getRange(r,5).setFormula('=SUMIF(\'CRM Data\'!D:D,"'+st+'",\'CRM Data\'!J:J)').setNumberFormat('#,##0.00');
  });
  s.getRange('A10:E10').setValues([['TOTAL','=SUM(B5:B9)','=SUM(C5:C9)','=SUM(D5:D9)','=SUM(E5:E9)']])
   .setFontWeight('bold').setBackground('#d0e4f7');

  s.getRange('A12').setValue('Monthly Trend').setFontSize(12).setFontWeight('bold');
  s.getRange('A13:D13').setValues([['Month','Billed','Collected','Pending']])
   .setFontWeight('bold').setBackground('#1a73e8').setFontColor('white').setHorizontalAlignment('center');
  var now=new Date();
  for(var mo=5;mo>=0;mo--){
    var d=new Date(now.getFullYear(),now.getMonth()-mo,1),yr=d.getFullYear(),mn=d.getMonth()+1,r=14+(5-mo);
    s.getRange(r,1).setValue(Utilities.formatDate(d,Session.getScriptTimeZone(),'MMM yyyy'));
    s.getRange(r,2).setFormula('=SUMPRODUCT((MONTH(\'CRM Data\'!E$2:E$101)='+mn+')*(YEAR(\'CRM Data\'!E$2:E$101)='+yr+')*(\'CRM Data\'!F$2:F$101))').setNumberFormat('#,##0.00');
    s.getRange(r,3).setFormula('=SUMPRODUCT((MONTH(\'CRM Data\'!E$2:E$101)='+mn+')*(YEAR(\'CRM Data\'!E$2:E$101)='+yr+')*(\'CRM Data\'!G$2:G$101))').setNumberFormat('#,##0.00');
    s.getRange(r,4).setFormula('=SUMPRODUCT((MONTH(\'CRM Data\'!E$2:E$101)='+mn+')*(YEAR(\'CRM Data\'!E$2:E$101)='+yr+')*(\'CRM Data\'!J$2:J$101))').setNumberFormat('#,##0.00');
  }
  [1,2,3,4,5].forEach(function(col){ s.setColumnWidth(col,140); });
  SpreadsheetApp.flush();

  s.insertChart(s.newChart().setChartType(Charts.ChartType.PIE)
    .addRange(s.getRange('A4:B9')).setPosition(3,7,0,0)
    .setOption('title','Status Distribution').setOption('width',400).setOption('height',280).setOption('is3D',true).build());
  s.insertChart(s.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(s.getRange('A4:A9')).addRange(s.getRange('C4:E9')).setPosition(3,13,0,0)
    .setOption('title','Amount by Status').setOption('width',460).setOption('height',280).setOption('legend',{position:'bottom'}).build());
  s.insertChart(s.newChart().setChartType(Charts.ChartType.LINE)
    .addRange(s.getRange('A13:D19')).setPosition(20,1,0,0)
    .setOption('title','Monthly Trend').setOption('width',650).setOption('height',300).setOption('legend',{position:'bottom'}).setOption('curveType','function').build());

  SpreadsheetApp.getUi().alert('All done! Your CRM is ready.');
}

// ─── ADD CUSTOMER ─────────────────────────────────────────────────────────────
function addNewCustomer() {
  var ui=SpreadsheetApp.getUi(), m=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master Data');
  if(!m){ui.alert('Run Step 1 first.');return;}
  var n=ui.prompt('Customer Name:',ui.ButtonSet.OK_CANCEL);   if(n.getSelectedButton()!==ui.Button.OK||!n.getResponseText().trim())return;
  var co=ui.prompt('Company Name:',ui.ButtonSet.OK_CANCEL);   if(co.getSelectedButton()!==ui.Button.OK)return;
  var p=ui.prompt('Phone Number:',ui.ButtonSet.OK_CANCEL);    if(p.getSelectedButton()!==ui.Button.OK)return;
  m.getRange(m.getLastRow()+1,1,1,3).setValues([[n.getResponseText().trim(),co.getResponseText().trim(),p.getResponseText().trim()]]);
  ui.alert('Customer added! Re-run Step 2 to refresh the dropdown.');
}

// ─── ADD COLLECTOR ────────────────────────────────────────────────────────────
function addNewCollector() {
  var ui=SpreadsheetApp.getUi(), m=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master Data');
  if(!m){ui.alert('Run Step 1 first.');return;}
  var r=ui.prompt('Collector Name:',ui.ButtonSet.OK_CANCEL);  if(r.getSelectedButton()!==ui.Button.OK||!r.getResponseText().trim())return;
  var vals=m.getRange('E:E').getValues(),last=1;
  vals.forEach(function(row,i){if(row[0]!=='')last=i+1;});
  m.getRange(last+1,5).setValue(r.getResponseText().trim());
  ui.alert('Collector added! Re-run Step 2 to refresh the dropdown.');
}
