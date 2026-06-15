// ── Google Apps Script สำหรับระบบจองตารางงานเภสัชกร ──
// วิธีใช้: Extensions → Apps Script → วางโค้ดนี้ → Deploy → Web App

var SHEET_NAME = 'Bookings';

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['Date', 'Name', 'Slot', 'Hours', 'Timestamp']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function doGet(e) {
  var p = e.parameter;
  var cb = p.cb || 'callback';
  var result;
  try {
    var action = p.action || 'get';
    if (action === 'get') {
      result = getBookings(p.month, p.year);
    } else if (action === 'save') {
      result = saveBooking(p.name, p.date, p.slot, parseFloat(p.hours));
    } else if (action === 'delete') {
      result = deleteBooking(p.name, p.date);
    } else {
      result = { error: 'unknown action' };
    }
  } catch (err) {
    result = { error: err.toString() };
  }
  return ContentService
    .createTextOutput(cb + '(' + JSON.stringify(result) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function getBookings(month, year) {
  var sh = getSheet();
  var data = sh.getDataRange().getValues();
  var result = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0] || !row[1]) continue;
    var ds = Utilities.formatDate(new Date(row[0]), 'Asia/Bangkok', 'yyyy-MM-dd');
    if (month && year) {
      var ym = year + '-' + (parseInt(month) < 10 ? '0' + parseInt(month) : '' + parseInt(month));
      if (!ds.startsWith(ym)) continue;
    }
    if (!result[ds]) result[ds] = [];
    result[ds].push({ name: row[1] + '', slot: row[2] + '', hours: parseFloat(row[3]) || 0 });
  }
  return result;
}

function saveBooking(name, date, slot, hours) {
  deleteBooking(name, date);
  var sh = getSheet();
  sh.appendRow([new Date(date), name, slot, hours, new Date()]);
  return { success: true };
}

function deleteBooking(name, date) {
  var sh = getSheet();
  var data = sh.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    var ds = Utilities.formatDate(new Date(data[i][0]), 'Asia/Bangkok', 'yyyy-MM-dd');
    if (ds === date && data[i][1] + '' === name) {
      sh.deleteRow(i + 1);
    }
  }
  return { success: true };
}
