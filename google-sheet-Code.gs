/**
 * منصات المراجعة – مدرسة جيمس وستمنستر – رأس الخيمة
 * يستقبل تقييمات الطلاب من المنصة ويسجلها في ورقة "Responses".
 * Receives student ratings from the revision platform and appends them to the "Responses" sheet.
 */
var SHEET_NAME = 'Responses';
var HEADERS = ['Timestamp', 'Subject', 'Subject code', 'Grade', 'Revision ID', 'Revision title',
  'Student name', 'Class', 'Happiness (1-5)', 'Helped prepare', 'Questions clear', 'Liked explanations',
  'Liked most', 'Comment', 'Score', 'Total', 'Answered', 'Percent', 'Minutes', 'Source'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var d = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS);
      sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#1f3b63').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }
    var clip = function (v, n) { return String(v == null ? '' : v).slice(0, n); };
    sh.appendRow([
      new Date(), clip(d.subjectName, 40), clip(d.subject, 20), clip(d.grade, 6), clip(d.revision, 10), clip(d.title, 80),
      clip(d.name, 60), clip(d.section, 12), Number(d.happy) || '', clip(d.helped, 3), clip(d.clear, 3), clip(d.feedback, 3),
      clip(d.liked, 80), clip(d.comment, 300), Number(d.score) || 0, Number(d.total) || 0, Number(d.answered) || 0,
      Number(d.percent) || 0, d.minutes === '' ? '' : Number(d.minutes) || '', clip(d.source, 60)
    ]);
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// افتح رابط تطبيق الويب في المتصفح للتأكد أنه يعمل
// Open the web-app URL in a browser to check it is live.
function doGet() {
  return ContentService.createTextOutput('Revision platform ratings endpoint is working ✓');
}
