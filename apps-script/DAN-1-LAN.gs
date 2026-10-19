// ============================================================
//  KHO BÁO GIÁ GÓC NHÀ – BẢN DÁN 1 LẦN, KHÔNG BAO GIỜ PHẢI SỬA LẠI
//  Mã chính tự tải từ GitHub (apps-script/Code.gs). Khi app cần
//  cập nhật, chỉ cần sửa trên GitHub – không phải triển khai lại.
// ============================================================
var NGUON_MA = 'https://raw.githubusercontent.com/GocNhaKCW/baogiaapp/main/apps-script/Code.gs';

function taiMa_() {
  var cache = CacheService.getScriptCache();
  var ma = cache.get('ma');
  if (ma) return ma;
  try {
    var r = UrlFetchApp.fetch(NGUON_MA + '?t=' + Date.now(), { muteHttpExceptions: true });
    if (r.getResponseCode() === 200 && r.getContentText().indexOf('function doPost') >= 0) {
      ma = r.getContentText();
      cache.put('ma', ma, 300);        // dùng lại trong 5 phút
      cache.put('ma_du_phong', ma, 21600);
      return ma;
    }
  } catch (e) {}
  ma = cache.get('ma_du_phong');       // GitHub lỗi tạm thời → dùng bản gần nhất
  if (ma) return ma;
  throw new Error('Không tải được mã từ GitHub');
}
function chay_() { return (new Function(taiMa_() + '\n;return { doGet: doGet, doPost: doPost };'))(); }

function doGet(e) {
  try { return chay_().doGet(e); }
  catch (x) { return ContentService.createTextOutput(JSON.stringify({ ok: false, err: String(x && x.message || x) })).setMimeType(ContentService.MimeType.JSON); }
}
function doPost(e) {
  try { return chay_().doPost(e); }
  catch (x) { return ContentService.createTextOutput(JSON.stringify({ ok: false, err: String(x && x.message || x) })).setMimeType(ContentService.MimeType.JSON); }
}

// Không chạy – chỉ để Google xin đủ quyền (Drive, tải mẫu, khoá, bộ nhớ) khi cấp quyền lần đầu
function capQuyen_() {
  DriveApp.getRootFolder(); DriveApp.getFileById('x'); DriveApp.getFolderById('x'); DriveApp.createFolder('x');
  PropertiesService.getScriptProperties(); LockService.getScriptLock(); CacheService.getScriptCache();
  UrlFetchApp.fetch('https://github.com'); ScriptApp.getOAuthToken(); Utilities.newBlob('');
}
