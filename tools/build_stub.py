# Tạo apps-script/DAN-1-LAN.gs = mã đầy đủ (dự phòng) + tự lấy bản mới hơn từ GitHub nếu tải được
import re
src=open('/home/claude/baogiaapp/apps-script/Code.gs',encoding='utf-8').read()
local=re.sub(r'\bfunction doGet\(',"function doGet_banDan_(",src)
local=re.sub(r'\bfunction doPost\(',"function doPost_banDan_(",local)
assert 'doGet_banDan_' in local and 'doPost_banDan_' in local
head=r"""// ============================================================
//  KHO BÁO GIÁ GÓC NHÀ – BẢN DÁN 1 LẦN
//  Tự lấy mã mới nhất từ GitHub; nếu không tải được thì dùng mã có sẵn bên dưới.
//  => Sau này cập nhật app KHÔNG cần dán lại / triển khai lại.
// ============================================================
var NGUON_MA_ = [
  'https://gocnhakcw.github.io/baogiaapp/apps-script/Code.gs',
  'https://raw.githubusercontent.com/GocNhaKCW/baogiaapp/main/apps-script/Code.gs',
  'https://cdn.jsdelivr.net/gh/GocNhaKCW/baogiaapp@main/apps-script/Code.gs'
];
var LOI_TAI_ = [];
function maMoi_() {
  var cache = CacheService.getScriptCache(), ma = cache.get('ma');
  if (ma) return ma === 'X' ? null : ma;
  for (var i = 0; i < NGUON_MA_.length; i++) {
    try {
      var r = UrlFetchApp.fetch(NGUON_MA_[i] + '?t=' + Date.now(), { muteHttpExceptions: true, followRedirects: true });
      var t = r.getContentText();
      if (r.getResponseCode() === 200 && t.indexOf('function doPost') >= 0) { try { cache.put('ma', t, 300); } catch (e) {} return t; }
      LOI_TAI_.push(NGUON_MA_[i].split('/')[2] + ': HTTP ' + r.getResponseCode());
    } catch (e) { LOI_TAI_.push(NGUON_MA_[i].split('/')[2] + ': ' + (e && e.message || e)); }
  }
  try { cache.put('ma', 'X', 300); cache.put('loi', LOI_TAI_.join(' | '), 300); } catch (e) {}
  return null;
}
function chon_() {
  var ma = null; try { ma = maMoi_(); } catch (e) { LOI_TAI_.push(String(e && e.message || e)); }
  if (ma) { try { var m = (new Function(ma + '\n;return { doGet: doGet, doPost: doPost };'))(); m.nguon = 'github'; return m; } catch (e) { LOI_TAI_.push('mã GitHub lỗi: ' + e.message); } }
  return { doGet: doGet_banDan_, doPost: doPost_banDan_, nguon: 'ban_dan' };
}
function doGet(e) {
  var m = chon_(), out = m.doGet(e);
  try { var o = JSON.parse(out.getContent()); o.nguon_ma = m.nguon; if (m.nguon !== 'github') o.loi_tai_github = LOI_TAI_.join(' | ') || (CacheService.getScriptCache().get('loi') || ''); return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); } catch (x) { return out; }
}
function doPost(e) { return chon_().doPost(e); }

// ===================== MÃ CÓ SẴN (dự phòng) =====================
"""
open('/home/claude/baogiaapp/apps-script/DAN-1-LAN.gs','w',encoding='utf-8').write(head+local)
print('built', len(head+local))
