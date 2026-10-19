// ============================================================
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
/**
 * KHO BÁO GIÁ & HỢP ĐỒNG CHUNG – Góc Nhà
 * Dán toàn bộ file này vào Google Apps Script (script.google.com) bằng tài khoản Google Workspace của công ty.
 *
 * Cài đặt trong: Cài đặt dự án (biểu tượng bánh răng) → Thuộc tính tập lệnh (Script properties):
 *   PIN        = mã PIN chung cho nhân viên (bắt buộc). Có người nghỉ → sửa PIN → Lưu. Không cần triển khai lại.
 *   FOLDER_ID  = ID thư mục Drive chung (không bắt buộc). Bỏ trống thì script tự tạo thư mục "Báo giá Góc Nhà (app)".
 *
 * Trong thư mục chung:
 *   - mỗi báo giá = 1 file .json (để app mở lại & sửa tiếp)
 *   - mỗi báo giá có 1 thư mục con "<Khách hàng> – <Số BG>" chứa: Báo giá.pdf, Hợp đồng.docx,
 *     Đề nghị tạm ứng.docx, Đề nghị thanh toán - Đợt N.docx, Biên bản giao nhận.docx
 *     → mỗi lần CSKH lưu lại báo giá, các file này được thay bằng bản mới nhất.
 */
function pin_() { return PropertiesService.getScriptProperties().getProperty('PIN') || ''; }

const FOLDER_NAME = 'Báo giá Góc Nhà (app)';

function folder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FOLDER_ID');
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  const f = DriveApp.createFolder(FOLDER_NAME);
  props.setProperty('FOLDER_ID', f.getId());
  return f;
}
function out_(o) { o.v = SCRIPT_VERSION; return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function inFolder_(file, folder) { const ps = file.getParents(); while (ps.hasNext()) { if (ps.next().getId() === folder.getId()) return true; } return false; }
function meta_(file) { try { return JSON.parse(file.getDescription() || '{}'); } catch (e) { return {}; } }
function subName_(kh, sobg) { return String((kh || 'Khách hàng') + (sobg ? ' – ' + sobg : '')).replace(/[\\\/:*?"<>|]/g, '-').slice(0, 120); }

// Thư mục con của 1 báo giá: luôn dùng lại đúng 1 thư mục.
// Thứ tự tìm: id app gửi lên → id lưu trong Script Properties (đọc tức thì, không bị trễ như mô tả file)
//             → id trong mô tả file → thư mục cùng tên → mới tạo.
function sub_(file, folder, kh, sobg, hint) {
  const props = PropertiesService.getScriptProperties(), key = 'SUB_' + file.getId();
  const tryId = function (id) { if (!id) return null; try { const f = DriveApp.getFolderById(id); return (!f.isTrashed() && inFolder_(f, folder)) ? f : null; } catch (e) { return null; } };
  let sub = tryId(hint) || tryId(props.getProperty(key)) || tryId(meta_(file).sub);
  if (!sub && kh !== undefined) { const it = folder.getFoldersByName(subName_(kh, sobg)); while (it.hasNext()) { const f = it.next(); if (!f.isTrashed()) { sub = f; break; } } }
  if (!sub) {
    if (kh === undefined) { const m = meta_(file); kh = m.kh || ''; sobg = m.sobg || ''; }
    const it = folder.getFoldersByName(subName_(kh, sobg));
    sub = it.hasNext() ? it.next() : folder.createFolder(subName_(kh, sobg));
  }
  if (kh !== undefined && sub.getName() !== subName_(kh, sobg)) sub.setName(subName_(kh, sobg));
  props.setProperty(key, sub.getId());
  return sub;
}

const SCRIPT_VERSION = '2026-10-19b';
// ===== Thư mục MẪU (app đọc mẫu Word từ đây; sửa mẫu ngay trên Drive) =====
const TPL_FOLDER_NAME = '_MẪU HỢP ĐỒNG & CHỨNG TỪ (app dùng)';
const TPL_README = 'ĐỌC TRƯỚC KHI SỬA MẪU.txt';
const TPL_README_TEXT = [
  'Các file "Mẫu ….docx" trong thư mục này là mẫu app báo giá dùng để tạo Hợp đồng, Đề nghị tạm ứng, Đề nghị thanh toán, Biên bản giao nhận.',
  'Sửa mẫu: mở file trên Drive (Google Tài liệu mở được file .docx) hoặc tải về sửa bằng Word rồi tải lên thay đúng TÊN FILE cũ. Lần xuất sau app tự dùng mẫu mới.',
  'KHÔNG đổi tên file. KHÔNG bấm "Lưu dưới dạng Google Tài liệu" (app chỉ đọc file .docx).',
  'GIỮ NGUYÊN các chữ trong ngoặc nhọn { } – đó là chỗ app tự điền:',
  '  {kh} tên khách · {kh_upper} tên khách IN HOA · {dc} địa chỉ · {mst} MST · {sdt} điện thoại · {em} email',
  '  {dd} người đại diện · {cv} chức vụ · {uq} số giấy uỷ quyền · {uqd} ngày uỷ quyền · {stk} số TK · {nh} ngân hàng',
  '  {dcg} địa chỉ giao hàng · {nguoinhan} người nhận · {sdtnhan} SĐT người nhận · {sd1}-{sd2} / {dh1}-{dh2} số ngày giao',
  '  {tong} tổng tiền · {tong_chu} tổng bằng chữ · {sau} tiền trước thuế · {vat} tiền thuế · {vat_pct} % VAT · {ky_a} tên người ký bên A',
  '  {#dot} … {/dot}: lặp từng đợt thanh toán, bên trong dùng {n} (số lần) và {mota} (nội dung đợt)',
  '  {#items} … {/items}: lặp từng sản phẩm trong bảng, dùng {stt} {ten} {kt} {mau} {dvt} {sl} {dg} {tt} {gc}',
  '  Đề nghị tạm ứng / thanh toán: {pct} %, {tien} số tiền, {chu} bằng chữ, {mota} nội dung đợt, {noidung} nội dung chuyển khoản',
  'Nếu mẫu bị lỗi (thiếu/sai ngoặc), app sẽ báo và tạm dùng mẫu dự phòng. Xoá file mẫu → lần xuất sau app tự cài lại mẫu gốc.'
].join('\n');
function tplFolder_() {
  const props = PropertiesService.getScriptProperties(), root = folder_();
  let f = null; const id = props.getProperty('TPL_FOLDER_ID');
  if (id) { try { f = DriveApp.getFolderById(id); if (f.isTrashed()) f = null; } catch (e) { f = null; } }
  if (!f) { const it = root.getFoldersByName(TPL_FOLDER_NAME); f = it.hasNext() ? it.next() : root.createFolder(TPL_FOLDER_NAME); props.setProperty('TPL_FOLDER_ID', f.getId()); }
  if (!f.getFilesByName(TPL_README).hasNext()) f.createFile(TPL_README, TPL_README_TEXT, MimeType.PLAIN_TEXT);
  return f;
}
function tplFile_(name) { const it = tplFolder_().getFilesByName(name); while (it.hasNext()) { const x = it.next(); if (!x.isTrashed()) return x; } return null; }
// Nguồn mẫu: nếu đã dán link file mẫu riêng (Script Property TPLSRC_<tên mẫu>) thì dùng file đó, không thì dùng file trong thư mục MẪU
function tplKey_(name) { return 'TPLSRC_' + Utilities.base64EncodeWebSafe(Utilities.newBlob(name).getBytes()); }
function tplResolve_(name) {
  const id = PropertiesService.getScriptProperties().getProperty(tplKey_(name));
  if (id) { try { const f = DriveApp.getFileById(id); if (!f.isTrashed()) return { file: f, src: 'link' }; } catch (e) { return { file: null, src: 'link', err: 'Không mở được file mẫu đã dán link (kiểm tra quyền chia sẻ cho tài khoản công ty)' }; } }
  const f = tplFile_(name); return { file: f, src: 'folder' };
}
function idFromLink_(s) { s = String(s || '').trim(); const m = s.match(/\/d\/([A-Za-z0-9_-]{20,})/) || s.match(/[?&]id=([A-Za-z0-9_-]{20,})/) || s.match(/^([A-Za-z0-9_-]{20,})$/); return m ? m[1] : ''; }
// Google Tài liệu → xuất ra .docx; file .docx → lấy nguyên bản
function tplBytes_(f) {
  if (f.getMimeType() === MimeType.GOOGLE_DOCS) {
    const r = UrlFetchApp.fetch('https://docs.google.com/document/d/' + f.getId() + '/export?format=docx', { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) throw new Error('Không xuất được Google Tài liệu sang Word (' + r.getResponseCode() + ')');
    return r.getContent();
  }
  return f.getBlob().getBytes();
}

// Mở link /exec trên trình duyệt để kiểm tra nhanh cấu hình
function doGet_banDan_() {
  const r = { ok: true, app: 'kho-bao-gia-goc-nha', version: SCRIPT_VERSION, pin_da_dat: !!pin_() };
  try { const f = folder_(); r.thu_muc = f.getName(); } catch (e) { r.thu_muc = 'LỖI: ' + e.message; }
  return out_(r);
}

function doPost_banDan_(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (x) { return out_({ ok: false, err: 'bad_request' }); }
  const PIN = pin_();
  if (!PIN || String(req.pin || '') !== PIN) return out_({ ok: false, err: 'pin' });
  const folder = folder_();
  const lock = LockService.getScriptLock();
  try {
    if (req.action === 'list') {
      const it = folder.getFilesByType('application/json'), list = [];
      while (it.hasNext()) {
        const f = it.next();
        list.push({ id: f.getId(), name: f.getName(), updated: f.getLastUpdated().getTime(), meta: meta_(f) });
      }
      list.sort(function (a, b) { return b.updated - a.updated; });
      return out_({ ok: true, list: list });
    }
    if (req.action === 'get') {
      const f = DriveApp.getFileById(req.id);
      if (!inFolder_(f, folder) || f.isTrashed()) return out_({ ok: false, err: 'not_found' });
      return out_({ ok: true, id: f.getId(), updated: f.getLastUpdated().getTime(), sub: meta_(f).sub || '', data: JSON.parse(f.getBlob().getDataAsString('UTF-8')) });
    }
    if (req.action === 'save') {
      lock.waitLock(20000);
      const d = req.data || {}, kh = d.kh || {};
      const content = JSON.stringify(d);
      let f;
      if (req.id) {
        f = DriveApp.getFileById(req.id);
        if (!inFolder_(f, folder) || f.isTrashed()) return out_({ ok: false, err: 'not_found' });
        const cur = f.getLastUpdated().getTime();
        if (!req.force && req.base && cur > req.base + 1500) {
          return out_({ ok: false, err: 'conflict', updated: cur, by: meta_(f).by || '' });
        }
        f.setContent(content);
        f.setName((d.name || 'Bao gia') + '.json');
      } else {
        f = folder.createFile((d.name || 'Bao gia') + '.json', content, 'application/json');
      }
      const sub = sub_(f, folder, kh.ten || '', kh.sobg || '', req.sub);
      f.setDescription(JSON.stringify({ kh: kh.ten || '', sobg: kh.sobg || '', nv: kh.nv || '', ngay: kh.ngay || '', tong: req.tong || 0, by: req.by || '', sub: sub.getId() }));
      return out_({ ok: true, id: f.getId(), sub: sub.getId(), updated: DriveApp.getFileById(f.getId()).getLastUpdated().getTime() });
    }
    if (req.action === 'putfile') {   // lưu/thay 1 file (PDF, Word) vào thư mục con của báo giá
      lock.waitLock(20000);
      const f = DriveApp.getFileById(req.id);
      if (!inFolder_(f, folder) || f.isTrashed()) return out_({ ok: false, err: 'not_found' });
      const sub = sub_(f, folder, undefined, undefined, req.sub);
      const name = String(req.name || 'file').replace(/[\\\/:*?"<>|]/g, '-');
      const old = sub.getFilesByName(name);
      while (old.hasNext()) old.next().setTrashed(true);   // bản cũ vào Thùng rác (khôi phục được 30 ngày)
      const blob = Utilities.newBlob(Utilities.base64Decode(req.b64 || ''), req.mime || 'application/octet-stream', name);
      const nf = sub.createFile(blob);
      return out_({ ok: true, fileId: nf.getId(), url: nf.getUrl(), sub: sub.getId() });
    }
    if (req.action === 'tplinfo') {   // nguồn & ngày sửa của từng mẫu
      const f = tplFolder_(), list = {};
      (req.names || []).forEach(function (n) {
        const r = tplResolve_(n);
        if (r.file) list[n] = { id: r.file.getId(), updated: r.file.getLastUpdated().getTime(), mime: r.file.getMimeType(), src: r.src, url: r.file.getUrl(), title: r.file.getName() };
        else if (r.err) list[n] = { src: 'link', err: r.err };
      });
      return out_({ ok: true, folder: f.getId(), files: list });
    }
    if (req.action === 'tplget') {
      const r = tplResolve_(req.name); if (!r.file) return out_({ ok: false, err: r.err || 'not_found' });
      return out_({ ok: true, updated: r.file.getLastUpdated().getTime(), b64: Utilities.base64Encode(tplBytes_(r.file)) });
    }
    if (req.action === 'tplset') {    // dán link file mẫu riêng (để trống = dùng mẫu trong thư mục MẪU)
      const props = PropertiesService.getScriptProperties(), key = tplKey_(req.name), id = idFromLink_(req.link);
      if (!String(req.link || '').trim()) { props.deleteProperty(key); return out_({ ok: true, src: 'folder' }); }
      if (!id) return out_({ ok: false, err: 'Link không hợp lệ' });
      let f; try { f = DriveApp.getFileById(id); } catch (e) { return out_({ ok: false, err: 'Tài khoản công ty chưa mở được file này – hãy chia sẻ file cho tài khoản công ty (hoặc "Bất kỳ ai có đường liên kết")' }); }
      const mt = f.getMimeType();
      if (mt !== MimeType.GOOGLE_DOCS && mt !== MimeType.MICROSOFT_WORD && !/wordprocessingml/.test(mt)) return out_({ ok: false, err: 'File phải là Word (.docx) hoặc Google Tài liệu' });
      props.setProperty(key, id);
      return out_({ ok: true, src: 'link', title: f.getName() });
    }
    if (req.action === 'tplput') {    // cài mẫu gốc lần đầu (không ghi đè mẫu đã có)
      lock.waitLock(20000);
      if (tplFile_(req.name)) return out_({ ok: true, existed: true });
      tplFolder_().createFile(Utilities.newBlob(Utilities.base64Decode(req.b64 || ''), 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', req.name));
      return out_({ ok: true });
    }
    if (req.action === 'delete') {
      lock.waitLock(20000);
      const f = DriveApp.getFileById(req.id);
      if (!inFolder_(f, folder)) return out_({ ok: false, err: 'not_found' });
      const m = meta_(f);
      f.setTrashed(true); // vào Thùng rác Drive, khôi phục được trong 30 ngày
      const sid = PropertiesService.getScriptProperties().getProperty('SUB_' + f.getId()) || m.sub;
      if (sid) { try { DriveApp.getFolderById(sid).setTrashed(true); } catch (e) {} }
      PropertiesService.getScriptProperties().deleteProperty('SUB_' + f.getId());
      return out_({ ok: true });
    }
    return out_({ ok: false, err: 'unknown_action' });
  } catch (x) {
    return out_({ ok: false, err: String(x && x.message || x) });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}
