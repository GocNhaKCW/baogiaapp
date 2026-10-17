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
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
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

const SCRIPT_VERSION = '2026-10-17l';
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

// Mở link /exec trên trình duyệt để kiểm tra nhanh cấu hình
function doGet() {
  const r = { ok: true, app: 'kho-bao-gia-goc-nha', version: SCRIPT_VERSION, pin_da_dat: !!pin_() };
  try { const f = folder_(); r.thu_muc = f.getName(); } catch (e) { r.thu_muc = 'LỖI: ' + e.message; }
  return out_(r);
}

function doPost(e) {
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
    if (req.action === 'tplinfo') {   // danh sách mẫu hiện có (tên → id, ngày sửa)
      const f = tplFolder_(), it = f.getFiles(), list = {};
      while (it.hasNext()) { const x = it.next(); if (!x.isTrashed()) list[x.getName()] = { id: x.getId(), updated: x.getLastUpdated().getTime(), mime: x.getMimeType() }; }
      return out_({ ok: true, folder: f.getId(), files: list });
    }
    if (req.action === 'tplget') {
      const x = tplFile_(req.name); if (!x) return out_({ ok: false, err: 'not_found' });
      return out_({ ok: true, updated: x.getLastUpdated().getTime(), mime: x.getMimeType(), b64: Utilities.base64Encode(x.getBlob().getBytes()) });
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
