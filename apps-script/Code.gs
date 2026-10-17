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

// Thư mục con của 1 báo giá (tạo nếu chưa có, đổi tên theo khách/số BG hiện tại)
function sub_(file, folder, kh, sobg) {
  const m = meta_(file);
  let sub = null;
  if (m.sub) { try { sub = DriveApp.getFolderById(m.sub); if (sub.isTrashed()) sub = null; } catch (e) { sub = null; } }
  if (!sub) sub = folder.createFolder(subName_(kh, sobg));
  else if (kh !== undefined && sub.getName() !== subName_(kh, sobg)) sub.setName(subName_(kh, sobg));
  return sub;
}

function doGet() { return out_({ ok: true, app: 'kho-bao-gia-goc-nha' }); }

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
      const sub = sub_(f, folder, kh.ten || '', kh.sobg || '');
      f.setDescription(JSON.stringify({ kh: kh.ten || '', sobg: kh.sobg || '', nv: kh.nv || '', ngay: kh.ngay || '', tong: req.tong || 0, by: req.by || '', sub: sub.getId() }));
      return out_({ ok: true, id: f.getId(), sub: sub.getId(), updated: DriveApp.getFileById(f.getId()).getLastUpdated().getTime() });
    }
    if (req.action === 'putfile') {   // lưu/thay 1 file (PDF, Word) vào thư mục con của báo giá
      lock.waitLock(20000);
      const f = DriveApp.getFileById(req.id);
      if (!inFolder_(f, folder) || f.isTrashed()) return out_({ ok: false, err: 'not_found' });
      const sub = sub_(f, folder);
      const name = String(req.name || 'file').replace(/[\\\/:*?"<>|]/g, '-');
      const old = sub.getFilesByName(name);
      while (old.hasNext()) old.next().setTrashed(true);   // bản cũ vào Thùng rác (khôi phục được 30 ngày)
      const blob = Utilities.newBlob(Utilities.base64Decode(req.b64 || ''), req.mime || 'application/octet-stream', name);
      const nf = sub.createFile(blob);
      return out_({ ok: true, fileId: nf.getId(), url: nf.getUrl(), sub: sub.getId() });
    }
    if (req.action === 'delete') {
      lock.waitLock(20000);
      const f = DriveApp.getFileById(req.id);
      if (!inFolder_(f, folder)) return out_({ ok: false, err: 'not_found' });
      const m = meta_(f);
      f.setTrashed(true); // vào Thùng rác Drive, khôi phục được trong 30 ngày
      if (m.sub) { try { DriveApp.getFolderById(m.sub).setTrashed(true); } catch (e) {} }
      return out_({ ok: true });
    }
    return out_({ ok: false, err: 'unknown_action' });
  } catch (x) {
    return out_({ ok: false, err: String(x && x.message || x) });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}
