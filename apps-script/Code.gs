/**
 * KHO BÁO GIÁ CHUNG – Góc Nhà
 * Dán toàn bộ file này vào Google Apps Script (script.google.com) bằng tài khoản Google Workspace của công ty.
 * Báo giá được lưu thành file .json trong thư mục Drive "Báo giá Góc Nhà (app)" của tài khoản đó.
 *
 * ĐỔI MÃ PIN Ở DÒNG DƯỚI trước khi Triển khai. Khi có nhân viên nghỉ: đổi PIN → Lưu (không cần triển khai lại).
 */
const PIN = 'DOI-MA-PIN-NAY';

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

function doGet() { return out_({ ok: true, app: 'kho-bao-gia-goc-nha' }); }

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (x) { return out_({ ok: false, err: 'bad_request' }); }
  if (String(req.pin || '') !== PIN) return out_({ ok: false, err: 'pin' });
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
      return out_({ ok: true, id: f.getId(), updated: f.getLastUpdated().getTime(), data: JSON.parse(f.getBlob().getDataAsString('UTF-8')) });
    }
    if (req.action === 'save') {
      lock.waitLock(20000);
      const d = req.data || {}, kh = d.kh || {};
      const meta = JSON.stringify({ kh: kh.ten || '', sobg: kh.sobg || '', nv: kh.nv || '', ngay: kh.ngay || '', tong: req.tong || 0, by: req.by || '' });
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
      f.setDescription(meta);
      return out_({ ok: true, id: f.getId(), updated: DriveApp.getFileById(f.getId()).getLastUpdated().getTime() });
    }
    if (req.action === 'delete') {
      lock.waitLock(20000);
      const f = DriveApp.getFileById(req.id);
      if (!inFolder_(f, folder)) return out_({ ok: false, err: 'not_found' });
      f.setTrashed(true); // vào Thùng rác Drive, khôi phục được trong 30 ngày
      return out_({ ok: true });
    }
    return out_({ ok: false, err: 'unknown_action' });
  } catch (x) {
    return out_({ ok: false, err: String(x && x.message || x) });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}
