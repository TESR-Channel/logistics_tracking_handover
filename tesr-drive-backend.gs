/**
 * TESR — Tracking & Handover backend (Google Apps Script)  v2.5 — root โฟลเดอร์ล็อกตายตัว (ไม่รับค่าจาก client), chunked upload, no 50MB limit
 * ---------------------------------------------------------------------------
 * โครงสร้างโฟลเดอร์:  TESR shop tracking and handover / <ปี> / <Tax ID> / <Invoice>_<DDMMYYYY>
 *
 * v2 เปลี่ยนจากส่งทุกไฟล์ใน request เดียว (ติด limit 50MB ของ Apps Script → "Load failed" บนมือถือ)
 * เป็นส่งทีละไฟล์ และไฟล์ใหญ่ (วิดีโอ) ส่งเป็นชิ้นละ 4MB ผ่าน Drive resumable upload
 *
 * actions ที่หน้าเว็บเรียก (POST, JSON):
 *   create  {record}                                  → สร้างโฟลเดอร์ root/ปี/TaxID/Invoice_วันที่ + README + Public → {ok, folderId, folderUrl}
 *   upload  {folderId, name, mime, data(base64)}      → ไฟล์เล็ก (≤ 15MB) อัปโหลดตรง             → {ok, fileId}
 *   begin   {folderId, name, mime, size}              → เปิด resumable session สำหรับไฟล์ใหญ่     → {ok, sessionUri}
 *   chunk   {sessionUri, start, end, total, data}     → ส่งชิ้นถัดไป (ขนาดต้องหาร 256KB ลงตัว ยกเว้นชิ้นสุดท้าย) → {ok, done, fileId?}
 *   finish  {folderId, record, fileCount}             → เขียน log ลง Google Sheet                → {ok}
 *
 * วิธี Deploy / อัปเดต
 *   1. วางโค้ดนี้ทับของเดิมทั้งหมด → Save
 *   2. **สำคัญ**: กด Run ▶ ที่ฟังก์ชัน `authorize` หนึ่งครั้ง แล้วกด Allow — v2 ต้องขอสิทธิ์เพิ่ม (เรียก Drive API ผ่าน UrlFetch)
 *   3. Deploy → Manage deployments → ✎ → Version: New version → Deploy   (URL เดิม ไม่ต้องเปลี่ยนในหน้าเว็บ)
 */

const CONFIG = {
  ROOT_FOLDER_ID: '1PuUBsXmCo0mnQDcdvDqbDwtH_Qiv_Hni', // "TESR shop tracking and handover" (ยืนยัน 10/10/2026)
  LOG_SHEET_NAME: 'TESR Handover Log',
  MAKE_PUBLIC:    true,
};

/** กด Run ที่ฟังก์ชันนี้หนึ่งครั้งหลังวางโค้ด เพื่อให้ Google ขอสิทธิ์ Drive + UrlFetch */
function authorize() {
  DriveApp.getRootFolder().getName();
  SpreadsheetApp.getActiveSpreadsheet;
  UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/about?fields=user', { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  Logger.log('authorized OK');
}

function doGet() {
  return json_({ ok: true, service: 'TESR Handover backend v2.5', time: new Date().toISOString(), root: (function(){ try { return DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID).getName(); } catch (e) { return 'ERROR: เปิด root ไม่ได้'; } })() });
}

function doPost(e) {
  try {
    const b = JSON.parse(e.postData.contents);
    switch (b.action) {
      case 'create': return json_(create_(b));
      case 'upload': return json_(upload_(b));
      case 'begin':  return json_(begin_(b));
      case 'chunk':  return json_(chunk_(b));
      case 'finish': return json_(finish_(b));
      default: throw new Error('unknown action "' + b.action + '" — หน้าเว็บอาจเป็นเวอร์ชันเก่า');
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/* ── actions ─────────────────────────────────────────────── */
function create_(b) {
  const r = b.record || {};
  if (!r.id || !r.taxId) throw new Error('record.id / record.taxId is required');
  // v2.5: root ล็อกตายตัวที่ CONFIG.ROOT_FOLDER_ID เท่านั้น — ไม่รับค่าจากหน้าเว็บอีก
  // (เคยมีเครื่องที่ Settings ค้างค่า root เก่า ทำให้ record ไปลงผิดที่)
  var root;
  try { root = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID); root.getName(); }
  catch (e) { throw new Error('เปิดโฟลเดอร์ root (' + CONFIG.ROOT_FOLDER_ID + ') ไม่ได้ — ตรวจ CONFIG.ROOT_FOLDER_ID ใน Apps Script'); }
  const year = String(r.invDate || r.createdAt || new Date().toISOString()).slice(0, 4);
  const customerFolder = getOrCreateFolder_(getOrCreateFolder_(root, year), safeName_(r.taxId));
  const base = safeName_(r.folderName || ((r.invoice || 'INV') + '_' + ddmmyyyy_(r.invDate)));
  let name = base, n = 1;
  while (customerFolder.getFoldersByName(name).hasNext()) name = base + '-' + (++n);
  const folder = customerFolder.createFolder(name);

  const readme = [
    'TESR Co., Ltd. — Handover Record', '=================================',
    'Record ID       : ' + r.id,
    'Customer        : ' + (r.company || '-') + ' (Tax ID ' + r.taxId + ')',
    'Contact         : ' + (r.contact || '-') + ' ' + (r.phone || ''),
    'Invoice/Receipt : ' + (r.invoice || '-') + '  (' + (r.invDateDisplay || r.invDate || '-') + ')',
    'Tracking Link   : ' + (r.trackLink || '-'),
    'Drive path      : ' + year + '/' + r.taxId + '/' + name,
    'Created         : ' + (r.createdAt || new Date().toISOString()) + ' by ' + (r.createdBy || '-'),
    '', 'Files:'].concat((r.files || []).map(function (f) { return ' - ' + f.name; }))
    .concat(['', 'Notes:', r.notes || '-', '', 'Robot · AI · IoT · Automation — Real-world Solutions']).join('\n');
  folder.createFile('README.txt', readme, MimeType.PLAIN_TEXT);
  if (CONFIG.MAKE_PUBLIC) folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { ok: true, folderId: folder.getId(), folderUrl: folder.getUrl(), folderName: name, drivePath: year + '/' + r.taxId + '/' + name };
}

function upload_(b) {
  const folder = DriveApp.getFolderById(b.folderId);
  const bytes = decode_(b.data, 'upload ' + b.name);
  if (b.size && bytes.length !== b.size) throw new Error('upload ' + b.name + ': received ' + bytes.length + ' of ' + b.size + ' bytes — ส่งซ้ำอีกครั้ง');
  // ถ้าชื่อซ้ำ (retry) ให้ลบของเก่าก่อน
  const old = folder.getFilesByName(b.name); while (old.hasNext()) old.next().setTrashed(true);
  const f = folder.createFile(Utilities.newBlob(bytes, b.mime || 'application/octet-stream', b.name));
  return { ok: true, fileId: f.getId() };
}

/** base64 → bytes พร้อมข้อความ error ที่บอกได้ว่าอะไรผิด (แทน "Could not decode string.") */
function decode_(data, ctx) {
  if (typeof data !== 'string' || !data.length) throw new Error(ctx + ': ไม่มีข้อมูลไฟล์ในคำขอ (data ว่าง) — หน้าเว็บอ่านไฟล์ไม่สำเร็จ ลองลบไฟล์แล้วเพิ่มใหม่');
  const clean = data.replace(/^data:[^,]*,/, '').replace(/\s/g, '');
  try { return Utilities.base64Decode(clean); }
  catch (e) { throw new Error(ctx + ': base64 เสียหาย (' + clean.length + ' chars) — ลองส่งอีกครั้ง'); }
}

function begin_(b) {
  const res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&supportsAllDrives=true', {
    method: 'post',
    contentType: 'application/json; charset=UTF-8',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken(), 'X-Upload-Content-Type': b.mime || 'application/octet-stream', 'X-Upload-Content-Length': String(b.size) },
    payload: JSON.stringify({ name: b.name, parents: [b.folderId] }),
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) throw new Error('begin failed: HTTP ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 200));
  const uri = res.getAllHeaders()['Location'] || res.getAllHeaders()['location'];
  if (!uri) throw new Error('no session URI returned');
  return { ok: true, sessionUri: uri };
}

function chunk_(b) {
  const bytes = decode_(b.data, 'chunk ' + b.start + '-' + b.end);
  if (bytes.length !== (b.end - b.start + 1)) throw new Error('chunk ' + b.start + '-' + b.end + ': received ' + bytes.length + ' bytes — ส่งซ้ำอีกครั้ง');
  const res = UrlFetchApp.fetch(b.sessionUri, {
    method: 'put',
    contentType: 'application/octet-stream',
    headers: { 'Content-Range': 'bytes ' + b.start + '-' + b.end + '/' + b.total },
    payload: bytes,
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  if (code === 308) return { ok: true, done: false };
  if (code === 200 || code === 201) { const j = JSON.parse(res.getContentText() || '{}'); return { ok: true, done: true, fileId: j.id }; }
  throw new Error('chunk failed: HTTP ' + code + ' ' + res.getContentText().slice(0, 200));
}

function finish_(b) {
  const r = b.record || {};
  const folder = DriveApp.getFolderById(b.folderId);
  logToSheet_(r, folder, b.fileCount || 0, b.drivePath || '');
  return { ok: true };
}

/* ── helpers ─────────────────────────────────────────────── */
function getOrCreateFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
function safeName_(s) { return String(s).replace(/[\\\/:*?"<>|]/g, '-').trim().slice(0, 120); }
function ddmmyyyy_(iso) {
  const d = iso ? new Date(iso + 'T00:00:00') : new Date();
  const p = function (n) { return String(n).padStart(2, '0'); };
  return p(d.getDate()) + p(d.getMonth() + 1) + d.getFullYear();
}
function logToSheet_(r, folder, fileCount, drivePath) {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('LOG_SHEET_ID'), ss;
  try { ss = id ? SpreadsheetApp.openById(id) : null; } catch (e) { ss = null; }
  if (!ss) {
    ss = SpreadsheetApp.create(CONFIG.LOG_SHEET_NAME);
    props.setProperty('LOG_SHEET_ID', ss.getId());
    ss.getActiveSheet().appendRow(['Timestamp', 'Record ID', 'Tax ID', 'Company', 'Contact', 'Phone', 'Invoice', 'Invoice Date', 'Tracking Link', 'Files', 'Created By', 'Drive Path', 'Drive Folder', 'Notes']);
    ss.getActiveSheet().setFrozenRows(1);
  }
  ss.getActiveSheet().appendRow([new Date(), r.id, r.taxId, r.company, r.contact, r.phone, r.invoice, r.invDateDisplay || r.invDate, r.trackLink, fileCount, r.createdBy, drivePath, folder.getUrl(), r.notes]);
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
