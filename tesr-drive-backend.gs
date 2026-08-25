/**
 * TESR — Tracking & Handover backend (Google Apps Script)
 * ------------------------------------------------------
 * รับข้อมูลจาก tesr-tracking-handover.html แล้ว:
 *   1. สร้างโฟลเดอร์  TESR shop tracking and handover / <ปี> / <Tax ID ลูกค้า> / <Invoice>_<DDMMYYYY>
 *      เช่น  TESR shop tracking and handover/2026/1234567891234/INV00001_01082026
 *   2. อัปโหลดไฟล์หลักฐานทุกไฟล์ (ตั้งชื่อไฟล์ตามช่อง เช่น "01 Pre-test Photo - xxx.jpg")
 *   3. เขียน README.txt สรุปข้อมูลการส่งมอบไว้ในโฟลเดอร์
 *   4. ตั้งสิทธิ์โฟลเดอร์เป็น Anyone with the link → Viewer
 *   5. บันทึก log ลง Google Sheet (สร้างให้อัตโนมัติถ้ายังไม่มี)
 *   6. ส่งกลับ { ok:true, folderUrl, folderId }
 *
 * วิธี Deploy
 *   1. https://script.google.com → New project → วางโค้ดนี้ทั้งหมด
 *   2. (ถ้าต้องการ) แก้ค่า CONFIG ด้านล่าง
 *   3. Deploy → New deployment → Type: Web app
 *        Execute as: Me   ·   Who has access: Anyone
 *   4. Copy URL ที่ลงท้ายด้วย /exec → วางใน Settings ของหน้าเว็บ → เลือก Live mode
 *
 * หมายเหตุ: หน้าเว็บส่ง POST แบบ text/plain เพื่อเลี่ยง CORS preflight — ห้ามเปลี่ยน
 */

const CONFIG = {
  ROOT_FOLDER_ID:   '1j7zETooR7NAMihHJJ2ZNOzD29-bdexXV', // "TESR shop tracking and handover" (ใช้เมื่อหน้าเว็บไม่ได้ส่ง parentFolderId)
  LOG_SHEET_NAME:   'TESR Handover Log',       // ชื่อ Google Sheet สำหรับ log
  MAKE_PUBLIC:      true,                      // false = แชร์เฉพาะคนในองค์กร
};

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const r = body.record || {};
    if (!r.id || !r.company) throw new Error('record.id / record.company is required');

    // 1) Folder structure
    //    Root → year → Tax ID → INVxxxxx_DDMMYYYY
    const root = DriveApp.getFolderById(body.parentFolderId || CONFIG.ROOT_FOLDER_ID);
    const year = String(r.invDate || r.createdAt || new Date().toISOString()).slice(0, 4);
    const yearFolder = getOrCreateFolder_(root, year);
    const customerFolder = getOrCreateFolder_(yearFolder, safeName_(r.taxId || 'unknown'));
    const baseName = safeName_(r.folderName || ((r.invoice || 'INV') + '_' + ddmmyyyy_(r.invDate)));
    // ถ้ามีโฟลเดอร์ชื่อซ้ำ (ส่งซ้ำ / แก้ไข) ให้ต่อท้ายด้วย -2, -3 ...
    let name = baseName, n = 1;
    while (customerFolder.getFoldersByName(name).hasNext()) name = baseName + '-' + (++n);
    const folder = customerFolder.createFolder(name);

    // 2) Files
    const counters = {};
    (body.files || []).forEach(function (f) {
      const idx = slotIndex_(f.slot);
      counters[f.slot] = (counters[f.slot] || 0) + 1;
      const name = idx + ' ' + f.slot + (counters[f.slot] > 1 ? ' (' + counters[f.slot] + ')' : '') + ' - ' + f.name;
      const blob = Utilities.newBlob(Utilities.base64Decode(f.data), f.mime || 'application/octet-stream', name);
      folder.createFile(blob);
    });

    // 3) README
    const readme = [
      'TESR Co., Ltd. — Handover Record',
      '=================================',
      'Record ID       : ' + r.id,
      'Customer        : ' + r.company + ' (Tax ID ' + (r.taxId || '-') + ')',
      'Contact         : ' + (r.contact || '-') + ' ' + (r.phone || ''),
      'Invoice/Receipt : ' + (r.invoice || '-') + '  (' + (r.invDateDisplay || r.invDate || '-') + ')',
      'Tracking Link   : ' + (r.trackLink || '-'),
      'Drive path      : ' + year + '/' + (r.taxId || '-') + '/' + name,
      'Created         : ' + (r.createdAt || new Date().toISOString()) + ' by ' + (r.createdBy || '-'),
      '',
      'Files:',
    ].concat((r.files || []).map(function (f) { return ' - [' + f.slot + '] ' + f.name; }))
     .concat(['', 'Notes:', r.notes || '-', '', 'Robot · AI · IoT · Automation — Real-world Solutions'])
     .join('\n');
    folder.createFile('README.txt', readme, MimeType.PLAIN_TEXT);

    // 4) Sharing
    if (CONFIG.MAKE_PUBLIC) folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // 5) Log sheet
    logToSheet_(r, folder, (body.files || []).length);

    return json_({ ok: true, folderUrl: folder.getUrl(), folderId: folder.getId() });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function doGet() {
  return json_({ ok: true, service: 'TESR Handover backend', time: new Date().toISOString() });
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
function slotIndex_(slot) {
  const order = ['Pre-test Photo', 'Test Video', 'Inside Box Photo', 'Packaging Photo', 'Other'];
  const i = order.indexOf(slot);
  return String(i < 0 ? 9 : i + 1).padStart(2, '0');   // unlabeled files → 09 Evidence - name
}
function logToSheet_(r, folder, fileCount) {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('LOG_SHEET_ID'), ss;
  try { ss = id ? SpreadsheetApp.openById(id) : null; } catch (e) { ss = null; }
  if (!ss) {
    ss = SpreadsheetApp.create(CONFIG.LOG_SHEET_NAME);
    props.setProperty('LOG_SHEET_ID', ss.getId());
    ss.getActiveSheet().appendRow(['Timestamp', 'Record ID', 'Tax ID', 'Company', 'Contact', 'Phone', 'Invoice', 'Invoice Date', 'Tracking Link', 'Files', 'Created By', 'Drive Path', 'Drive Folder', 'Notes']);
    ss.getActiveSheet().setFrozenRows(1);
  }
  ss.getActiveSheet().appendRow([new Date(), r.id, r.taxId, r.company, r.contact, r.phone, r.invoice, r.invDateDisplay || r.invDate, r.trackLink, fileCount, r.createdBy, folder.getParents().next().getParents().next().getName() + '/' + r.taxId + '/' + folder.getName(), folder.getUrl(), r.notes]);
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
