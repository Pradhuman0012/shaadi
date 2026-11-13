// ===== Wedding Planner: Google Sheet + Drive photos =====
// Change FAMILY_CODE to your own code. This is the code family members type to open the app.
const FAMILY_CODE = "shaadi-7171";
const SHEET_NAME = "data";
const FOLDER_NAME = "Wedding Planner Photos";

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME); sh.appendRow(["kind", "id", "json", "updated"]); }
  return sh;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function folder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}
function doGet(e) {
  if (String(e.parameter.code || "") !== FAMILY_CODE) return out_({ ok: false, error: "code" });
  if (e.parameter.photo) {
    try {
      const f = DriveApp.getFileById(String(e.parameter.photo));
      const inFolder = f.getParents().hasNext() && f.getParents().next().getName() === FOLDER_NAME;
      if (!inFolder || f.isTrashed()) return out_({ ok: false, error: "notfound" });
      const b = f.getBlob();
      return out_({ ok: true, mime: b.getContentType(), data: Utilities.base64Encode(b.getBytes()) });
    } catch (err) { return out_({ ok: false, error: "notfound" }); }
  }
  const rows = sheet_().getDataRange().getValues();
  const res = { ok: true, items: {}, custom: {} };
  for (let i = 1; i < rows.length; i++) {
    const k = rows[i][0], id = rows[i][1], j = rows[i][2];
    if (res[k] && id) { try { res[k][String(id)] = JSON.parse(j); } catch (err) {} }
  }
  return out_(res);
}
function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: "bad" }); }
  if (String(req.code || "") !== FAMILY_CODE) return out_({ ok: false, error: "code" });

  if (req.action === "upload") {
    try {
      const blob = Utilities.newBlob(Utilities.base64Decode(req.data), req.mime || "image/jpeg", req.name || "photo.jpg");
      const file = folder_().createFile(blob);
      return out_({ ok: true, id: file.getId() });
    } catch (err) { return out_({ ok: false, error: "upload" }); }
  }
  if (req.action === "photo_del") {
    try {
      const f = DriveApp.getFileById(String(req.id));
      if (f.getParents().hasNext() && f.getParents().next().getName() === FOLDER_NAME) f.setTrashed(true);
      return out_({ ok: true });
    } catch (err) { return out_({ ok: false, error: "notfound" }); }
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    const rows = sh.getDataRange().getValues();
    const idx = {};
    for (let i = 1; i < rows.length; i++) idx[rows[i][0] + "/" + rows[i][1]] = i + 1;
    const now = new Date(), del = [];
    (req.ops || []).slice(0, 500).forEach(function (op) {
      if (op.k !== "items" && op.k !== "custom") return;
      const key = op.k + "/" + op.id, r = idx[key];
      if (op.b === null) { if (r) del.push(r); return; }
      const vals = [[op.k, String(op.id), JSON.stringify(op.b), now]];
      if (r) sh.getRange(r, 1, 1, 4).setValues(vals);
      else { sh.appendRow(vals[0]); idx[key] = sh.getLastRow(); }
    });
    del.sort(function (a, b) { return b - a; }).forEach(function (r) { sh.deleteRow(r); });
  } finally { lock.releaseLock(); }
  return out_({ ok: true });
}
