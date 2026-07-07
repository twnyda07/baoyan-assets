/**
 * 寶嚴禪寺財產管理系統 — 後端 (Google Apps Script Web App)
 * 資料存於 Google 試算表（首次執行自動建立），照片存於 Google Drive 資料夾
 * 管理密碼存於 Script Properties（預設 baoyan2026）
 */

var TZ = 'Asia/Taipei';

var ASSET_HEADERS = ['id', '財產編號', '名稱', '分類', '存放位置', '數量', '狀態', '保管人', '照片', '備註', '借用人', '借出日期', '建立時間', '更新時間', '金額'];
var LOG_HEADERS = ['時間', '財產編號', '名稱', '動作', '說明'];

function props() { return PropertiesService.getScriptProperties(); }

function nowStr() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }
function todayStr() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }

function getPassword() {
  var pw = props().getProperty('adminPassword');
  if (!pw) { pw = 'baoyan2026'; props().setProperty('adminPassword', pw); }
  return pw;
}

/* ---------- 試算表與資料夾（首次自動建立） ---------- */

function getSpreadsheet() {
  var id = props().getProperty('ssId');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* 被刪除則重建 */ }
  }
  var ss = SpreadsheetApp.create('寶嚴財產管理系統資料庫');
  var sh = ss.getSheets()[0];
  sh.setName('財產清冊');
  sh.appendRow(ASSET_HEADERS);
  sh.setFrozenRows(1);
  var log = ss.insertSheet('異動紀錄');
  log.appendRow(LOG_HEADERS);
  log.setFrozenRows(1);
  props().setProperty('ssId', ss.getId());
  return ss;
}

function assetSheet() {
  var sh = getSpreadsheet().getSheetByName('財產清冊');
  // 舊資料庫沒有「金額」欄→補上表頭（既有列留白視為 0）
  if (String(sh.getRange(1, ASSET_HEADERS.length).getValue()) !== '金額') {
    sh.getRange(1, 1, 1, ASSET_HEADERS.length).setValues([ASSET_HEADERS]);
  }
  return sh;
}
function logSheet() { return getSpreadsheet().getSheetByName('異動紀錄'); }

function getPhotoFolder() {
  var id = props().getProperty('photoFolderId');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* 重建 */ }
  }
  var folder = DriveApp.createFolder('寶嚴財產照片');
  props().setProperty('photoFolderId', folder.getId());
  return folder;
}

/* ---------- 讀寫財產 ---------- */

function cellStr(v) {
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm').replace(' 00:00', '');
  }
  return String(v);
}

function rowToAsset(row) {
  return {
    id: cellStr(row[0]),
    code: cellStr(row[1]),
    name: cellStr(row[2]),
    category: cellStr(row[3]),
    location: cellStr(row[4]),
    qty: Number(row[5]) || 0,
    status: cellStr(row[6]) || '正常',
    keeper: cellStr(row[7]),
    photo: cellStr(row[8]),
    note: cellStr(row[9]),
    borrower: cellStr(row[10]),
    borrowDate: cellStr(row[11]),
    createdAt: cellStr(row[12]),
    updatedAt: cellStr(row[13]),
    amount: Number(row[14]) || 0
  };
}

function assetToRow(a) {
  return [a.id, a.code, a.name, a.category, a.location, a.qty, a.status, a.keeper,
          a.photo, a.note, a.borrower, a.borrowDate, a.createdAt, a.updatedAt, Number(a.amount) || 0];
}

function loadAssets() {
  var sh = assetSheet();
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, ASSET_HEADERS.length).getValues().map(rowToAsset);
}

// 回傳 { rowIndex(1-based), asset } 或 null
function findAsset(id) {
  var sh = assetSheet();
  var last = sh.getLastRow();
  if (last < 2) return null;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === id) {
      var row = sh.getRange(i + 2, 1, 1, ASSET_HEADERS.length).getValues()[0];
      return { rowIndex: i + 2, asset: rowToAsset(row) };
    }
  }
  return null;
}

function writeAsset(rowIndex, a) {
  assetSheet().getRange(rowIndex, 1, 1, ASSET_HEADERS.length).setValues([assetToRow(a)]);
}

function nextCode() {
  var max = 0;
  loadAssets().forEach(function (a) {
    var m = /^BY-(\d+)$/.exec(a.code);
    if (m) max = Math.max(max, +m[1]);
  });
  return 'BY-' + ('0000' + (max + 1)).slice(-4);
}

function addLog(code, name, action, detail) {
  var sh = logSheet();
  sh.insertRowBefore(2);
  sh.getRange(2, 1, 1, LOG_HEADERS.length).setValues([[nowStr(), code, name, action, detail || '']]);
  // 只保留最近 500 筆
  var last = sh.getLastRow();
  if (last > 501) sh.deleteRows(502, last - 501);
}

function recentLogs(n) {
  var sh = logSheet();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var count = Math.min(n, last - 1);
  return sh.getRange(2, 1, count, LOG_HEADERS.length).getValues().map(function (r) {
    return { time: cellStr(r[0]), code: cellStr(r[1]), name: cellStr(r[2]), action: cellStr(r[3]), detail: cellStr(r[4]) };
  });
}

/* ---------- 照片 ---------- */

function savePhoto(photo, code, name) {
  if (!photo || !photo.data) return '';
  var bytes = Utilities.base64Decode(photo.data);
  if (bytes.length > 3 * 1024 * 1024) throw new Error('照片過大（超過 3MB）');
  var mime = photo.mime || 'image/jpeg';
  var ext = mime === 'image/png' ? '.png' : '.jpg';
  var blob = Utilities.newBlob(bytes, mime, code + '_' + name + ext);
  var file = getPhotoFolder().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000';
}

function deletePhotoByUrl(url) {
  var m = /[?&]id=([^&]+)/.exec(url || '');
  if (!m) return;
  try { DriveApp.getFileById(m[1]).setTrashed(true); } catch (e) { /* 已不存在 */ }
}

/* ---------- API ---------- */

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return jsonOut({
    ok: true,
    today: todayStr(),
    assets: loadAssets(),
    logs: recentLogs(30)
  });
}

function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); }
  catch (err) { return jsonOut({ ok: false, error: '無法解析請求' }); }

  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); }
  catch (err) { return jsonOut({ ok: false, error: '系統忙碌中，請稍後再試' }); }

  var res;
  try {
    switch (req.action) {
      case 'verify': res = apiVerify(req); break;
      case 'add': res = apiAdd(req); break;
      case 'update': res = apiUpdate(req); break;
      case 'remove': res = apiRemove(req); break;
      case 'borrow': res = apiBorrow(req); break;
      case 'giveback': res = apiGiveback(req); break;
      case 'setPassword': res = apiSetPassword(req); break;
      default: res = { ok: false, error: '未知的操作' };
    }
  } catch (err) {
    res = { ok: false, error: '系統錯誤：' + err.message };
  } finally {
    lock.releaseLock();
  }
  return jsonOut(res);
}

function checkAdmin(req) { return String(req.password || '') === getPassword(); }

function apiVerify(req) {
  if (!checkAdmin(req)) return { ok: false, error: '管理密碼錯誤' };
  return { ok: true, message: '登入成功', sheetUrl: getSpreadsheet().getUrl() };
}

var VALID_STATUS = ['正常', '維修中', '報廢'];

function cleanAssetInput(req, existing) {
  var a = req.asset || {};
  var name = String(a.name || '').trim();
  if (!name) return { error: '請填寫財產名稱' };
  if (name.length > 50) return { error: '名稱過長' };
  var qty = parseInt(a.qty, 10);
  if (isNaN(qty) || qty < 0 || qty > 999999) return { error: '數量必須是 0 以上的整數' };
  var status = String(a.status || '正常');
  if (VALID_STATUS.indexOf(status) < 0) return { error: '狀態不正確' };
  var code = String(a.code || '').trim();
  if (code.length > 30) return { error: '財產編號過長' };
  var amount = Number(a.amount) || 0;
  if (amount < 0 || amount > 999999999) return { error: '金額必須是 0 以上的數字' };
  var fields = {
    code: code,
    name: name,
    category: String(a.category || '其他').trim().slice(0, 20) || '其他',
    location: String(a.location || '').trim().slice(0, 50),
    qty: qty,
    status: status,
    keeper: String(a.keeper || '').trim().slice(0, 30),
    note: String(a.note || '').trim().slice(0, 200),
    amount: Math.round(amount)
  };
  // 編號不可重複（排除自己）
  if (fields.code) {
    var dup = loadAssets().some(function (x) {
      return x.code === fields.code && (!existing || x.id !== existing.id);
    });
    if (dup) return { error: '財產編號「' + fields.code + '」已存在' };
  }
  return { fields: fields };
}

function apiAdd(req) {
  if (!checkAdmin(req)) return { ok: false, error: '管理密碼錯誤' };
  var c = cleanAssetInput(req, null);
  if (c.error) return { ok: false, error: c.error };
  var f = c.fields;
  if (!f.code) f.code = nextCode();
  var photoUrl = '';
  if (req.photo) photoUrl = savePhoto(req.photo, f.code, f.name);
  var a = {
    id: 'a' + Date.now() + Math.floor(Math.random() * 1000),
    code: f.code, name: f.name, category: f.category, location: f.location,
    qty: f.qty, status: f.status, keeper: f.keeper, photo: photoUrl, note: f.note,
    borrower: '', borrowDate: '', createdAt: nowStr(), updatedAt: nowStr(), amount: f.amount
  };
  assetSheet().appendRow(assetToRow(a));
  addLog(a.code, a.name, '新增', '分類：' + a.category + '｜位置：' + a.location + '｜數量：' + a.qty + (a.amount ? '｜金額：' + a.amount : ''));
  return { ok: true, message: '已新增財產「' + a.name + '」（編號 ' + a.code + '）', code: a.code };
}

function apiUpdate(req) {
  if (!checkAdmin(req)) return { ok: false, error: '管理密碼錯誤' };
  var found = findAsset(String(req.id || ''));
  if (!found) return { ok: false, error: '找不到這筆財產' };
  var old = found.asset;
  var c = cleanAssetInput(req, old);
  if (c.error) return { ok: false, error: c.error };
  var f = c.fields;
  if (!f.code) f.code = old.code || nextCode();
  var photoUrl = old.photo;
  if (req.removePhoto) { deletePhotoByUrl(old.photo); photoUrl = ''; }
  if (req.photo) {
    deletePhotoByUrl(old.photo);
    photoUrl = savePhoto(req.photo, f.code, f.name);
  }
  var a = {
    id: old.id, code: f.code, name: f.name, category: f.category, location: f.location,
    qty: f.qty, status: f.status, keeper: f.keeper, photo: photoUrl, note: f.note,
    borrower: old.borrower, borrowDate: old.borrowDate,
    createdAt: old.createdAt, updatedAt: nowStr(), amount: f.amount
  };
  writeAsset(found.rowIndex, a);
  addLog(a.code, a.name, '修改', '');
  return { ok: true, message: '已更新「' + a.name + '」' };
}

function apiRemove(req) {
  if (!checkAdmin(req)) return { ok: false, error: '管理密碼錯誤' };
  var found = findAsset(String(req.id || ''));
  if (!found) return { ok: false, error: '找不到這筆財產' };
  deletePhotoByUrl(found.asset.photo);
  assetSheet().deleteRow(found.rowIndex);
  addLog(found.asset.code, found.asset.name, '刪除', '');
  return { ok: true, message: '已刪除「' + found.asset.name + '」' };
}

function apiBorrow(req) {
  var borrower = String(req.borrower || '').trim();
  if (!borrower) return { ok: false, error: '請填寫借用人姓名' };
  if (borrower.length > 30) return { ok: false, error: '姓名過長' };
  var found = findAsset(String(req.id || ''));
  if (!found) return { ok: false, error: '找不到這筆財產' };
  var a = found.asset;
  if (a.borrower) return { ok: false, error: '「' + a.name + '」目前由 ' + a.borrower + ' 借用中，請先歸還' };
  if (a.status === '報廢') return { ok: false, error: '此財產已報廢，無法借用' };
  a.borrower = borrower;
  a.borrowDate = todayStr();
  a.updatedAt = nowStr();
  writeAsset(found.rowIndex, a);
  var note = String(req.note || '').trim().slice(0, 100);
  addLog(a.code, a.name, '借出', '借用人：' + borrower + (note ? '｜' + note : ''));
  return { ok: true, message: '已登記借出：「' + a.name + '」→ ' + borrower };
}

function apiGiveback(req) {
  var found = findAsset(String(req.id || ''));
  if (!found) return { ok: false, error: '找不到這筆財產' };
  var a = found.asset;
  if (!a.borrower) return { ok: false, error: '「' + a.name + '」目前沒有借出紀錄' };
  var borrower = a.borrower, since = a.borrowDate;
  a.borrower = '';
  a.borrowDate = '';
  a.updatedAt = nowStr();
  writeAsset(found.rowIndex, a);
  addLog(a.code, a.name, '歸還', borrower + ' 歸還（' + since + ' 借出）');
  return { ok: true, message: '「' + a.name + '」已歸還' };
}

function apiSetPassword(req) {
  if (!checkAdmin(req)) return { ok: false, error: '管理密碼錯誤' };
  var np = String(req.newPassword || '');
  if (np.length < 6) return { ok: false, error: '新密碼至少 6 個字元' };
  props().setProperty('adminPassword', np);
  return { ok: true, message: '已更改管理密碼' };
}
