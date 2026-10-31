/**
 * 英文單字小考 — 雲端後台（Google Apps Script + Google 試算表）
 *
 * 試算表裡會有四個工作表：
 *   帳號      帳號 | 密碼 | 姓名 | 班級 | 身分（學生／老師）
 *   題庫      單元 | 打勾 | 英文 | 中文 | 例句 | 例句中文
 *   作答紀錄  每一題的作答（時間、誰、哪個字、對錯、答了什麼）
 *   單字狀態  每個學生每個字的答對／答錯次數、連續答對次數
 *
 * 安裝步驟請看 repo 裡的 SETUP.md
 */

var SHEETS = {
  users: { name: '帳號',     headers: ['帳號', '密碼', '姓名', '班級', '身分'] },
  bank:  { name: '題庫',     headers: ['單元', '打勾', '英文', '中文', '例句', '例句中文'] },
  log:   { name: '作答紀錄', headers: ['時間', '帳號', '姓名', '班級', '單字', '題型', '對錯', '學生答案', '正確答案'] },
  stats: { name: '單字狀態', headers: ['帳號', '單字', '答對次數', '答錯次數', '連續答對', '最後作答'] }
};

/** 第一次安裝時，在編輯器裡選這個函式按「執行」：建立工作表和範例帳號 */
function setup() {
  var ss = SpreadsheetApp.getActive();
  Object.keys(SHEETS).forEach(function (k) { sheet_(SHEETS[k]); });
  var users = ss.getSheetByName(SHEETS.users.name);
  users.getRange('A:B').setNumberFormat('@');   // 帳號、密碼當文字（避免 0012 變成 12）
  if (users.getLastRow() < 2) {
    users.getRange(2, 1, 3, 5).setValues([
      ['teacher', '1234', '老師', '', '老師'],
      ['s001', '1234', '王小明', '七年一班', '學生'],
      ['s002', '1234', '林小美', '七年一班', '學生']
    ]);
  }
  ['工作表1', 'Sheet1'].forEach(function (n) {
    var sh = ss.getSheetByName(n);
    if (sh && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
  // 觸發 Drive 授權（照片／PDF 辨識要用）
  DriveApp.getRootFolder();
}

function doGet() {
  return json_({ ok: true, msg: '英文單字小考後台運作中 ✅' });
}

function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents || '{}');
    var action = req.action;
    if (action === 'ping') return json_({ ok: true });
    if (action === 'getBank') return json_(getBank_());

    var u = auth_(req.user, req.pass);
    if (!u) return json_({ ok: false, error: '帳號或密碼錯誤' });

    switch (action) {
      case 'login':       return json_({ ok: true, user: { id: u.id, name: u.name, cls: u.cls, role: u.role } });
      case 'getProgress': return json_({ ok: true, progress: getProgress_(u.id) });
      case 'saveResults': return json_(saveResults_(u, req.results));
    }

    if (u.role !== '老師') return json_({ ok: false, error: '只有老師可以使用這個功能' });
    switch (action) {
      case 'overview':      return json_(overview_());
      case 'studentDetail': return json_({ ok: true, words: getProgressList_(String(req.id || '').toLowerCase()) });
      case 'saveBank':      return json_(saveBank_(req.items));
      case 'ocr':           return json_(ocr_(req.name, req.mime, req.data));
    }
    return json_({ ok: false, error: '未知的動作：' + action });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/* ---------------- 共用 ---------------- */
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function sheet_(def) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(def.name);
  if (!sh) sh = ss.insertSheet(def.name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(def.headers);
    sh.getRange(1, 1, 1, def.headers.length).setFontWeight('bold').setBackground('#e8efff');
    sh.setFrozenRows(1);
  }
  return sh;
}
function rows_(def) {
  var sh = sheet_(def);
  var n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, def.headers.length).getValues() : [];
}
// 避免學生輸入 =xxx 被試算表當成公式
function safe_(v) {
  v = String(v == null ? '' : v).slice(0, 300);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}
function required_(wrong) { return wrong >= 3 ? 3 : 2; }   // 錯越多次，要連續答對越多次才算過關
function inWrongBook_(r, w, s) { return w > 0 && s < required_(w); }

function auth_(user, pass) {
  user = String(user || '').trim().toLowerCase();
  pass = String(pass || '').trim();
  if (!user || !pass) return null;
  var rows = rows_(SHEETS.users);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() === user && String(rows[i][1]).trim() === pass) {
      return { id: user, name: String(rows[i][2] || user), cls: String(rows[i][3] || ''),
               role: String(rows[i][4] || '').indexOf('老師') >= 0 ? '老師' : '學生' };
    }
  }
  return null;
}

/* ---------------- 題庫 ---------------- */
function getBank_() {
  var rows = rows_(SHEETS.bank);
  var lines = [], group = null;
  rows.forEach(function (r) {
    var en = String(r[2] || '').trim();
    if (!en) return;
    var g = String(r[0] || '').trim();
    if (g !== group) { group = g; if (g) lines.push('', '# ' + g); }
    var star = /^(✓|✔|v|y|yes|true|1|是|★|\*)$/i.test(String(r[1]).trim()) ? '✓ ' : '';
    var parts = [en, r[3], r[4], r[5]].map(function (x) { return String(x || '').replace(/\|/g, '／').trim(); });
    while (parts.length > 1 && !parts[parts.length - 1]) parts.pop();
    lines.push(star + parts.join(' | '));
  });
  return { ok: true, text: lines.join('\n').trim() };
}

function saveBank_(items) {
  if (!Array.isArray(items)) return { ok: false, error: '題庫格式錯誤' };
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = sheet_(SHEETS.bank);
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 6).clearContent();
    var rows = items.slice(0, 3000).map(function (x) {
      return [safe_(x.group), x.star ? '✓' : '', safe_(x.en), safe_(x.zh), safe_(x.sentence), safe_(x.sentZh)];
    });
    if (rows.length) sh.getRange(2, 1, rows.length, 6).setValues(rows);
    return { ok: true, saved: rows.length };
  } finally { lock.releaseLock(); }
}

/* ---------------- 學習紀錄 ---------------- */
function getProgressList_(id) {
  return rows_(SHEETS.stats).filter(function (r) { return String(r[0]).toLowerCase() === id; })
    .map(function (r) {
      return { key: String(r[1]), r: +r[2] || 0, w: +r[3] || 0, s: +r[4] || 0,
               t: r[5] instanceof Date ? r[5].getTime() : 0 };
    });
}
function getProgress_(id) {
  var out = {};
  getProgressList_(id).forEach(function (p) { out[p.key] = { r: p.r, w: p.w, s: p.s, t: p.t }; });
  return out;
}

function saveResults_(u, results) {
  if (!Array.isArray(results) || !results.length) return { ok: true, saved: 0 };
  results = results.slice(0, 500);
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    var log = sheet_(SHEETS.log);
    var logRows = results.map(function (r) {
      return [new Date(r.t || Date.now()), u.id, u.name, u.cls, safe_(r.en), safe_(r.type),
              r.ok ? '✓' : '✗', safe_(r.ans), safe_(r.correct)];
    });
    log.getRange(log.getLastRow() + 1, 1, logRows.length, logRows[0].length).setValues(logRows);

    var st = sheet_(SHEETS.stats);
    var data = rows_(SHEETS.stats);
    var idx = {};
    data.forEach(function (r, i) { if (String(r[0]).toLowerCase() === u.id) idx[String(r[1]).toLowerCase()] = i; });
    results.forEach(function (r) {
      var key = String(r.key || r.en || '').toLowerCase().trim().slice(0, 200);
      if (!key) return;
      if (!(key in idx)) { data.push([u.id, safe_(key), 0, 0, 0, '']); idx[key] = data.length - 1; }
      var row = data[idx[key]];
      if (r.ok) { row[2] = (+row[2] || 0) + 1; row[4] = (+row[4] || 0) + 1; }
      else      { row[3] = (+row[3] || 0) + 1; row[4] = 0; }
      row[5] = new Date(r.t || Date.now());
    });
    if (data.length) st.getRange(2, 1, data.length, 6).setValues(data);
    return { ok: true, saved: results.length };
  } finally { lock.releaseLock(); }
}

/* ---------------- 老師後台 ---------------- */
function overview_() {
  var users = rows_(SHEETS.users).filter(function (r) { return String(r[0]).trim(); });
  var students = {};
  users.forEach(function (r) {
    if (String(r[4] || '').indexOf('老師') >= 0) return;
    var id = String(r[0]).trim().toLowerCase();
    students[id] = { id: id, name: String(r[2] || id), cls: String(r[3] || ''), r: 0, w: 0, words: 0, wrongBook: 0, mastered: 0, last: 0 };
  });
  var words = {};
  rows_(SHEETS.stats).forEach(function (r) {
    var s = students[String(r[0]).toLowerCase()];
    if (!s) return;
    var R = +r[2] || 0, W = +r[3] || 0, S = +r[4] || 0, key = String(r[1]);
    s.r += R; s.w += W; s.words++;
    if (inWrongBook_(R, W, S)) s.wrongBook++;
    if (S >= 3) s.mastered++;
    var t = r[5] instanceof Date ? r[5].getTime() : 0;
    if (t > s.last) s.last = t;
    if (!words[key]) words[key] = { key: key, w: 0, n: 0 };
    words[key].w += W;
    if (inWrongBook_(R, W, S)) words[key].n++;
  });
  var top = Object.keys(words).map(function (k) { return words[k]; })
    .filter(function (x) { return x.w > 0; })
    .sort(function (a, b) { return b.n - a.n || b.w - a.w; }).slice(0, 15);
  return { ok: true, students: Object.keys(students).map(function (k) { return students[k]; }), topWrong: top };
}

/* ---------------- 照片／PDF 文字辨識（用 Google 雲端硬碟的 OCR） ---------------- */
function ocr_(name, mime, data) {
  if (typeof Drive === 'undefined') {
    return { ok: false, error: '還沒開啟 Drive API：請在 Apps Script 編輯器左邊「服務」按 ＋，加入「Drive API」後重新部署' };
  }
  if (!data) return { ok: false, error: '沒有收到檔案' };
  var blob = Utilities.newBlob(Utilities.base64Decode(data), mime || 'image/jpeg', name || 'upload');
  var file;
  if (Drive.Files.create) {
    file = Drive.Files.create({ name: 'ocr-' + Date.now(), mimeType: MimeType.GOOGLE_DOCS }, blob, { ocrLanguage: 'zh-TW' });
  } else {
    file = Drive.Files.insert({ title: 'ocr-' + Date.now(), mimeType: MimeType.GOOGLE_DOCS }, blob, { ocr: true, ocrLanguage: 'zh-TW' });
  }
  try {
    var text = DocumentApp.openById(file.id).getBody().getText();
    return { ok: true, text: text };
  } finally {
    DriveApp.getFileById(file.id).setTrashed(true);
  }
}
