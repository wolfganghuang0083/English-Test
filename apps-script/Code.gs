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

// 後台用的 Google 試算表 ID（留空 = 使用這個程式所在的試算表；都沒有就自動建立一個）
var SPREADSHEET_ID = '';
// 第一次使用時自動建立的老師帳號密碼（建立後請到「帳號」工作表修改）
var DEFAULT_TEACHER_PASSWORD = '1234';

var ss_cache_ = null;
function ss_() {
  if (ss_cache_) return ss_cache_;
  var props = PropertiesService.getScriptProperties();
  if (SPREADSHEET_ID) {
    if (props.getProperty('SHEET_ID') !== SPREADSHEET_ID) props.setProperty('SHEET_ID', SPREADSHEET_ID);   // 記住，之後更新程式不怕 ID 被清空
    return (ss_cache_ = SpreadsheetApp.openById(SPREADSHEET_ID));
  }
  var active = SpreadsheetApp.getActive();
  if (active) return (ss_cache_ = active);
  // 獨立的 Apps Script 專案又沒填 ID：用記住的試算表，沒有就自動建立一個
  var id = props.getProperty('SHEET_ID');
  if (id) return (ss_cache_ = SpreadsheetApp.openById(id));
  var ss = SpreadsheetApp.create('英文小考後台（成績與帳號）');
  props.setProperty('SHEET_ID', ss.getId());
  return (ss_cache_ = ss);
}

var SHEETS = {
  users: { name: '帳號',     headers: ['帳號', '密碼', '姓名', '班級', '身分'] },
  bank:  { name: '題庫',     headers: ['單元', '打勾', '英文', '中文', '例句', '例句中文', '所屬單字', '指定學生'] },
  log:   { name: '作答紀錄', headers: ['時間', '帳號', '姓名', '班級', '單字', '題型', '對錯', '學生答案', '正確答案', '紀錄ID', '測驗ID'] },
  stats: { name: '單字狀態', headers: ['帳號', '單字', '答對次數', '答錯次數', '連續答對', '最後作答'] }
};

/** 建立工作表和範例帳號（第一次有人連線時會自動執行，也可以在編輯器手動執行） */
function setup() {
  var ss = ss_();
  Object.keys(SHEETS).forEach(function (k) { sheet_(SHEETS[k]); });
  var users = ss.getSheetByName(SHEETS.users.name);
  if (users.getLastRow() < 2) {
    users.getRange('A:B').setNumberFormat('@');   // 帳號、密碼當文字（避免 0012 變成 12）
    users.getRange(2, 1, 4, 5).setValues([
      ['teacher', DEFAULT_TEACHER_PASSWORD, '老師', '', '老師'],
      ['s001', '1234', '範例學生1', '七年一班', '學生'],
      ['s002', '1234', '範例學生2', '七年一班', '學生'],
      ['s003', '1234', '範例學生3', '七年二班', '學生']
    ]);
  }
  ['工作表1', 'Sheet1'].forEach(function (n) {
    var sh = ss.getSheetByName(n);
    if (sh && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
}
function ensureSetup_() {
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('TZ_SET')) {   // 試算表時區設成台灣，日期才不會差一天
    try { ss_().setSpreadsheetTimeZone(Session.getScriptTimeZone() || 'Asia/Taipei'); } catch (e) {}
    props.setProperty('TZ_SET', '1');
  }
  if (!props.getProperty('CLEANUP_V1')) cleanupDuplicates_(props);
  if (ss_().getSheetByName(SHEETS.users.name)) return;
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { if (!ss_().getSheetByName(SHEETS.users.name)) setup(); } finally { lock.releaseLock(); }
}

/** 一次性清理：刪掉完全相同的重複作答紀錄（同一時間、同帳號、同單字、同答案），並修正「單字狀態」的次數 */
function cleanupDuplicates_(props) {
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    if (props.getProperty('CLEANUP_V1')) return;
    var ss = ss_();
    var log = ss.getSheetByName(SHEETS.log.name);
    var n = log ? log.getLastRow() - 1 : 0;
    if (n > 0) {
      var width = SHEETS.log.headers.length;
      var data = log.getRange(2, 1, n, width).getValues();
      var seen = {}, keep = [], removed = [];
      data.forEach(function (r) {
        var t = r[0] instanceof Date ? r[0].getTime() : String(r[0]);
        var k = r[9] ? 'id:' + r[9] : [t, r[1], r[4], r[5], r[6], r[7]].join('\u0001');
        if (seen[k]) removed.push(r); else { seen[k] = 1; keep.push(r); }
      });
      if (removed.length) {
        log.getRange(2, 1, n, width).clearContent();
        log.getRange(2, 1, keep.length, width).setValues(keep);
        var st = ss.getSheetByName(SHEETS.stats.name);
        var m = st ? st.getLastRow() - 1 : 0;
        if (m > 0) {
          var stats = st.getRange(2, 1, m, 6).getValues();
          removed.forEach(function (r) {
            for (var i = 0; i < stats.length; i++) {
              if (String(stats[i][0]).toLowerCase() === String(r[1]).toLowerCase() &&
                  String(stats[i][1]).toLowerCase() === String(r[4]).toLowerCase()) {
                if (r[6] === '✓') { stats[i][2] = Math.max(0, (+stats[i][2] || 0) - 1); stats[i][4] = Math.max(0, (+stats[i][4] || 0) - 1); }
                else stats[i][3] = Math.max(0, (+stats[i][3] || 0) - 1);
                break;
              }
            }
          });
          st.getRange(2, 1, m, 6).setValues(stats);
        }
      }
    }
    props.setProperty('CLEANUP_V1', '1');
  } finally { lock.releaseLock(); }
}

function doGet() {
  ensureSetup_();
  return json_({ ok: true, msg: '英文單字小考後台運作中 ✅' });
}

function doPost(e) {
  try {
    ensureSetup_();
    var req = JSON.parse(e.postData.contents || '{}');
    var action = req.action;
    if (action === 'ping') return json_({ ok: true });
    var u = auth_(req.user, req.pass);
    if (!u) return json_({ ok: false, error: '帳號或密碼錯誤' });
    if (action === 'getBank') return json_(getBank_(u));   // 要登入；學生只拿得到指定給自己（或全部學生）的單元

    switch (action) {
      case 'login':       return json_({ ok: true, user: { id: u.id, name: u.name, cls: u.cls, role: u.role } });
      case 'getProgress': return json_({ ok: true, progress: getProgress_(u.id) });
      case 'getHistory':  return json_({ ok: true, sessions: getHistory_(u.id) });
      case 'saveResults': return json_(saveResults_(u, req.results));
    }

    if (u.role !== '老師') return json_({ ok: false, error: '只有老師可以使用這個功能' });
    switch (action) {
      case 'overview':      return json_(overview_());
      case 'listStudents':  return json_({ ok: true, students: rows_(SHEETS.users).filter(function (r) { return String(r[0]).trim() && String(r[4] || '').indexOf('老師') < 0; })
                              .map(function (r) { return { id: String(r[0]).trim().toLowerCase(), name: String(r[2] || r[0]), cls: String(r[3] || '') }; }) });
      case 'studentDetail': {
        var sid = String(req.id || '').toLowerCase();
        return json_({ ok: true, words: getProgressList_(sid), sessions: getHistory_(sid).slice(0, 20), wrongAns: getWrongAnswers_(sid) });
      }
      case 'saveBank':      return json_(saveBank_(req.items));
      case 'ocr':           return json_(ocr_(req.name, req.mime, req.data));
      case 'getAiKey':      return json_({ ok: true, key: PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY') || '' });
      case 'setAiKey': {
        var key = String(req.key || '').trim();
        if (key && !/^sk-ant-[\w-]{20,}$/.test(key)) return json_({ ok: false, error: '金鑰格式不對，應該是 sk-ant- 開頭' });
        if (key) PropertiesService.getScriptProperties().setProperty('ANTHROPIC_API_KEY', key);
        else PropertiesService.getScriptProperties().deleteProperty('ANTHROPIC_API_KEY');
        return json_({ ok: true });
      }
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
  var ss = ss_();
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
function whoList_(v) {
  return String(v || '').toLowerCase().split(/[\s,，、;；]+/).map(function (x) { return x.trim(); }).filter(String);
}
function getBank_(u) {
  var rows = rows_(SHEETS.bank);
  // 題庫是空的：寫入預設題庫（只做一次，老師之後清空也不會再自動寫入）
  if (!rows.length && typeof SEED_BANK !== 'undefined' && SEED_BANK.length) {
    var props = PropertiesService.getScriptProperties();
    if (!props.getProperty('BANK_SEEDED')) {
      saveBank_(SEED_BANK);
      props.setProperty('BANK_SEEDED', '1');
      rows = rows_(SHEETS.bank);
    }
  }
  var teacher = u && u.role === '老師';
  var lines = [], group = null;
  rows.forEach(function (r) {
    var en = String(r[2] || '').trim();
    if (!en) return;
    var who = whoList_(r[7]);
    if (!teacher && who.length && who.indexOf(u.id) < 0) return;   // 不是指定給這位學生的單元
    var g = String(r[0] || '').trim();
    if (g !== group) {
      group = g;
      // 老師看得到每個單元指定給誰（# Unit 3 @s001 @s002）；學生只看到單元名稱
      if (g) lines.push('', '# ' + g + (teacher && who.length ? ' @' + who.join(' @') : ''));
    }
    var star = /^(✓|✔|v|y|yes|true|1|是|★|\*)$/i.test(String(r[1]).trim()) ? '✓ ' : '';
    var parts = [en, r[3], r[4], r[5]].map(function (x) { return String(x || '').replace(/\|/g, '／').trim(); });
    while (parts.length > 1 && !parts[parts.length - 1]) parts.pop();
    var sub = String(r[6] || '').trim() ? '- ' : '';   // 片語、例句等延伸題，掛在某個單字底下
    lines.push(sub + star + parts.join(' | '));
  });
  return { ok: true, text: lines.join('\n').trim() };
}

function saveBank_(items) {
  if (!Array.isArray(items)) return { ok: false, error: '題庫格式錯誤' };
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = sheet_(SHEETS.bank);
    if (sh.getRange(1, 7).getValue() !== '所屬單字') sh.getRange(1, 7).setValue('所屬單字').setFontWeight('bold').setBackground('#e8efff');
    if (sh.getRange(1, 8).getValue() !== '指定學生') sh.getRange(1, 8).setValue('指定學生').setFontWeight('bold').setBackground('#e8efff');
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 8).clearContent();
    var rows = items.slice(0, 3000).map(function (x) {
      return [safe_(x.group), x.star ? '✓' : '', safe_(x.en), safe_(x.zh), safe_(x.sentence), safe_(x.sentZh), safe_(x.head),
              whoList_(Array.isArray(x.who) ? x.who.join(',') : x.who).join(',')];
    });
    if (rows.length) sh.getRange(2, 1, rows.length, 8).setValues(rows);
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
    if (log.getRange(1, 10).getValue() !== '紀錄ID') log.getRange(1, 10).setValue('紀錄ID').setFontWeight('bold').setBackground('#e8efff');
    if (log.getRange(1, 11).getValue() !== '測驗ID') log.getRange(1, 11).setValue('測驗ID').setFontWeight('bold').setBackground('#e8efff');
    // 同一筆紀錄重送（網路逾時後補傳）時不要重複記錄
    var n = log.getLastRow() - 1, seen = {};
    if (n > 0) {
      var start = Math.max(2, log.getLastRow() - 3000 + 1);
      log.getRange(start, 10, log.getLastRow() - start + 1, 1).getValues().forEach(function (r) { if (r[0]) seen[String(r[0])] = 1; });
    }
    results = results.filter(function (r) {
      if (!r.id) return true;
      var id = String(r.id).slice(0, 40);
      if (seen[id]) return false;
      seen[id] = 1; return true;
    });
    if (!results.length) return { ok: true, saved: 0 };
    var logRows = results.map(function (r) {
      return [new Date(r.t || Date.now()), u.id, u.name, u.cls, safe_(r.en), safe_(r.type),
              r.ok ? '✓' : '✗', safe_(r.ans), safe_(r.correct), String(r.id || '').slice(0, 40), String(r.qid || '').slice(0, 40)];
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

/** 學生每個字最近答錯時寫了什麼（每個字最多 3 筆，新的在前） */
function getWrongAnswers_(id) {
  var out = {};
  rows_(SHEETS.log).forEach(function (r) {
    if (String(r[1]).toLowerCase() !== id || r[6] !== '✗') return;
    var k = String(r[4]).toLowerCase().trim();
    (out[k] = out[k] || []).push({ a: String(r[7] || ''), type: String(r[5] || ''), t: r[0] instanceof Date ? r[0].getTime() : 0 });
  });
  Object.keys(out).forEach(function (k) { out[k] = out[k].sort(function (a, b) { return b.t - a.t; }).slice(0, 3); });
  return out;
}

/** 每一次考試的成績：同一個測驗ID算一次；舊資料沒有測驗ID，就把間隔 10 分鐘以內的作答算同一次 */
function getHistory_(id) {
  var rows = rows_(SHEETS.log).filter(function (r) { return String(r[1]).toLowerCase() === id && r[0] instanceof Date; });
  rows.sort(function (a, b) { return a[0].getTime() - b[0].getTime(); });
  var sessions = {}, order = [], lastAuto = null;
  rows.forEach(function (r) {
    var t = r[0].getTime(), key;
    if (r[10]) key = 'q:' + r[10];
    else {
      if (!lastAuto || t - lastAuto.end > 10 * 60 * 1000) lastAuto = { key: 'a:' + t, end: t };
      lastAuto.end = t; key = lastAuto.key;
    }
    if (!sessions[key]) { sessions[key] = { t: t, n: 0, ok: 0, wrong: [] }; order.push(key); }
    var s = sessions[key];
    s.n++;
    if (r[6] === '✓') s.ok++; else if (s.wrong.length < 30) s.wrong.push(String(r[4]));
  });
  return order.map(function (k) { return sessions[k]; }).sort(function (a, b) { return b.t - a.t; }).slice(0, 100);
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
    if (!words[key]) words[key] = { key: key, w: 0, n: 0, who: [] };
    words[key].w += W;
    if (inWrongBook_(R, W, S)) { words[key].n++; words[key].who.push(s.name + (s.cls ? '（' + s.cls + '）' : '')); }
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
    file = Drive.Files.create({ name: 'ocr-' + Date.now(), mimeType: MimeType.GOOGLE_DOCS }, blob);   // 不指定語言：自動辨識中英文（Drive 不接受 zh-TW）
  } else {
    file = Drive.Files.insert({ title: 'ocr-' + Date.now(), mimeType: MimeType.GOOGLE_DOCS }, blob, { ocr: true });
  }
  try {
    var text = DocumentApp.openById(file.id).getBody().getText();
    return { ok: true, text: text };
  } finally {
    try { Drive.Files.remove ? Drive.Files.remove(file.id) : Drive.Files.trash(file.id); } catch (e) {}
  }
}
