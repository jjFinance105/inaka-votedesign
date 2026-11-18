/**
 * 名刺デザイン投票 バックエンド (Google Apps Script)
 * - 投票は Google スプレッドシート「名刺デザイン投票」に自動保存されます（初回実行時に自動作成）。
 * - GET  ?action=list  : 全員分の投票を JSON で返す
 * - POST (text/plain JSON) : 1人分の投票を id で上書き保存
 * 秘密情報は含みません。
 */
var SHEET_NAME = 'votes';
var HEADERS = ['id', 'name', 'picks', 'none', 'comment', 'done', 'ts', 'updatedAt'];
var IDS = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h'];
var MAX_ROWS = 500;
var MAX_BODY = 4000;

function getSheet_() {
  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty('SSID');
  var ss = null;
  if (ssId) {
    try { ss = SpreadsheetApp.openById(ssId); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('名刺デザイン投票');
    props.setProperty('SSID', ss.getId());
  }
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.getSheets()[0];
    sh.setName(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.getRange(1, 1, sh.getMaxRows(), HEADERS.length).setNumberFormat('@'); // 数式として解釈させない
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function clean_(v, max) {
  return String(v == null ? '' : v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '').slice(0, max);
}

// 公開用ID: 投票者ID(書き込み鍵)そのものは返さず、ハッシュの先頭16桁だけを返す
function pid_(id) {
  var d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, id, Utilities.Charset.UTF_8);
  return d.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('').slice(0, 16);
}
// 表計算の数式注入対策: 先頭が = + - @ タブ 改行 の文字列は ' を付けて保存し、読み出し時に外す
function esc_(s) { return /^[=+\-@\t\r]/.test(s) ? "'" + s : s; }
function unesc_(s) { return /^'[=+\-@\t\r]/.test(s) ? s.slice(1) : s; }

function doGet(e) {
  var action = e && e.parameter && e.parameter.action;
  if (action !== 'list') return json_({ ok: true, service: 'meishi-vote' });
  var sh = getSheet_();
  var last = sh.getLastRow();
  var votes = [];
  if (last > 1) {
    var rows = sh.getRange(2, 1, Math.min(last - 1, MAX_ROWS), HEADERS.length).getValues();
    rows.forEach(function (r) {
      if (!r[0]) return;
      votes.push({
        pid: pid_(String(r[0])),
        name: unesc_(String(r[1] || '')),
        picks: r[2] ? String(r[2]).split(',') : [],
        none: String(r[3]) === 'true',
        comment: unesc_(String(r[4] || '')),
        done: String(r[5]) === 'true',
        ts: Number(r[6]) || 0
      });
    });
  }
  return json_({ ok: true, votes: votes });
}

function doPost(e) {
  try {
    var raw = (e && e.postData && e.postData.contents) || '';
    if (!raw || raw.length > MAX_BODY) return json_({ ok: false, error: 'size' });
    var d = JSON.parse(raw);
    var id = String(d.id || '');
    if (!/^[A-Za-z0-9_-]{8,64}$/.test(id)) return json_({ ok: false, error: 'id' });
    var picks = [];
    (Array.isArray(d.picks) ? d.picks : []).forEach(function (p) {
      if (typeof p === 'string' && IDS.indexOf(p) >= 0 && picks.indexOf(p) < 0 && picks.length < 2) picks.push(p);
    });
    var ts = Number(d.ts);
    if (!isFinite(ts) || ts < 0) ts = 0;
    var row = [
      id,
      esc_(clean_(d.name, 30).trim()),
      picks.join(','),
      String(d.none === true && picks.length === 0),
      esc_(clean_(d.comment, 500)),
      String(d.done === true),
      String(Math.round(ts)),
      new Date().toISOString()
    ];
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var sh = getSheet_();
      var last = sh.getLastRow();
      var target = 0;
      if (last > 1) {
        var ids = sh.getRange(2, 1, last - 1, 1).getValues();
        for (var i = 0; i < ids.length; i++) {
          if (String(ids[i][0]) === id) { target = i + 2; break; }
        }
      }
      if (!target) {
        if (last - 1 >= MAX_ROWS) return json_({ ok: false, error: 'full' });
        target = last + 1;
      }
      var range = sh.getRange(target, 1, 1, HEADERS.length);
      range.setNumberFormat('@');
      range.setValues([row]);
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: 'bad_request' });
  }
}
