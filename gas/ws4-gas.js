/**
 * WS-4 アプリ バックエンド（Google Apps Script）
 * S-TEC Output 4 / 2026-10-07, 08
 *
 * 役割
 *   - 参加者の回答をスプレッドシートに保存する
 *   - クメール語の回答を英語に自動翻訳して同じ行に入れる
 *   - 運営の操作（段階公開・タイマー・言語・サマリー）を全端末に配る
 *   - 同じ内容の二重送信を弾き、内容が変わったときだけ版を増やす
 *
 * 使い方
 *   1. スプレッドシートを新規作成する
 *   2. 拡張機能 → Apps Script を開き、このコードを全部貼る
 *   3. 関数 setup を一度だけ実行する（シートと見出しができる）
 *   4. デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 *        次のユーザーとして実行： 自分
 *        アクセスできるユーザー： 全員
 *   5. 出てきた URL を Claude に渡す
 *
 * コードを直したときは、必ず「デプロイを管理」→ 鉛筆 → バージョン「新バージョン」
 * で更新すること。保存しただけでは反映されない。
 */

var NOTES  = 'WS4Notes';
var CONFIG = 'WS4Config';

var NOTES_HEAD = [
  'timestamp', 'day', 'batch', 'qid', 'subject', 'group',
  'scope', 'lang', 'text', 'text_en',
  'revision', 'is_latest', 'key', 'client'
];

/** 最初に一度だけ実行する */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var n = ss.getSheetByName(NOTES) || ss.insertSheet(NOTES);
  if (n.getLastRow() === 0) {
    n.appendRow(NOTES_HEAD);
    n.setFrozenRows(1);
  }
  var c = ss.getSheetByName(CONFIG) || ss.insertSheet(CONFIG);
  if (c.getLastRow() === 0) {
    c.appendRow(['key', 'value']);
    c.setFrozenRows(1);
  }
  return 'ok';
}

/* ───────────────── 受信 ───────────────── */

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return json_({ ok: false, error: 'busy' });
  }
  try {
    var d = JSON.parse(e.postData.contents || '{}');

    if (d.type === 'state')   { setState_(d.state || {}); return json_({ ok: true }); }
    if (d.type === 'summary') { setConfig_('summary', String(d.text || '')); return json_({ ok: true }); }
    if (d.type === 'submit')  { return json_(submit_(d)); }

    return json_({ ok: false, error: 'unknown type' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/**
 * まとめ送信。
 * 1グループ・1セッションぶんの回答が answers にまとまって届く。
 * 同じ内容なら何もしない。内容が変わっていたら版を1つ上げて追記し、
 * 前の版の is_latest を FALSE にする。
 */
function submit_(d) {
  var day     = String(d.day || '');
  var batch   = String(d.batch || '');
  var subject = String(d.subject || '');
  var group   = String(d.group || '');
  var scope   = String(d.scope || 'group');
  var lang    = String(d.lang || '');
  var client  = String(d.client || '');
  var answers = Array.isArray(d.answers) ? d.answers : [];

  if (!batch || !answers.length) return { ok: false, error: 'empty' };

  // 個人回答は人ごとに別レコードにしたいので、鍵に端末を含める
  var key = scope === 'individual'
    ? [day, batch, subject, group, client].join('|')
    : [day, batch, subject, group].join('|');

  var payload = answers.map(function (a) {
    return String(a.qid || '') + '\u0001' + String(a.text || '').slice(0, 2000);
  }).join('\u0002');
  var stamp = hash_(payload);

  var sh   = sheet_(NOTES);
  var last = sh.getLastRow();
  var rev  = 0;
  var rows = [];

  if (last > 1) {
    var vals = sh.getRange(2, 1, last - 1, NOTES_HEAD.length).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (String(vals[i][12]) === key) {          // key 列
        rows.push(i + 2);
        if (Number(vals[i][10]) > rev) rev = Number(vals[i][10]);   // revision 列
      }
    }
  }

  // 直前の版と同じ内容なら無視する（連打対策）
  if (rev > 0 && getConfig_('h:' + key) === stamp) {
    return { ok: true, duplicate: true, revision: rev };
  }

  // 前の版を最新ではなくする
  for (var k = 0; k < rows.length; k++) {
    sh.getRange(rows[k], 12).setValue(false);     // is_latest 列
  }

  var now = new Date();
  var out = [];
  for (var m = 0; m < answers.length; m++) {
    var src = String(answers[m].text || '').slice(0, 2000);
    out.push([
      now, day, batch, String(answers[m].qid || ''), subject, group,
      scope, lang, src, tr_(src, 'en'),
      rev + 1, true, key, client
    ]);
  }
  sh.getRange(sh.getLastRow() + 1, 1, out.length, NOTES_HEAD.length).setValues(out);
  setConfig_('h:' + key, stamp);

  return { ok: true, revision: rev + 1, saved: out.length };
}

/* ───────────────── 配信 ───────────────── */

function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    if (p.action === 'notes') return json_({ ok: true, notes: notes_(p.day || '') });
    return json_({ ok: true, state: getState_() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** ボード用。最新版のみ、直近400件 */
function notes_(day) {
  var sh = sheet_(NOTES);
  var last = sh.getLastRow();
  if (last < 2) return [];

  var from = Math.max(2, last - 399);
  var vals = sh.getRange(from, 1, last - from + 1, NOTES_HEAD.length).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var v = vals[i];
    if (v[11] !== true && String(v[11]).toUpperCase() !== 'TRUE') continue;
    if (day && String(v[1]) !== String(day)) continue;
    out.push({
      qid: String(v[3]), subject: String(v[4]), group: String(v[5]),
      scope: String(v[6]), text: String(v[8]), text_en: String(v[9]),
      batch: String(v[2]), rev: Number(v[10])
    });
  }
  return out;
}

/* ───────────────── 運営の状態 ───────────────── */

function getState_() {
  return {
    day:      getConfig_('day')  || '1',
    lang:     getConfig_('lang') || 'km',
    lock:     getConfig_('lock') === '0' ? 0 : 1,
    open:     getConfig_('open') || '',
    timerEnd: Number(getConfig_('timerEnd') || 0),
    timerRun: getConfig_('timerRun') === '1' ? 1 : 0,
    summary:  getConfig_('summary') || ''
  };
}

function setState_(s) {
  var keys = ['day', 'lang', 'lock', 'open', 'timerEnd', 'timerRun'];
  for (var i = 0; i < keys.length; i++) {
    if (s[keys[i]] !== undefined && s[keys[i]] !== null) {
      setConfig_(keys[i], String(s[keys[i]]));
    }
  }
}

/* ───────────────── 小さな道具 ───────────────── */

function sheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) { setup(); sh = ss.getSheetByName(name); }
  return sh;
}

function getConfig_(k) {
  var sh = sheet_(CONFIG);
  var last = sh.getLastRow();
  if (last < 2) return '';
  var vals = sh.getRange(2, 1, last - 1, 2).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) === k) return String(vals[i][1]);
  }
  return '';
}

function setConfig_(k, v) {
  var sh = sheet_(CONFIG);
  var last = sh.getLastRow();
  if (last > 1) {
    var vals = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (String(vals[i][0]) === k) { sh.getRange(i + 2, 2).setValue(v); return; }
    }
  }
  sh.appendRow([k, v]);
}

/** クメール語などを英語にする。失敗しても送信は止めない */
function tr_(src, to) {
  if (!src) return '';
  try { return LanguageApp.translate(src, '', to); }
  catch (err) { return ''; }
}

function hash_(s) {
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s, Utilities.Charset.UTF_8);
  var out = '';
  for (var i = 0; i < b.length; i++) {
    out += ('0' + (b[i] & 0xFF).toString(16)).slice(-2);
  }
  return out;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
