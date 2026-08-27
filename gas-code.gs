/**
 * 投票クイズ — バックエンド（Google Apps Script）
 * 1問ずつ「開始 → 締め切り → 発表」を進める方式
 *
 * スプレッドシートに「rooms」「votes」の2シートを自動生成します。
 * ウェブアプリとしてデプロイし、/exec の URL を index.html の API_URL に設定してください。
 */

// スプレッドシートに紐づけたスクリプト（拡張機能 > Apps Script）なら空のままでOK。
// 独立したスクリプトにする場合だけ、対象シートのIDを入れてください。
const SHEET_ID = '';

const S_ROOMS = 'rooms';
const S_VOTES = 'votes';
const ROOM_COLS = ['room','hostKey','title','phase','questions','correct','updatedAt','cur','qphase','revealed'];
const VOTE_COLS = ['room','pid','name','answers','answered','updatedAt'];
const CACHE_SEC = 3;

/* phase  : edit（準備中） / live（進行中） / end（終了・全体成績） */
/* qphase : idle（未開始） / voting（受付中） / closed（締め切り） / results（発表済み） */

/* ===================== 入口 ===================== */

function doGet(e) { return handle((e && e.parameter) || {}); }

function doPost(e) {
  var p = {};
  try { p = JSON.parse(e.postData.contents); }
  catch (err) { p = (e && e.parameter) || {}; }
  return handle(p);
}

function handle(p) {
  var res;
  try { res = route(p); }
  catch (err) { res = { ok: false, error: String((err && err.message) || err) }; }
  var json = JSON.stringify(res);
  if (p.callback) {
    return ContentService.createTextOutput(p.callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function route(p) {
  switch (p.action) {
    case 'ping':       return { ok: true, ver: 2 };
    case 'createRoom': return createRoom(p);
    case 'getConfig':  return getConfig(p);
    case 'saveConfig': return saveConfig(p);
    case 'setPhase':   return setPhase(p);
    case 'setQ':       return setQ(p);
    case 'vote':       return vote(p);
    case 'myVote':     return myVote(p);
    case 'status':     return status(p);
    case 'results':    return results(p);
    case 'clearVotes': return clearVotes(p);
    case 'deleteRoom': return deleteRoom(p);
    default:           return { ok: false, error: 'action が不明です: ' + p.action };
  }
}

/* ===================== シート ===================== */

function book() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActive();
}

function sheet(name, cols) {
  var bk = book();
  var sh = bk.getSheetByName(name);
  if (!sh) {
    sh = bk.insertSheet(name);
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
    return sh;
  }
  // 旧バージョンのシートに列が足りなければ見出しを引き直す
  if (sh.getLastColumn() < cols.length) {
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
  }
  return sh;
}

function rooms() { return sheet(S_ROOMS, ROOM_COLS); }
function votes() { return sheet(S_VOTES, VOTE_COLS); }

function rowsOf(sh, width) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, width || sh.getLastColumn()).getValues();
}

function findRoom(code) {
  var data = rowsOf(rooms(), ROOM_COLS.length);
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === code) return { row: i + 2, v: data[i] };
  }
  return null;
}

function roomObj(r) {
  var phase = String(r.v[3] || 'edit');
  if (phase === 'voting' || phase === 'closed' || phase === 'results') phase = 'live'; // 旧データの読み替え
  var cur = (r.v[7] === '' || r.v[7] === undefined || r.v[7] === null) ? -1 : Number(r.v[7]);
  return {
    room: String(r.v[0]),
    hostKey: String(r.v[1]),
    title: String(r.v[2]),
    phase: phase,
    questions: JSON.parse(r.v[4] || '[]'),
    correct: JSON.parse(r.v[5] || '[]'),
    cur: cur,
    qphase: String(r.v[8] || 'idle'),
    revealed: JSON.parse(r.v[9] || '[]')
  };
}

function writeRoom(row, o) {
  rooms().getRange(row, 1, 1, ROOM_COLS.length).setValues([[
    o.room, o.hostKey, o.title, o.phase,
    JSON.stringify(o.questions), JSON.stringify(o.correct), new Date(),
    o.cur, o.qphase, JSON.stringify(o.revealed)
  ]]);
}

function cacheKill(code) {
  try { CacheService.getScriptCache().remove('cfg_' + code); } catch (e) {}
}

/* ===================== 各アクション ===================== */

function createRoom(p) {
  var code = newCode();
  var hostKey = Utilities.getUuid();
  var nq = Math.min(50, Math.max(1, Number(p.nq) || 10));
  var nc = Math.min(10, Math.max(2, Number(p.nc) || 4));
  var KEYS = ['A','B','C','D','E','F','G','H','I','J'];
  var qs = [], cor = [];
  for (var i = 0; i < nq; i++) {
    var ch = [];
    for (var j = 0; j < nc; j++) ch.push('選択肢 ' + KEYS[j]);
    qs.push({ text: '設問 ' + (i + 1), choices: ch, multi: false });
    cor.push([]);
  }
  rooms().appendRow([code, hostKey, '投票クイズ', 'edit',
    JSON.stringify(qs), JSON.stringify(cor), new Date(), -1, 'idle', '[]']);
  return { ok: true, room: code, hostKey: hostKey,
    config: { title: '投票クイズ', phase: 'edit', questions: qs, cur: -1, qphase: 'idle',
              revealed: [], total: qs.length },
    correct: cor };
}

function newCode() {
  var a = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  for (var t = 0; t < 40; t++) {
    var s = '';
    for (var i = 0; i < 4; i++) s += a.charAt(Math.floor(Math.random() * a.length));
    if (!findRoom(s)) return s;
  }
  throw new Error('部屋コードを発行できませんでした');
}

/**
 * 参加者には「到達済みの設問」しか渡さない。
 * 先の設問は null に伏せ、正解は発表済みのものだけ返す。
 */
function maskForGuest(o) {
  var qs = o.questions.map(function (q, i) { return i <= o.cur ? q : null; });
  var cor = o.correct.map(function (c, i) { return o.revealed.indexOf(i) >= 0 ? c : null; });
  return {
    config: { title: o.title, phase: o.phase, questions: qs, cur: o.cur, qphase: o.qphase,
              revealed: o.revealed, total: o.questions.length },
    correct: cor
  };
}

function getConfig(p) {
  var code = String(p.room || '').toUpperCase();
  var isHost = !!p.hostKey;
  var cache = CacheService.getScriptCache();

  if (!isHost) {
    var hit = cache.get('cfg_' + code);
    if (hit) {
      var c = JSON.parse(hit);
      if (p.rev && String(p.rev) === String(c.rev)) return { ok: true, same: true, rev: c.rev };
      return c;
    }
  }

  var r = findRoom(code);
  if (!r) return { ok: false, error: 'notfound' };
  var o = roomObj(r);
  if (isHost && o.hostKey !== p.hostKey) return { ok: false, error: 'auth' };

  var rev = new Date(r.v[6]).getTime();
  var res;
  if (isHost) {
    res = { ok: true, rev: rev,
      config: { title: o.title, phase: o.phase, questions: o.questions, cur: o.cur,
                qphase: o.qphase, revealed: o.revealed, total: o.questions.length },
      correct: o.correct };
    return res;
  }
  var m = maskForGuest(o);
  res = { ok: true, rev: rev, config: m.config, correct: m.correct };
  cache.put('cfg_' + code, JSON.stringify(res), CACHE_SEC);
  if (p.rev && String(p.rev) === String(rev)) return { ok: true, same: true, rev: rev };
  return res;
}

function saveConfig(p) {
  var r = mustHost(p);
  var o = roomObj(r);
  var cfg = typeof p.config === 'string' ? JSON.parse(p.config) : p.config;
  var cor = typeof p.correct === 'string' ? JSON.parse(p.correct) : p.correct;
  o.title = cfg.title || o.title;
  o.questions = cfg.questions || o.questions;
  o.correct = cor || o.correct;
  if (o.phase === 'edit') { o.cur = -1; o.qphase = 'idle'; o.revealed = []; }
  writeRoom(r.row, o);
  cacheKill(o.room);
  return { ok: true };
}

function setPhase(p) {
  var r = mustHost(p);
  var o = roomObj(r);
  var ph = String(p.phase);
  if (['edit', 'live', 'end'].indexOf(ph) < 0) throw new Error('phase が不正です');
  o.phase = ph;
  if (ph === 'edit') { o.cur = -1; o.qphase = 'idle'; }
  if (ph === 'live' && o.cur < 0) { o.cur = 0; o.qphase = 'idle'; }
  writeRoom(r.row, o);
  cacheKill(o.room);
  return { ok: true, phase: ph, cur: o.cur, qphase: o.qphase };
}

/** 現在の設問と、その設問の状態を切り替える */
function setQ(p) {
  var r = mustHost(p);
  var o = roomObj(r);
  var cur = (p.cur === undefined || p.cur === null || p.cur === '') ? o.cur : Number(p.cur);
  var qp = String(p.qphase || o.qphase);
  if (cur < 0 || cur >= o.questions.length) throw new Error('設問番号が範囲外です');
  if (['idle', 'voting', 'closed', 'results'].indexOf(qp) < 0) throw new Error('qphase が不正です');
  o.cur = cur;
  o.qphase = qp;
  o.phase = 'live';
  if (qp === 'results' && o.revealed.indexOf(cur) < 0) o.revealed.push(cur);
  writeRoom(r.row, o);
  cacheKill(o.room);
  return { ok: true, cur: o.cur, qphase: o.qphase, revealed: o.revealed };
}

function mustHost(p) {
  var r = findRoom(String(p.room || '').toUpperCase());
  if (!r) throw new Error('部屋が見つかりません');
  if (String(r.v[1]) !== String(p.hostKey || '')) throw new Error('出題者の権限がありません');
  return r;
}

/** 1問ぶんの回答を受け取り、既存の回答にマージする */
function vote(p) {
  var code = String(p.room || '').toUpperCase();
  var r = findRoom(code);
  if (!r) return { ok: false, error: 'notfound' };
  var o = roomObj(r);

  var joining = (p.q === undefined || p.q === null || p.q === '');
  var qi = Number(p.q);
  if (!joining) {
    if (o.phase !== 'live' || o.qphase !== 'voting') return { ok: false, error: 'closed' };
    if (qi !== o.cur) return { ok: false, error: 'notcurrent' };
  }

  var pid = String(p.pid || '');
  var name = String(p.name || '');
  var pick = joining ? null :
    (typeof p.pick === 'string' ? JSON.parse(p.pick || '[]') : (p.pick || []));

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = votes();
    var data = rowsOf(sh, VOTE_COLS.length);
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]) === code && String(data[i][1]) === pid) {
        var ans = JSON.parse(data[i][3] || '{}');
        if (!joining) {
          if (pick && pick.length) ans[qi] = pick; else delete ans[qi];
        }
        sh.getRange(i + 2, 1, 1, VOTE_COLS.length)
          .setValues([[code, pid, name, JSON.stringify(ans), countKeys(ans), new Date()]]);
        return { ok: true, updated: true };
      }
    }
    var fresh = {};
    if (!joining && pick && pick.length) fresh[qi] = pick;
    sh.appendRow([code, pid, name, JSON.stringify(fresh), countKeys(fresh), new Date()]);
    return { ok: true, created: true };
  } finally { lock.releaseLock(); }
}

function countKeys(o) {
  var n = 0;
  for (var k in o) if (o[k] && o[k].length) n++;
  return n;
}

function myVote(p) {
  var code = String(p.room || '').toUpperCase();
  var pid = String(p.pid || '');
  var data = rowsOf(votes(), VOTE_COLS.length);
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === code && String(data[i][1]) === pid) {
      return { ok: true, found: true, answers: JSON.parse(data[i][3] || '{}') };
    }
  }
  return { ok: true, found: false };
}

/** 投票中でも安全に返せる「人数だけ」の情報 */
function status(p) {
  var code = String(p.room || '').toUpperCase();
  var r = findRoom(code);
  if (!r) return { ok: false, error: 'notfound' };
  var o = roomObj(r);
  var data = rowsOf(votes(), VOTE_COLS.length);
  var peers = 0, answered = 0;
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) !== code) continue;
    peers++;
    var a = JSON.parse(data[i][3] || '{}')[o.cur];
    if (a && a.length) answered++;
  }
  return { ok: true, peers: peers, answered: answered, cur: o.cur, qphase: o.qphase };
}

/** 集計。出題者は全問、参加者は発表済みの設問だけ */
function results(p) {
  var code = String(p.room || '').toUpperCase();
  var r = findRoom(code);
  if (!r) return { ok: false, error: 'notfound' };
  var o = roomObj(r);
  var isHost = String(p.hostKey || '') === o.hostKey;
  if (!isHost && !o.revealed.length) return { ok: false, error: 'hidden' };

  var data = rowsOf(votes(), VOTE_COLS.length);
  var list = [];
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) !== code) continue;
    var ans = JSON.parse(data[i][3] || '{}');
    if (!isHost) {                                   // 未発表の設問は落とす
      var keep = {};
      for (var k in ans) if (o.revealed.indexOf(Number(k)) >= 0) keep[k] = ans[k];
      ans = keep;
    }
    list.push({ name: String(data[i][2]), answers: ans });
  }
  var m = isHost
    ? { config: { title: o.title, phase: o.phase, questions: o.questions, cur: o.cur,
                  qphase: o.qphase, revealed: o.revealed, total: o.questions.length },
        correct: o.correct }
    : maskForGuest(o);
  return { ok: true, config: m.config, correct: m.correct, votes: list };
}

function clearVotes(p) {
  var r = mustHost(p);
  var o = roomObj(r);
  dropVotes(o.room);
  o.revealed = []; o.cur = -1; o.qphase = 'idle'; o.phase = 'edit';
  writeRoom(r.row, o);
  cacheKill(o.room);
  return { ok: true };
}

function deleteRoom(p) {
  var r = mustHost(p);
  var o = roomObj(r);
  dropVotes(o.room);
  rooms().deleteRow(r.row);
  cacheKill(o.room);
  return { ok: true };
}

function dropVotes(code) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = votes();
    var data = rowsOf(sh, VOTE_COLS.length);
    for (var i = data.length - 1; i >= 0; i--) {
      if (String(data[i][0]) === code) sh.deleteRow(i + 2);
    }
  } finally { lock.releaseLock(); }
}

/* ===================== 動作確認用 ===================== */

function setup() {
  rooms(); votes();
  Logger.log('シートを準備しました: ' + book().getUrl());
}
