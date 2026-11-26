/**
 * 投票クイズ — バックエンド（Google Apps Script）
 * スプレッドシートに「rooms」「votes」の2シートを自動生成します。
 *
 * 使い方は README を参照。ウェブアプリとしてデプロイし、
 * 発行された /exec の URL を index.html の API_URL に設定してください。
 */

// スプレッドシートに紐づけたスクリプト（拡張機能 > Apps Script）なら空のままでOK。
// 独立したスクリプトにする場合だけ、対象シートのIDを入れてください。
const SHEET_ID = '';

const S_ROOMS = 'rooms';
const S_VOTES = 'votes';
const ROOM_COLS = ['room', 'hostKey', 'title', 'phase', 'questions', 'correct', 'updatedAt'];
const VOTE_COLS = ['room', 'pid', 'name', 'answers', 'done', 'updatedAt'];
const CACHE_SEC = 5;

/* ===================== 入口 ===================== */

function doGet(e) { return handle((e && e.parameter) || {}); }

function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); }
  catch (err) { p = (e && e.parameter) || {}; }
  return handle(p);
}

function handle(p) {
  let res;
  try { res = route(p); }
  catch (err) { res = { ok: false, error: String((err && err.message) || err) }; }
  const json = JSON.stringify(res);
  if (p.callback) {
    return ContentService.createTextOutput(p.callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function route(p) {
  switch (p.action) {
    case 'ping':       return { ok: true, ver: 1 };
    case 'createRoom': return createRoom(p);
    case 'getConfig':  return getConfig(p);
    case 'saveConfig': return saveConfig(p);
    case 'setPhase':   return setPhase(p);
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
  const bk = book();
  let sh = bk.getSheetByName(name);
  if (!sh) {
    sh = bk.insertSheet(name);
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

const rooms = () => sheet(S_ROOMS, ROOM_COLS);
const votes = () => sheet(S_VOTES, VOTE_COLS);

/** 指定シートの全データ（ヘッダー除く）を配列で返す */
function rowsOf(sh) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
}

function findRoom(code) {
  const data = rowsOf(rooms());
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === code) return { row: i + 2, v: data[i] };
  }
  return null;
}

function roomObj(r) {
  return {
    room: String(r.v[0]),
    hostKey: String(r.v[1]),
    title: String(r.v[2]),
    phase: String(r.v[3]),
    questions: JSON.parse(r.v[4] || '[]'),
    correct: JSON.parse(r.v[5] || '[]')
  };
}

function touch(row, values) {
  const sh = rooms();
  sh.getRange(row, 1, 1, ROOM_COLS.length).setValues([values]);
}

function cacheKill(code) {
  try { CacheService.getScriptCache().remove('cfg_' + code); } catch (e) {}
}

/* ===================== 各アクション ===================== */

function createRoom(p) {
  const code = newCode();
  const hostKey = Utilities.getUuid();
  const nq = Math.min(50, Math.max(1, Number(p.nq) || 10));
  const nc = Math.min(10, Math.max(2, Number(p.nc) || 4));
  const KEYS = ['A','B','C','D','E','F','G','H','I','J'];
  const qs = [];
  const cor = [];
  for (let i = 0; i < nq; i++) {
    const ch = [];
    for (let j = 0; j < nc; j++) ch.push('選択肢 ' + KEYS[j]);
    qs.push({ text: '設問 ' + (i + 1), choices: ch, multi: false });
    cor.push([]);
  }
  rooms().appendRow([code, hostKey, '投票クイズ', 'edit',
    JSON.stringify(qs), JSON.stringify(cor), new Date()]);
  return { ok: true, room: code, hostKey: hostKey,
           config: { title: '投票クイズ', phase: 'edit', questions: qs }, correct: cor };
}

function newCode() {
  const a = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  for (let t = 0; t < 40; t++) {
    let s = '';
    for (let i = 0; i < 4; i++) s += a.charAt(Math.floor(Math.random() * a.length));
    if (!findRoom(s)) return s;
  }
  throw new Error('部屋コードを発行できませんでした');
}

function getConfig(p) {
  const code = String(p.room || '').toUpperCase();
  const isHost = !!p.hostKey;
  const cache = CacheService.getScriptCache();

  if (!isHost) {                       // 参加者向けは5秒キャッシュ（読み込み負荷を抑える）
    const hit = cache.get('cfg_' + code);
    if (hit) {
      const o = JSON.parse(hit);
      if (p.rev && String(p.rev) === String(o.rev)) return { ok: true, same: true, rev: o.rev };
      return o;
    }
  }

  const r = findRoom(code);
  if (!r) return { ok: false, error: 'notfound' };
  const o = roomObj(r);
  if (isHost && o.hostKey !== p.hostKey) return { ok: false, error: 'auth' };

  const rev = new Date(r.v[6]).getTime();
  const res = {
    ok: true, rev: rev,
    config: { title: o.title, phase: o.phase, questions: o.questions },
    // 正解は「出題者」か「結果公開中」のときだけ返す
    correct: (isHost || o.phase === 'results') ? o.correct : null
  };
  if (!isHost) {
    cache.put('cfg_' + code, JSON.stringify(res), CACHE_SEC);
    if (p.rev && String(p.rev) === String(rev)) return { ok: true, same: true, rev: rev };
  }
  return res;
}

function saveConfig(p) {
  const r = mustHost(p);
  const o = roomObj(r);
  const cfg = typeof p.config === 'string' ? JSON.parse(p.config) : p.config;
  const cor = typeof p.correct === 'string' ? JSON.parse(p.correct) : p.correct;
  touch(r.row, [o.room, o.hostKey, cfg.title || o.title, cfg.phase || o.phase,
    JSON.stringify(cfg.questions || o.questions),
    JSON.stringify(cor || o.correct), new Date()]);
  cacheKill(o.room);
  return { ok: true, rev: Date.now() };
}

function setPhase(p) {
  const r = mustHost(p);
  const o = roomObj(r);
  const ph = String(p.phase);
  if (['edit', 'voting', 'closed', 'results'].indexOf(ph) < 0) throw new Error('phase が不正です');
  touch(r.row, [o.room, o.hostKey, o.title, ph,
    JSON.stringify(o.questions), JSON.stringify(o.correct), new Date()]);
  cacheKill(o.room);
  return { ok: true, phase: ph };
}

function mustHost(p) {
  const r = findRoom(String(p.room || '').toUpperCase());
  if (!r) throw new Error('部屋が見つかりません');
  if (String(r.v[1]) !== String(p.hostKey || '')) throw new Error('出題者の権限がありません');
  return r;
}

function vote(p) {
  const code = String(p.room || '').toUpperCase();
  const r = findRoom(code);
  if (!r) return { ok: false, error: 'notfound' };
  const phase = String(r.v[3]);
  if (phase !== 'voting') return { ok: false, error: 'closed' };

  const pid = String(p.pid || '');
  const rec = [code, pid, String(p.name || ''),
    typeof p.answers === 'string' ? p.answers : JSON.stringify(p.answers || {}),
    p.done ? 'done' : '', new Date()];

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = votes();
    const data = rowsOf(sh);
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === code && String(data[i][1]) === pid) {
        sh.getRange(i + 2, 1, 1, VOTE_COLS.length).setValues([rec]);
        return { ok: true, updated: true };
      }
    }
    sh.appendRow(rec);
    return { ok: true, created: true };
  } finally { lock.releaseLock(); }
}

function myVote(p) {
  const code = String(p.room || '').toUpperCase();
  const pid = String(p.pid || '');
  const data = rowsOf(votes());
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === code && String(data[i][1]) === pid) {
      return { ok: true, found: true,
               answers: JSON.parse(data[i][3] || '{}'), done: data[i][4] === 'done' };
    }
  }
  return { ok: true, found: false };
}

/** 投票中でも安全に返せる「人数だけ」の情報 */
function status(p) {
  const code = String(p.room || '').toUpperCase();
  const data = rowsOf(votes());
  let peers = 0, done = 0;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) !== code) continue;
    peers++;
    if (data[i][4] === 'done') done++;
  }
  return { ok: true, peers: peers, done: done };
}

/** 集計。出題者、または結果公開中のときだけ返す */
function results(p) {
  const code = String(p.room || '').toUpperCase();
  const r = findRoom(code);
  if (!r) return { ok: false, error: 'notfound' };
  const o = roomObj(r);
  const isHost = String(p.hostKey || '') === o.hostKey;
  if (!isHost && o.phase !== 'results') return { ok: false, error: 'hidden' };

  const data = rowsOf(votes());
  const list = [];
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) !== code) continue;
    list.push({ name: String(data[i][2]), answers: JSON.parse(data[i][3] || '{}'),
                done: data[i][4] === 'done' });
  }
  return { ok: true, config: { title: o.title, phase: o.phase, questions: o.questions },
           correct: o.correct, votes: list };
}

function clearVotes(p) {
  const o = roomObj(mustHost(p));
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = votes();
    const data = rowsOf(sh);
    for (let i = data.length - 1; i >= 0; i--) {
      if (String(data[i][0]) === o.room) sh.deleteRow(i + 2);
    }
  } finally { lock.releaseLock(); }
  return { ok: true };
}

function deleteRoom(p) {
  const r = mustHost(p);
  const o = roomObj(r);
  clearVotes(p);
  rooms().deleteRow(r.row);
  cacheKill(o.room);
  return { ok: true };
}

/* ===================== 動作確認用 ===================== */

function setup() {
  rooms(); votes();
  Logger.log('シートを準備しました: ' + book().getUrl());
}
