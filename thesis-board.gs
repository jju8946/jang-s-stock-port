const POS_HEADERS  = ['name', 'ticker', 'status', 'avgPrice', 'shares', 'idea', 'note', 'market'];
const HIST_HEADERS = ['ts', 'datetime', 'json'];
const HIST_KEEP = 300;

// 사용자별 시트 이름: positions_jinwook, meta_alice 등
function sn_(base, user) { return base + '_' + user; }

/* ── GET 핸들러 ── */
function doGet(e) {
  var p = e.parameter || {};
  var user = String(p.user || '').trim();

  if (p.action === 'users') return json_(getUsers_());

  if (p.action === 'prices') {
    var tickers = (p.tickers || '').split(',').filter(function(t){ return t; });
    return json_({ ok: true, prices: fetchPricesGoogle_(tickers) });
  }

  if (p.action === 'comments') {
    if (!user) return json_({ ok: false, error: 'no user' });
    return json_({ ok: true, comments: getComments_(user) });
  }

  if (p.action === 'addcomment') {
    if (!user) return json_({ ok: false });
    var nick = String(p.nickname || '').trim().substring(0, 20);
    var content = String(p.content || '').trim().substring(0, 300);
    if (nick && content) { addComment_(user, nick, content); return json_({ ok: true }); }
    return json_({ ok: false });
  }

  if (!user) return json_({ ok: false, error: 'no user' });
  return json_(publicBoard_(user));
}

/* ── POST 핸들러 ── */
function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) {}
  var user = String(body.user || '').trim();
  if (!user) return json_({ ok: false, error: 'no user' });

  if (body.action === 'auth') {
    if (!checkPw_(user, body.password)) return json_({ ok: false });
    return json_({ ok: true, board: fullBoard_(user) });
  }

  if (body.action === 'history') {
    if (!checkPw_(user, body.password)) return json_({ ok: false, error: 'auth' });
    return json_({ ok: true, items: history_(user, 20) });
  }

  if (body.action === 'save') {
    if (!checkPw_(user, body.password)) return json_({ ok: false, error: 'auth' });
    var lock = LockService.getScriptLock();
    try {
      lock.waitLock(8000);
      snapshotCurrent_(user);
      var updatedAt = writeBoard_(user, body.board || {});
      return json_({ ok: true, updatedAt: updatedAt });
    } catch (err) {
      return json_({ ok: false, error: String(err) });
    } finally {
      try { lock.releaseLock(); } catch (e2) {}
    }
  }

  return json_({ ok: false, error: 'unknown_action' });
}

/* ── 사용자 목록 ── */
function getUsers_() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty('USERS') || '';
  var ids = raw.split(',').map(function(s){ return s.trim(); }).filter(Boolean);
  return {
    users: ids.map(function(id) {
      return { id: id, name: props.getProperty('NAME_' + id) || id };
    })
  };
}

/* ── 보드 읽기/쓰기 ── */
function fullBoard_(user) {
  var ss = SpreadsheetApp.getActive();
  var meta = ss.getSheetByName(sn_('meta', user));
  var pos  = ss.getSheetByName(sn_('positions', user));
  var board = { title: '', overall: '', updatedAt: null, positions: [] };
  if (meta) {
    board.title   = meta.getRange('B1').getValue() || '';
    board.overall = meta.getRange('B2').getValue() || '';
    var u = meta.getRange('B3').getValue();
    board.updatedAt = u ? Number(u) : null;
  }
  if (pos) {
    var last = pos.getLastRow();
    if (last >= 2) {
      var rows = pos.getRange(2, 1, last - 1, POS_HEADERS.length).getValues();
      board.positions = rows
        .filter(function(r){ return String(r[0]).trim() !== '' || String(r[1]).trim() !== ''; })
        .map(function(r){
          return {
            name:     String(r[0] || ''),
            ticker:   String(r[1] || '').padStart(6, '0'),
            status:   String(r[2] || 'valid'),
            avgPrice: Number(r[3]) || 0,
            shares:   Number(r[4]) || 0,
            idea:     String(r[5] || ''),
            note:     String(r[6] || ''),
            market:   String(r[7] || 'KOSPI')
          };
        });
    }
  }
  return board;
}

function publicBoard_(user) {
  var full = fullBoard_(user);
  var total = 0;
  full.positions.forEach(function(p){ total += (p.avgPrice || 0) * (p.shares || 0); });
  return {
    title: full.title,
    overall: full.overall,
    updatedAt: full.updatedAt,
    positions: full.positions.map(function(p){
      var amt = (p.avgPrice || 0) * (p.shares || 0);
      var w = total > 0 ? Math.round(amt / total * 1000) / 10 : 0;
      return { name: p.name, ticker: p.ticker, status: p.status, idea: p.idea, note: p.note, weight: w, market: p.market };
    })
  };
}

function writeBoard_(user, board) {
  var ss = SpreadsheetApp.getActive();
  var meta = ss.getSheetByName(sn_('meta', user))      || ss.insertSheet(sn_('meta', user));
  var pos  = ss.getSheetByName(sn_('positions', user)) || ss.insertSheet(sn_('positions', user));
  var now = Date.now();
  meta.getRange('A1').setValue('title');     meta.getRange('B1').setValue(board.title || '');
  meta.getRange('A2').setValue('overall');   meta.getRange('B2').setValue(board.overall || '');
  meta.getRange('A3').setValue('updatedAt'); meta.getRange('B3').setValue(now);
  pos.getRange(1, 1, 1, POS_HEADERS.length).setValues([POS_HEADERS]);
  var last = pos.getLastRow();
  if (last >= 2) pos.getRange(2, 1, last - 1, POS_HEADERS.length).clearContent();
  var list = board.positions || [];
  if (list.length) {
    var out = list.map(function(p){
      return [p.name||'', p.ticker||'', p.status||'valid',
              Number(p.avgPrice)||0, Number(p.shares)||0,
              p.idea||'', p.note||'', p.market||'KOSPI'];
    });
    pos.getRange(2, 1, out.length, POS_HEADERS.length).setValues(out);
  }
  return now;
}

function snapshotCurrent_(user) {
  var ss = SpreadsheetApp.getActive();
  var hist = ss.getSheetByName(sn_('history', user)) || ss.insertSheet(sn_('history', user));
  hist.getRange(1, 1, 1, HIST_HEADERS.length).setValues([HIST_HEADERS]);
  var now = Date.now();
  hist.appendRow([now, new Date(now), JSON.stringify(fullBoard_(user))]);
  var last = hist.getLastRow();
  var excess = last - 1 - HIST_KEEP;
  if (excess > 0) hist.deleteRows(2, excess);
}

function history_(user, n) {
  var ss = SpreadsheetApp.getActive();
  var hist = ss.getSheetByName(sn_('history', user));
  if (!hist) return [];
  var last = hist.getLastRow(); if (last < 2) return [];
  var start = Math.max(2, last - n + 1);
  var rows = hist.getRange(start, 1, last - start + 1, HIST_HEADERS.length).getValues();
  var items = rows.map(function(r){
    var b = null; try { b = JSON.parse(r[2]); } catch(e) {}
    return { ts: Number(r[0]), board: b };
  }).filter(function(it){ return it.board; });
  items.reverse(); return items;
}

/* ── 현재가 (GOOGLEFINANCE) ── */
function fetchPricesGoogle_(tickers) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('_prices') || ss.insertSheet('_prices');
  sheet.hideSheet();
  var formulas = tickers.map(function(t){ return ['=IFERROR(GOOGLEFINANCE("' + t + '","price"),0)']; });
  sheet.getRange(1, 1, tickers.length, 1).setFormulas(formulas);
  SpreadsheetApp.flush();
  var values = sheet.getRange(1, 1, tickers.length, 1).getValues();
  var result = {};
  tickers.forEach(function(t, i){ var v = Number(values[i][0]); if (v > 0) result[t] = v; });
  sheet.getRange(1, 1, tickers.length, 1).clearContent();
  return result;
}

/* ── 댓글 ── */
function getComments_(user) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(sn_('comments', user));
  if (!sheet) return [];
  var last = sheet.getLastRow(); if (last < 2) return [];
  var rows = sheet.getRange(2, 1, last - 1, 4).getValues();
  return rows
    .filter(function(r){ return r[0]; })
    .map(function(r){ return { ts: Number(r[0]), nickname: String(r[2]), content: String(r[3]) }; })
    .reverse();
}

function addComment_(user, nickname, content) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(sn_('comments', user)) || ss.insertSheet(sn_('comments', user));
  if (sheet.getLastRow() < 1) sheet.appendRow(['ts', 'datetime', 'nickname', 'content']);
  var now = Date.now();
  sheet.appendRow([now, new Date(now), nickname, content]);
}

/* ── 인증 ── */
function checkPw_(user, input) {
  var props = PropertiesService.getScriptProperties();
  var stored = props.getProperty('PW_' + user);
  return stored != null && String(input) === String(stored);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ══════════════════════════════════════════════════════
   관리 함수 — Apps Script 편집기에서 직접 실행하세요
   ══════════════════════════════════════════════════════ */

// 새 사용자 추가: setupUser_('alice', '앨리스의 포트폴리오', '1234')
function setupUser_(userId, displayName, password) {
  var ss = SpreadsheetApp.getActive();
  var sheets = ['meta','positions','history'];
  sheets.forEach(function(base){
    var name = sn_(base, userId);
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });
  var meta = ss.getSheetByName(sn_('meta', userId));
  meta.getRange('A1').setValue('title');     meta.getRange('B1').setValue(displayName || userId + '의 포트폴리오');
  meta.getRange('A2').setValue('overall');   meta.getRange('B2').setValue('');
  meta.getRange('A3').setValue('updatedAt'); meta.getRange('B3').setValue(Date.now());
  var pos = ss.getSheetByName(sn_('positions', userId));
  pos.getRange(1, 1, 1, POS_HEADERS.length).setValues([POS_HEADERS]);
  var hist = ss.getSheetByName(sn_('history', userId));
  hist.getRange(1, 1, 1, HIST_HEADERS.length).setValues([HIST_HEADERS]);

  var props = PropertiesService.getScriptProperties();
  props.setProperty('PW_' + userId, password || '0000');
  props.setProperty('NAME_' + userId, displayName || userId);
  var raw = props.getProperty('USERS') || '';
  var ids = raw.split(',').map(function(s){ return s.trim(); }).filter(Boolean);
  if (ids.indexOf(userId) === -1) { ids.push(userId); props.setProperty('USERS', ids.join(',')); }
  Logger.log('setupUser_ 완료: ' + userId);
}

// 기존 시트(접미사 없음) → 특정 사용자 시트로 복사
// 예: 기존 positions 시트를 positions_jinwook 으로 복사
// migrateToUser_('jinwook') 실행 후 기존 시트는 수동으로 삭제하세요
function migrateToUser_(userId) {
  var ss = SpreadsheetApp.getActive();
  var bases = ['positions','meta','history','comments'];
  bases.forEach(function(base){
    var old = ss.getSheetByName(base);
    var newName = sn_(base, userId);
    if (old && !ss.getSheetByName(newName)) {
      old.copyTo(ss).setName(newName);
      Logger.log(base + ' → ' + newName + ' 복사 완료');
    }
  });
  Logger.log('migrateToUser_ 완료. 기존 시트는 수동으로 확인 후 삭제하세요.');
}
