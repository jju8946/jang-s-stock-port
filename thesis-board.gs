const SHEET_POS  = 'positions';
const SHEET_META = 'meta';
const SHEET_HIST = 'history';
const SHEET_CMT  = 'comments';
const POS_HEADERS  = ['name', 'ticker', 'status', 'avgPrice', 'shares', 'idea', 'note', 'market'];
const HIST_HEADERS = ['ts', 'datetime', 'json'];
const HIST_KEEP = 300;

function doGet(e) {
  if (e.parameter && e.parameter.action === 'prices') {
    var tickers = (e.parameter.tickers || '').split(',').filter(function(t){ return t; });
    var result = fetchPricesGoogle_(tickers);
    return json_({ ok: true, prices: result });
  }
  if (e.parameter && e.parameter.action === 'comments') {
    return json_({ ok: true, comments: getComments_() });
  }
  if (e.parameter && e.parameter.action === 'addcomment') {
    var nick = String(e.parameter.nickname || '').trim().substring(0, 20);
    var content = String(e.parameter.content || '').trim().substring(0, 300);
    if (nick && content) { addComment_(nick, content); return json_({ ok: true }); }
    return json_({ ok: false });
  }
  return json_(publicBoard_());
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) {}

  if (body.action === 'auth') {
    if (!checkPw_(body.password)) return json_({ ok: false });
    return json_({ ok: true, board: fullBoard_() });
  }

  if (body.action === 'history') {
    if (!checkPw_(body.password)) return json_({ ok: false, error: 'auth' });
    return json_({ ok: true, items: history_(20) });
  }

  if (body.action === 'save') {
    if (!checkPw_(body.password)) return json_({ ok: false, error: 'auth' });
    var lock = LockService.getScriptLock();
    try {
      lock.waitLock(8000);
      snapshotCurrent_();
      var updatedAt = writeBoard_(body.board || {});
      return json_({ ok: true, updatedAt: updatedAt });
    } catch (err) {
      return json_({ ok: false, error: String(err) });
    } finally {
      try { lock.releaseLock(); } catch (e2) {}
    }
  }

  return json_({ ok: false, error: 'unknown_action' });
}

function fullBoard_() {
  var ss = SpreadsheetApp.getActive();
  var meta = ss.getSheetByName(SHEET_META);
  var pos  = ss.getSheetByName(SHEET_POS);
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
        .filter(function (r) { return String(r[0]).trim() !== '' || String(r[1]).trim() !== ''; })
        .map(function (r) {
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

function publicBoard_() {
  var full = fullBoard_();
  var total = 0;
  full.positions.forEach(function (p) { total += (p.avgPrice || 0) * (p.shares || 0); });
  return {
    title: full.title,
    overall: full.overall,
    updatedAt: full.updatedAt,
    positions: full.positions.map(function (p) {
      var amt = (p.avgPrice || 0) * (p.shares || 0);
      var w = total > 0 ? Math.round(amt / total * 1000) / 10 : 0;
      return {
        name: p.name, ticker: p.ticker, status: p.status,
        idea: p.idea, note: p.note, weight: w, market: p.market
      };
    })
  };
}

function writeBoard_(board) {
  var ss = SpreadsheetApp.getActive();
  var meta = ss.getSheetByName(SHEET_META) || ss.insertSheet(SHEET_META);
  var pos  = ss.getSheetByName(SHEET_POS)  || ss.insertSheet(SHEET_POS);
  var now = Date.now();
  meta.getRange('A1').setValue('title');     meta.getRange('B1').setValue(board.title || '');
  meta.getRange('A2').setValue('overall');   meta.getRange('B2').setValue(board.overall || '');
  meta.getRange('A3').setValue('updatedAt'); meta.getRange('B3').setValue(now);
  pos.getRange(1, 1, 1, POS_HEADERS.length).setValues([POS_HEADERS]);
  var last = pos.getLastRow();
  if (last >= 2) pos.getRange(2, 1, last - 1, POS_HEADERS.length).clearContent();
  var list = board.positions || [];
  if (list.length) {
    var out = list.map(function (p) {
      return [p.name || '', p.ticker || '', p.status || 'valid',
              Number(p.avgPrice) || 0, Number(p.shares) || 0,
              p.idea || '', p.note || '', p.market || 'KOSPI'];
    });
    pos.getRange(2, 1, out.length, POS_HEADERS.length).setValues(out);
  }
  return now;
}

function snapshotCurrent_() {
  var ss = SpreadsheetApp.getActive();
  var hist = ss.getSheetByName(SHEET_HIST) || ss.insertSheet(SHEET_HIST);
  hist.getRange(1, 1, 1, HIST_HEADERS.length).setValues([HIST_HEADERS]);
  var current = fullBoard_();
  var now = Date.now();
  hist.appendRow([now, new Date(now), JSON.stringify(current)]);
  var last = hist.getLastRow();
  var excess = last - 1 - HIST_KEEP;
  if (excess > 0) hist.deleteRows(2, excess);
}

function history_(n) {
  var ss = SpreadsheetApp.getActive();
  var hist = ss.getSheetByName(SHEET_HIST);
  if (!hist) return [];
  var last = hist.getLastRow();
  if (last < 2) return [];
  var start = Math.max(2, last - n + 1);
  var rows = hist.getRange(start, 1, last - start + 1, HIST_HEADERS.length).getValues();
  var items = rows.map(function (r) {
    var board = null;
    try { board = JSON.parse(r[2]); } catch (e) {}
    return { ts: Number(r[0]), board: board };
  }).filter(function (it) { return it.board; });
  items.reverse();
  return items;
}

function fetchPricesGoogle_(tickers) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('_prices') || ss.insertSheet('_prices');
  sheet.hideSheet();
  var formulas = tickers.map(function(t) {
    return ['=IFERROR(GOOGLEFINANCE("' + t + '","price"),0)'];
  });
  sheet.getRange(1, 1, tickers.length, 1).setFormulas(formulas);
  SpreadsheetApp.flush();
  var values = sheet.getRange(1, 1, tickers.length, 1).getValues();
  var result = {};
  tickers.forEach(function(t, i) {
    var v = Number(values[i][0]);
    if (v > 0) result[t] = v;
  });
  sheet.getRange(1, 1, tickers.length, 1).clearContent();
  return result;
}

function getComments_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(SHEET_CMT);
  if (!sheet) return [];
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var rows = sheet.getRange(2, 1, last - 1, 4).getValues();
  return rows
    .filter(function(r){ return r[0]; })
    .map(function(r){ return { ts: Number(r[0]), nickname: String(r[2]), content: String(r[3]) }; })
    .reverse();
}

function addComment_(nickname, content) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(SHEET_CMT) || ss.insertSheet(SHEET_CMT);
  if (sheet.getLastRow() < 1) sheet.appendRow(['ts', 'datetime', 'nickname', 'content']);
  var now = Date.now();
  sheet.appendRow([now, new Date(now), nickname, content]);
}

function checkPw_(input) {
  var stored = PropertiesService.getScriptProperties().getProperty('EDIT_PASSWORD');
  return stored != null && String(input) === String(stored);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function setup() {
  var ss = SpreadsheetApp.getActive();
  var meta = ss.getSheetByName(SHEET_META) || ss.insertSheet(SHEET_META);
  var pos  = ss.getSheetByName(SHEET_POS)  || ss.insertSheet(SHEET_POS);
  var hist = ss.getSheetByName(SHEET_HIST) || ss.insertSheet(SHEET_HIST);
  pos.getRange(1, 1, 1, POS_HEADERS.length).setValues([POS_HEADERS]);
  if (pos.getLastRow() < 2) {
    pos.getRange(2, 1, 2, POS_HEADERS.length).setValues([
      ['하나머티리얼즈', '166090', 'valid', 0, 0, '', '', 'KOSPI'],
      ['원익QnC',        '074600', 'watch', 0, 0, '', '', 'KOSPI']
    ]);
  }
  meta.getRange('A1').setValue('title');     meta.getRange('B1').setValue('진욱의 포트폴리오');
  meta.getRange('A2').setValue('overall');   meta.getRange('B2').setValue('');
  meta.getRange('A3').setValue('updatedAt'); meta.getRange('B3').setValue(Date.now());
  hist.getRange(1, 1, 1, HIST_HEADERS.length).setValues([HIST_HEADERS]);
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('EDIT_PASSWORD')) props.setProperty('EDIT_PASSWORD', '0000');
}
