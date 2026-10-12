/**
 * Aster DEX — read-only history fetcher (Google Apps Script).
 * Writes Order History, Trade History and Transaction History (income) to sheets.
 *
 * Setup (Project Settings → Script Properties):
 *   ASTER_USER                main wallet address
 *   ASTER_SIGNER              API wallet address (Read permission only)
 *   ASTER_SIGNER_PRIVATE_KEY  API wallet private key — NEVER the main wallet key
 *
 * READ-ONLY: only GET requests are made. There is intentionally no POST/PUT/DELETE helper.
 */

var BASE_URL = 'https://fapi.asterdex.com';
// Oldest date to look back to (UTC). Newest data is fetched first, then history going backwards.
var START_DATE = '2024-01-01';
// Stop looking further back after this many consecutive empty weeks (account had no activity yet).
var STOP_AFTER_EMPTY_WEEKS = 12;
var WEEK_MS = 7 * 86400000 - 1;

var RUN_BUDGET_MS = 4.5 * 60 * 1000;  // GAS kills a run at 6 min; stop early and resume via trigger
// Everything in [TAIL, HEAD) is already in the sheets.
var HEAD_KEY = 'ASTER_HEAD_MS';         // newest fetched boundary; new data is fetched from here to now
var TAIL_KEY = 'ASTER_TAIL_MS';         // oldest fetched boundary; history continues backwards from here
var EMPTY_KEY = 'ASTER_EMPTY_WEEKS';    // consecutive empty weeks seen while walking backwards
var DONE_KEY = 'ASTER_HISTORY_DONE';    // backward walk finished

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Aster')
    .addItem('Fetch / continue', 'fetchAll')
    .addItem('Fetch date range…', 'showDateRangeDialog')
    .addItem('Fetch open positions', 'fetchPositions')
    .addItem('Build position history', 'buildPositionHistory')
    .addItem('Check config', 'checkConfig')
    .addItem('Reset (clear sheets, start over)', 'resetAll')
    .addToUi();
}

/**
 * 1) Fetches anything new since the last run (HEAD → now).
 * 2) Continues walking history backwards one week at a time (TAIL → START_DATE),
 *    stopping early after STOP_AFTER_EMPTY_WEEKS empty weeks in a row.
 * If time runs short it saves progress and schedules itself to continue in 1 minute.
 * Sheets are kept sorted newest-first.
 */
function fetchAll() {
  try {
    fetchAllInner_();
  } catch (e) {
    log_('ERROR: ' + e.message, 'ERROR');
    throw e;
  } finally {
    ['Orders', 'Trades', 'Transactions'].forEach(sortNewestFirst_);
  }
}

function fetchAllInner_() {
  clearTriggers_();
  var cfg = loadConfig_();
  log_('Start — user ' + mask_(cfg.user) + ', signer ' + mask_(cfg.signer));
  var props = PropertiesService.getScriptProperties();
  var started = Date.now(), now = Date.now();
  var floor = new Date(START_DATE + 'T00:00:00Z').getTime();
  var head = Number(props.getProperty(HEAD_KEY));
  if (!head) { head = now; props.setProperty(HEAD_KEY, String(head)); }  // first run: history goes back from here
  var tail = Number(props.getProperty(TAIL_KEY)) || now;
  var empty = Number(props.getProperty(EMPTY_KEY)) || 0;
  var total = { orders: 0, trades: 0, income: 0 };
  var outOfTime = function () { return Date.now() - started > RUN_BUDGET_MS; };

  // 1) New data since last run, oldest window first.
  while (head < now) {
    if (outOfTime()) return pause_('new data at ' + day_(head), total);
    var he = Math.min(head + WEEK_MS, now);
    fetchWeek_(head, he, total);
    head = he + 1;
    props.setProperty(HEAD_KEY, String(head));
  }

  // 2) History, walking backwards.
  if (!props.getProperty(DONE_KEY)) {
    while (tail > floor && empty < STOP_AFTER_EMPTY_WEEKS) {
      if (outOfTime()) return pause_('history at ' + day_(tail), total);
      var ts = Math.max(tail - WEEK_MS - 1, floor);
      empty = fetchWeek_(ts, tail - 1, total) ? 0 : empty + 1;
      tail = ts;
      props.setProperty(TAIL_KEY, String(tail));
      props.setProperty(EMPTY_KEY, String(empty));
    }
    props.setProperty(DONE_KEY, '1');
    log_('History complete back to ' + day_(tail) +
      (empty >= STOP_AFTER_EMPTY_WEEKS ? ' (no activity in the ' + empty + ' weeks before that)' : ''));
  }
  log_('Up to date as of ' + new Date(now).toISOString() + ' (this run: ' + summary_(total) + ')');
}

/** Fetches one window, writes all three sheets, returns true if anything was found. */
function fetchWeek_(ws, we, total, sheets) {
  sheets = sheets || { orders: 'Orders', trades: 'Trades', income: 'Transactions' };
  var orders = fetchRange_('/fapi/v3/allOrders', {}, ws, we);
  var income = fetchRange_('/fapi/v3/income', {}, ws, we);

  // userTrades needs a symbol: only query symbols that show activity (orders or income) this week.
  var symbols = {};
  orders.concat(income).forEach(function (r) { if (r.symbol) symbols[r.symbol] = true; });
  var trades = [];
  Object.keys(symbols).sort().forEach(function (sym) {
    trades = trades.concat(fetchRange_('/fapi/v3/userTrades', { symbol: sym }, ws, we));
  });

  // Write all three before the caller advances its cursor, so a crash mid-week just redoes that week.
  appendRows_(sheets.orders, orders);
  appendRows_(sheets.trades, trades);
  appendRows_(sheets.income, income);
  total.orders += orders.length; total.trades += trades.length; total.income += income.length;

  var found = orders.length || trades.length || income.length;
  if (found) {
    log_(day_(ws) + ' → ' + day_(we) + ': ' +
      summary_({ orders: orders.length, trades: trades.length, income: income.length }) +
      ' [' + Object.keys(symbols).sort().join(', ') + ']', 'INFO', true);
  } else {
    console.log(day_(ws) + ' → ' + day_(we) + ': empty');
  }
  return !!found;
}

// ---------- Positions ----------

/** Replaces the "Positions" sheet with a snapshot of currently open positions. */
function fetchPositions() {
  try {
    loadConfig_();
    // positionRisk returns every symbol; keep only open ones (positionAmt != 0).
    var rows = signedGet_('/fapi/v3/positionRisk', {}).filter(function (r) { return Number(r.positionAmt) !== 0; });
    var fetchedAt = new Date();
    rows = rows.map(function (r) {
      // One-way mode reports positionSide BOTH, so derive the side from the sign of positionAmt.
      var side = r.positionSide === 'LONG' || r.positionSide === 'SHORT' ? r.positionSide
        : Number(r.positionAmt) > 0 ? 'LONG' : 'SHORT';
      var out = { symbol: r.symbol, side: side };  // put side right after symbol
      Object.keys(r).forEach(function (k) { out[k] = r[k]; });
      out.positionAmt = String(r.positionAmt).replace(/^[-+]/, '');  // direction is in `side`
      out.fetchedAt = fetchedAt;
      return out;
    });
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Positions');
    if (sh) sh.clearContents();
    appendRows_('Positions', rows);
    log_(rows.length ? rows.length + ' open position(s): ' + rows.map(function (r) {
      return r.symbol + ' ' + r.side + ' ' + r.positionAmt + ' uPnL ' + Number(r.unRealizedProfit).toFixed(2);
    }).join(', ') : 'No open positions');
  } catch (e) {
    log_('ERROR: ' + e.message, 'ERROR');
    throw e;
  }
}

/**
 * Aster has no position-history endpoint, so rebuild closed positions from the "Trades" and
 * "Transactions" sheets (run Fetch / continue first). A position starts when the net size
 * for a symbol/positionSide leaves 0 and ends when it returns to 0 (a flip closes one and opens another).
 */
function buildPositionHistory() {
  try {
    var trades = readSheet_('Trades'), income = readSheet_('Transactions');
    if (!trades.length) throw new Error('Trades sheet is empty — run "Fetch / continue" first');
    var rows = reconstructPositions_(trades, income);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Position History');
    if (sh) sh.clearContents();
    appendRows_('Position History', rows);
    sortNewestFirst_('Position History');
    var pnl = rows.reduce(function (a, r) { return a + r.netPnl; }, 0);
    log_(rows.length + ' closed position(s), total net PnL ' + pnl.toFixed(2));
  } catch (e) {
    log_('ERROR: ' + e.message, 'ERROR');
    throw e;
  }
}

function reconstructPositions_(trades, income) {
  var EPS = 1e-9, out = [], groups = {};
  trades.forEach(function (t) {
    var k = t.symbol + '|' + (t.positionSide || 'BOTH');
    (groups[k] = groups[k] || []).push(t);
  });
  var funding = income.filter(function (r) { return r.incomeType === 'FUNDING_FEE'; });

  Object.keys(groups).forEach(function (k) {
    var list = groups[k].sort(function (a, b) { return ms_(a.time) - ms_(b.time) || Number(a.id) - Number(b.id); });
    var pos = 0, cur = null;
    var open = function (t, qty, dir) {
      cur = { symbol: t.symbol, positionSide: t.positionSide || 'BOTH', side: dir > 0 ? 'LONG' : 'SHORT',
        openTime: ms_(t.time), entryCost: 0, entryQty: 0, exitCost: 0, exitQty: 0, maxQty: 0,
        realizedPnl: 0, fees: 0, trades: 0 };
    };
    list.forEach(function (t) {
      var q = Number(t.qty), p = Number(t.price), dir = t.side === 'BUY' ? 1 : -1;
      if (Math.abs(pos) < EPS) open(t, q, dir);
      cur.trades++;
      cur.fees += Number(t.commission) || 0;
      cur.realizedPnl += Number(t.realizedPnl) || 0;
      if (dir * pos >= -EPS) {  // opening / adding
        cur.entryCost += p * q; cur.entryQty += q; pos += dir * q;
        cur.maxQty = Math.max(cur.maxQty, Math.abs(pos));
        return;
      }
      var c = Math.min(q, Math.abs(pos));  // reducing / closing
      cur.exitCost += p * c; cur.exitQty += c; pos += dir * c;
      if (Math.abs(pos) < EPS) {
        pos = 0;
        cur.closeTime = ms_(t.time);
        out.push(finishPosition_(cur, funding));
        var left = q - c;
        if (left > EPS) {  // flipped: the rest of this trade opens the opposite side
          open(t, left, dir);
          cur.entryCost = p * left; cur.entryQty = left; cur.maxQty = left; pos = dir * left;
        }
      }
    });
  });
  return out;
}

function finishPosition_(c, funding) {
  var fund = funding.filter(function (f) {
    return f.symbol === c.symbol && ms_(f.time) >= c.openTime && ms_(f.time) <= c.closeTime;
  }).reduce(function (a, f) { return a + Number(f.income); }, 0);
  return {
    symbol: c.symbol, side: c.side, positionSide: c.positionSide,
    openTime: new Date(c.openTime), closeTime: new Date(c.closeTime),
    maxQty: round_(c.maxQty), avgEntry: round_(c.entryCost / c.entryQty), avgExit: round_(c.exitCost / c.exitQty),
    realizedPnl: round_(c.realizedPnl), fees: round_(-c.fees), funding: round_(fund),
    netPnl: round_(c.realizedPnl - c.fees + fund), trades: c.trades
  };
}

function readSheet_(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  var v = sh.getDataRange().getValues(), h = v.shift();
  return v.map(function (r) { var o = {}; h.forEach(function (k, i) { o[k] = r[i]; }); return o; });
}
function ms_(t) { return t instanceof Date ? t.getTime() : Number(t); }
function round_(x) { return Math.round(x * 1e8) / 1e8; }

// ---------- Pick-date fetch ----------

var RANGE_SHEETS_ = { orders: 'Range Orders', trades: 'Range Trades', income: 'Range Transactions' };

/** Opens a small dialog with two date pickers; submitting calls fetchDateRange(from, to). */
function showDateRangeDialog() {
  var today = Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd');
  var html = HtmlService.createHtmlOutput(
    '<style>body{font-family:Arial,sans-serif;font-size:13px}label{display:block;margin:8px 0 2px}' +
    'input{width:100%;padding:4px}button{margin-top:12px;padding:6px 14px}#s{margin-top:8px;color:#555}</style>' +
    '<label>From (UTC)</label><input type="date" id="f" value="' + today + '" max="' + today + '">' +
    '<label>To (UTC, inclusive)</label><input type="date" id="t" value="' + today + '" max="' + today + '">' +
    '<button id="b" onclick="go()">Fetch</button><div id="s"></div>' +
    '<script>function go(){var f=document.getElementById("f").value,t=document.getElementById("t").value;' +
    'if(!f||!t||f>t){document.getElementById("s").textContent="Pick a valid range";return;}' +
    'document.getElementById("b").disabled=true;document.getElementById("s").textContent="Fetching…";' +
    'google.script.run.withSuccessHandler(function(m){document.getElementById("s").textContent=m;' +
    'document.getElementById("b").disabled=false;})' +
    '.withFailureHandler(function(e){document.getElementById("s").textContent="Error: "+e.message;' +
    'document.getElementById("b").disabled=false;}).fetchDateRange(f,t);}</script>'
  ).setWidth(280).setHeight(220);
  SpreadsheetApp.getUi().showModalDialog(html, 'Aster — fetch date range');
}

/**
 * Fetches orders, trades and transactions between two dates (yyyy-MM-dd, UTC, inclusive)
 * into the "Range …" sheets, replacing what was there. Independent of Fetch / continue.
 */
function fetchDateRange(fromDate, toDate) {
  try {
    var start = new Date(fromDate + 'T00:00:00Z').getTime();
    var end = Math.min(new Date(toDate + 'T23:59:59.999Z').getTime(), Date.now());
    if (isNaN(start) || isNaN(end) || start > end) throw new Error('Invalid range: ' + fromDate + ' → ' + toDate);
    var cfg = loadConfig_();
    log_('Range fetch ' + fromDate + ' → ' + toDate + ' — user ' + mask_(cfg.user));

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    Object.keys(RANGE_SHEETS_).forEach(function (k) {
      var sh = ss.getSheetByName(RANGE_SHEETS_[k]);
      if (sh) sh.clearContents();
    });

    var started = Date.now(), total = { orders: 0, trades: 0, income: 0 };
    for (var ws = start; ws <= end; ws += WEEK_MS + 1) {
      if (Date.now() - started > RUN_BUDGET_MS) {
        throw new Error('Range too large to finish in one run — stopped at ' + day_(ws) +
          '. Pick a shorter range (data up to that date is in the sheets).');
      }
      fetchWeek_(ws, Math.min(ws + WEEK_MS, end), total, RANGE_SHEETS_);
    }
    Object.keys(RANGE_SHEETS_).forEach(function (k) { sortNewestFirst_(RANGE_SHEETS_[k]); });
    var msg = 'Range ' + fromDate + ' → ' + toDate + ' done: ' + summary_(total);
    log_(msg);
    return msg;
  } catch (e) {
    log_('ERROR: ' + e.message, 'ERROR');
    throw e;
  }
}

function pause_(where, total) {
  ScriptApp.newTrigger('fetchAll').timeBased().after(60 * 1000).create();
  log_('Paused (' + where + ', this run: ' + summary_(total) + ') — continuing automatically in ~1 min');
}

function resetAll() {
  clearTriggers_();
  var props = PropertiesService.getScriptProperties();
  [HEAD_KEY, TAIL_KEY, EMPTY_KEY, DONE_KEY, 'ASTER_CURSOR_MS'].forEach(function (k) { props.deleteProperty(k); });
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ['Orders', 'Trades', 'Transactions'].forEach(function (n) {
    var sh = ss.getSheetByName(n);
    if (sh) sh.clearContents();
  });
  log_('Reset done — run "Fetch / continue" to start again from the newest data');
}

/** Validates Script Properties and logs which ones are set (never logs the private key). */
function checkConfig() {
  try {
    var cfg = loadConfig_();
    log_('Config OK — user ' + mask_(cfg.user) + ', signer ' + mask_(cfg.signer) + ', private key set');
  } catch (e) {
    log_('ERROR: ' + e.message, 'ERROR');
  }
}

/** Reads ASTER_USER / ASTER_SIGNER / ASTER_SIGNER_PRIVATE_KEY from Script Properties. */
var config_ = null;
function loadConfig_() {
  if (config_) return config_;
  var props = PropertiesService.getScriptProperties();
  var cfg = {
    user: (props.getProperty('ASTER_USER') || '').trim(),
    signer: (props.getProperty('ASTER_SIGNER') || '').trim(),
    pk: (props.getProperty('ASTER_SIGNER_PRIVATE_KEY') || '').trim()
  };
  var missing = ['ASTER_USER', 'ASTER_SIGNER', 'ASTER_SIGNER_PRIVATE_KEY']
    .filter(function (k, i) { return !cfg[['user', 'signer', 'pk'][i]]; });
  if (missing.length) throw new Error('Missing Script Properties: ' + missing.join(', '));
  if (!/^0x[0-9a-fA-F]{40}$/.test(cfg.user)) throw new Error('ASTER_USER is not a valid address');
  if (!/^0x[0-9a-fA-F]{40}$/.test(cfg.signer)) throw new Error('ASTER_SIGNER is not a valid address');
  if (!/^(0x)?[0-9a-fA-F]{64}$/.test(cfg.pk)) throw new Error('ASTER_SIGNER_PRIVATE_KEY must be 64 hex chars');
  if (cfg.user.toLowerCase() === cfg.signer.toLowerCase()) {
    throw new Error('ASTER_SIGNER must be the API wallet, not the main wallet');
  }
  return (config_ = cfg);
}

function mask_(addr) { return addr.slice(0, 6) + '…' + addr.slice(-4); }
function day_(ms) { return new Date(ms).toISOString().slice(0, 10); }
function summary_(t) { return t.orders + ' orders, ' + t.trades + ' trades, ' + t.income + ' transactions'; }

function clearTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'fetchAll') ScriptApp.deleteTrigger(t);
  });
}

/**
 * Logs to Apps Script Executions (console) and shows a toast. No log sheet.
 * Pass quiet=true to skip the toast (for per-week progress lines).
 */
function log_(msg, level, quiet) {
  level = level || 'INFO';
  (level === 'ERROR' ? console.error : level === 'WARN' ? console.warn : console.log)(msg);
  if (quiet) return;
  try {
    SpreadsheetApp.getActiveSpreadsheet()
      .toast(msg, level === 'ERROR' ? 'Aster ⚠️' : 'Aster', level === 'ERROR' ? 30 : 10);
  } catch (e) {}
}

// ---------- HTTP ----------

function signedGet_(path, params) {
  var cfg = loadConfig_(), user = cfg.user, signer = cfg.signer, pk = cfg.pk;

  var p = {};
  Object.keys(params).forEach(function (k) { if (params[k] != null) p[k] = params[k]; });
  p.user = user; p.signer = signer; p.nonce = nextNonce_();
  var query = Object.keys(p).map(function (k) {
    return encodeURIComponent(k) + '=' + encodeURIComponent(p[k]);
  }).join('&');

  var url = BASE_URL + path + '?' + query + '&signature=' + signAster_(query, pk);
  var res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
  if (res.getResponseCode() === 429) {  // rate limited: back off once, then retry with a fresh nonce
    log_('Rate limited (429) on ' + path + ' — waiting 61s', 'WARN');
    Utilities.sleep(61000);
    return signedGet_(path, params);
  }
  // Stay under the per-IP weight limit (2400/min): pause when the minute's budget is nearly used.
  var used = Number(res.getHeaders()['X-MBX-USED-WEIGHT-1M'] || res.getHeaders()['x-mbx-used-weight-1m'] || 0);
  if (used > 1800) {
    console.log('Weight ' + used + '/2400 — pausing until next minute');
    Utilities.sleep(61000 - (Date.now() % 60000));
  }
  if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode() + ': ' + res.getContentText());
  return JSON.parse(res.getContentText());
}

var lastNonce_ = 0;
function nextNonce_() {
  // Microsecond timestamp, strictly increasing (server rejects reused nonces).
  lastNonce_ = Math.max(Date.now() * 1000, lastNonce_ + 1);
  return lastNonce_;
}

/** One window; if it hits the 1000-row cap, split it in half so nothing is dropped. */
function fetchRange_(path, params, startMs, endMs) {
  var q = Object.assign({}, params, { startTime: startMs, endTime: endMs, limit: 1000 });
  var rows = signedGet_(path, q);
  if (rows.length < 1000 || endMs - startMs < 1000) return rows;
  var mid = Math.floor((startMs + endMs) / 2);
  return fetchRange_(path, params, startMs, mid).concat(fetchRange_(path, params, mid + 1, endMs));
}

// ---------- Sheets ----------

var TIME_KEYS_ = { time: true, updateTime: true };

/** Appends rows under the existing header, adding new columns if the API returns new fields. */
function appendRows_(name, rows) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (!rows.length) return;
  var header = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
  rows.forEach(function (r) { Object.keys(r).forEach(function (k) { if (header.indexOf(k) < 0) header.push(k); }); });
  sh.getRange(1, 1, 1, header.length).setValues([header]);
  sh.setFrozenRows(1);
  var values = rows.map(function (r) {
    return header.map(function (k) {
      var v = r[k];
      if (TIME_KEYS_[k] && typeof v === 'number') return new Date(v);
      return v == null ? '' : v;
    });
  });
  sh.getRange(sh.getLastRow() + 1, 1, values.length, header.length).setValues(values);
}

// BigInt constants (Apps Script rejects 123n literals).
var B_0_ = BigInt(0);
var B_xff_ = BigInt(0xff);
var B_1_ = BigInt(1);
var B_1666_ = BigInt(1666);
var B_2_ = BigInt(2);
var B_3_ = BigInt(3);
var B_64_ = BigInt(64);
var B_8_ = BigInt(8);

/** Sorts data rows by the `time` column, newest on top. */
function sortNewestFirst_(name) {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
    if (!sh || sh.getLastRow() < 3) return;
    var col = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].indexOf('time') + 1;
    if (col) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).sort({ column: col, ascending: false });
  } catch (e) {
    console.error('Sort ' + name + ' failed: ' + e.message);
  }
}

// ---------- EIP-712 signing (Aster V3) ----------

/** Signs the urlencoded query as EIP-712 Message(string msg) under the AsterSignTransaction domain. */
function signAster_(msg, privateKeyHex) {
  var domainType = keccak256_(utf8_('EIP712Domain(string name,string version,uint256 chainId,address verifyingContract)'));
  var domainSep = keccak256_(concat_([domainType, keccak256_(utf8_('AsterSignTransaction')),
    keccak256_(utf8_('1')), uint256_(B_1666_), new Uint8Array(32)]));
  var structHash = keccak256_(concat_([keccak256_(utf8_('Message(string msg)')), keccak256_(utf8_(msg))]));
  var digest = keccak256_(concat_([new Uint8Array([0x19, 0x01]), domainSep, structHash]));
  return ecdsaSign_(digest, BigInt('0x' + privateKeyHex.replace(/^0x/, '')));
}

function utf8_(s) {
  var e = unescape(encodeURIComponent(s)), out = new Uint8Array(e.length);
  for (var i = 0; i < e.length; i++) out[i] = e.charCodeAt(i);
  return out;
}
function concat_(arrs) {
  var n = arrs.reduce(function (a, b) { return a + b.length; }, 0), out = new Uint8Array(n), o = 0;
  arrs.forEach(function (a) { out.set(a, o); o += a.length; });
  return out;
}
function uint256_(x) {
  var out = new Uint8Array(32);
  for (var i = 31; i >= 0; i--) { out[i] = Number(x & B_xff_); x >>= B_8_; }
  return out;
}
function toBig_(bytes) {
  var h = '0x';
  for (var i = 0; i < bytes.length; i++) h += ('0' + bytes[i].toString(16)).slice(-2);
  return BigInt(h);
}
function hex_(bytes) {
  var h = '';
  for (var i = 0; i < bytes.length; i++) h += ('0' + bytes[i].toString(16)).slice(-2);
  return h;
}

// Keccak-256 (Ethereum variant, 0x01 padding), 64-bit lanes as BigInt.
var KECCAK_RC_ = ['0x1', '0x8082', '0x800000000000808a', '0x8000000080008000', '0x808b', '0x80000001',
  '0x8000000080008081', '0x8000000000008009', '0x8a', '0x88', '0x80008009', '0x8000000a', '0x8000808b',
  '0x800000000000008b', '0x8000000000008089', '0x8000000000008003', '0x8000000000008002',
  '0x8000000000000080', '0x800a', '0x800000008000000a', '0x8000000080008081', '0x8000000000008080',
  '0x80000001', '0x8000000080008008'].map(BigInt);
var KECCAK_ROT_ = [0, 1, 62, 28, 27, 36, 44, 6, 55, 20, 3, 10, 43, 25, 39, 41, 45, 15, 21, 8, 18, 2, 61, 56, 14];
var M64_ = (B_1_ << B_64_) - B_1_;

function keccak256_(data) {
  var rate = 136, padded = new Uint8Array(Math.ceil((data.length + 1) / rate) * rate);
  padded.set(data); padded[data.length] ^= 0x01; padded[padded.length - 1] ^= 0x80;
  var s = new Array(25).fill(B_0_);
  for (var off = 0; off < padded.length; off += rate) {
    for (var i = 0; i < rate / 8; i++) {
      var lane = B_0_;
      for (var b = 7; b >= 0; b--) lane = (lane << B_8_) | BigInt(padded[off + i * 8 + b]);
      s[i] ^= lane;
    }
    keccakF_(s);
  }
  var out = new Uint8Array(32);
  for (var j = 0; j < 32; j++) out[j] = Number((s[j >> 3] >> BigInt((j & 7) * 8)) & B_xff_);
  return out;
}
function rotl_(x, n) { n = BigInt(n); return n === B_0_ ? x : ((x << n) | (x >> (B_64_ - n))) & M64_; }
function keccakF_(s) {
  for (var r = 0; r < 24; r++) {
    var c = [], d = [], x, y;
    for (x = 0; x < 5; x++) c[x] = s[x] ^ s[x + 5] ^ s[x + 10] ^ s[x + 15] ^ s[x + 20];
    for (x = 0; x < 5; x++) d[x] = c[(x + 4) % 5] ^ rotl_(c[(x + 1) % 5], 1);
    for (var i = 0; i < 25; i++) s[i] ^= d[i % 5];
    var bArr = new Array(25);
    for (x = 0; x < 5; x++) for (y = 0; y < 5; y++)
      bArr[y + 5 * ((2 * x + 3 * y) % 5)] = rotl_(s[x + 5 * y], KECCAK_ROT_[x + 5 * y]);
    for (x = 0; x < 5; x++) for (y = 0; y < 5; y++)
      s[x + 5 * y] = bArr[x + 5 * y] ^ ((~bArr[(x + 1) % 5 + 5 * y] & M64_) & bArr[(x + 2) % 5 + 5 * y]);
    s[0] ^= KECCAK_RC_[r];
  }
}

// secp256k1 ECDSA with RFC 6979 deterministic k; returns 0x r||s||v (v = 27/28), low-s.
var P_ = BigInt('0xfffffffffffffffffffffffffffffffffffffffffffffffffffffffefffffc2f');
var N_ = BigInt('0xfffffffffffffffffffffffffffffffebaaedce6af48a03bbfd25e8cd0364141');
var G_ = [BigInt('0x79be667ef9dcbbac55a06295ce870b07029bfcdb2dce28d959f2815b16f81798'),
  BigInt('0x483ada7726a3c4655da4fbfc0e1108a8fd17b448a68554199c47d08ffb10d4b8')];

function mod_(a, m) { var r = a % m; return r < B_0_ ? r + m : r; }
function inv_(a, m) {
  var lm = B_1_, hm = B_0_, low = mod_(a, m), high = m;
  while (low > B_1_) { var q = high / low; var nm = hm - lm * q, nw = high - low * q; hm = lm; lm = nm; high = low; low = nw; }
  return mod_(lm, m);
}
function ptAdd_(p, q) {
  if (!p) return q; if (!q) return p;
  var l;
  if (p[0] === q[0]) {
    if (mod_(p[1] + q[1], P_) === B_0_) return null;
    l = mod_(B_3_ * p[0] * p[0] * inv_(B_2_ * p[1], P_), P_);
  } else {
    l = mod_((q[1] - p[1]) * inv_(q[0] - p[0], P_), P_);
  }
  var x = mod_(l * l - p[0] - q[0], P_);
  return [x, mod_(l * (p[0] - x) - p[1], P_)];
}
function ptMul_(k, p) {
  var r = null;
  while (k > B_0_) { if (k & B_1_) r = ptAdd_(r, p); p = ptAdd_(p, p); k >>= B_1_; }
  return r;
}
function hmac_(key, msg) {
  return new Uint8Array(Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256,
    Array.from(msg).map(function (b) { return b > 127 ? b - 256 : b; }),
    Array.from(key).map(function (b) { return b > 127 ? b - 256 : b; })).map(function (b) { return b & 0xff; }));
}
function rfc6979K_(digest, d) {
  var x = uint256_(d), h = uint256_(mod_(toBig_(digest), N_));
  var v = new Uint8Array(32).fill(1), k = new Uint8Array(32);
  k = hmac_(k, concat_([v, new Uint8Array([0]), x, h])); v = hmac_(k, v);
  k = hmac_(k, concat_([v, new Uint8Array([1]), x, h])); v = hmac_(k, v);
  for (;;) {
    v = hmac_(k, v);
    var t = toBig_(v);
    if (t > B_0_ && t < N_) return t;
    k = hmac_(k, concat_([v, new Uint8Array([0])])); v = hmac_(k, v);
  }
}
function ecdsaSign_(digest, d) {
  var z = toBig_(digest), k = rfc6979K_(digest, d);
  var R = ptMul_(k, G_), r = mod_(R[0], N_);
  var s = mod_(inv_(k, N_) * (z + r * d), N_), recId = Number(R[1] & B_1_);
  if (s > N_ / B_2_) { s = N_ - s; recId ^= 1; }
  return '0x' + hex_(uint256_(r)) + hex_(uint256_(s)) + (27 + recId).toString(16);
}
