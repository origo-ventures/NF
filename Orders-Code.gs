/**
 * NICE FASHION - order sheet (Google Apps Script)  v6.1
 * Saves every order / inquiry sent from the website into a Google Sheet.
 *
 * SETUP (same as before):
 *  1) Create a Google Sheet (any name). Extensions > Apps Script. Paste this file, Save.
 *  2) Optional: set NOTIFY_EMAIL below to get an email for every new order.
 *  3) Deploy > New deployment > type "Web app": Execute as "Me", Who has access "Anyone". Deploy and copy the /exec URL.
 *  4) Control panel > Settings > "Order sheet link": paste the /exec URL, then Save & Publish.
 *  UPDATING from v5/v6.0: paste this file over the old code, Save, then Deploy > Manage deployments > Edit (pencil)
 *  > Version: "New version" > Deploy. The /exec URL stays the same.
 *
 * SECURITY NOTES (v6.1)
 *  - The web app must be public ("Anyone") because customers' browsers post to it. Anything in the website's
 *    JavaScript is public, so there is NO secret here. Every request is therefore validated.
 *  - Prices, discounts and totals sent by the browser are NOT trusted. They are saved for reference only.
 *    The "Verification" column tells staff whether the items' subtotal matched the real prices.
 *    Staff must confirm the amount (and any payment) before accepting an order. Status always starts as "New".
 *  - "Payment" and "TrxID" are what the CUSTOMER typed. They do not mean the order is paid.
 *  - The limits below are best-effort only (Apps Script has no real rate limiting). Do not rely on them against a determined attacker.
 *  - Share the spreadsheet only with people who need it. It contains customer names, phone numbers and addresses.
 */
var NOTIFY_EMAIL = '';        // e.g. 'you@gmail.com' (leave empty for no email)
var SHEET_NAME   = 'Orders';
var PRODUCTS_URL = '';        // OPTIONAL: your site's product file, e.g. 'https://YOUR-SITE/data/products.js'. Enables the subtotal check.
var MAX_PER_MIN  = 60;        // best-effort flood limit for the whole site
var MAX_PER_PHONE = 5;        // best-effort: max submissions per phone number per 10 minutes
var MAX_EMAILS_PER_HOUR = 30; // best-effort: stops a flood from filling your inbox / email quota

var MAX_BODY = 20000, MAX_ITEMS = 50, MAX_QTY = 100, MAX_MONEY = 100000000;
var MODES = ['order', 'availability', 'inquiry'];
var PAYMENTS = ['', 'cod', 'bkash', 'nagad', 'qr'];

function doPost(e) {
  var lock = null;
  try {
    if (!e || !e.postData || !e.postData.contents || e.postData.contents.length > MAX_BODY) return out('error');
    var d;
    try { d = JSON.parse(e.postData.contents); } catch (x) { return out('error'); }
    var r = validate(d);
    if (!r) return out('error');

    lock = LockService.getScriptLock();
    lock.waitLock(10000);
    if (overLimit('all', MAX_PER_MIN, 60)) return out('busy');
    if (overLimit('p' + r.phoneDigits, MAX_PER_PHONE, 600)) return out('busy');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) {
      sh.appendRow(['Time', 'Ref', 'Type', 'Name', 'Phone', 'Address', 'Items', 'Subtotal', 'Coupon', 'Discount', 'Delivery', 'Total', 'Payment', 'TrxID', 'Status', 'Verification']);
      sh.setFrozenRows(1);
    } else if (sh.getRange(1, 16).getValue() === '') {
      sh.getRange(1, 16).setValue('Verification');
    }
    var check = verify(r);
    sh.appendRow([new Date(), safe(r.ref), safe(r.mode), safe(r.name), "'" + r.phone, safe(r.address), safe(r.itemsText),
      r.subtotal, safe(r.coupon), r.discount, r.delivery, r.total, safe(r.payment), safe(r.trx), 'New', check]);

    if (NOTIFY_EMAIL && !overLimit('mail', MAX_EMAILS_PER_HOUR, 3600)) {
      try {
        MailApp.sendEmail(NOTIFY_EMAIL, 'New ' + r.mode + ' ' + r.ref,
          'Name: ' + r.name + '\nPhone: ' + r.phone + '\n\n' + r.itemsText + '\n\nTotal claimed by browser (NOT verified): ' + r.total + '\nCheck: ' + check);
      } catch (x) { /* email problems must never lose the order */ }
    }
    return out('ok');
  } catch (err) {
    return out('error'); // no details are returned to the public
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (x) {} }
  }
}

function out(s) { return ContentService.createTextOutput(s); }

/* ---- validation: returns a clean record, or null if the request is not acceptable ---- */
function validate(d) {
  if (!d || typeof d !== 'object' || Array.isArray(d)) return null;
  if (MODES.indexOf(d.mode) < 0) return null;
  if (typeof d.ref !== 'string' || !/^NF-\d{4}-\d{3}$/.test(d.ref)) return null;
  var name = text(d.name, 100);
  if (!name) return null;
  var phone = normPhone(d.phone);
  if (!phone) return null;
  if (!Array.isArray(d.items) || d.items.length < 1 || d.items.length > MAX_ITEMS) return null;
  var items = [], lines = [];
  for (var i = 0; i < d.items.length; i++) {
    var it = d.items[i];
    if (!it || typeof it !== 'object' || Array.isArray(it)) return null;
    if (typeof it.id !== 'string' || !/^[A-Za-z0-9][A-Za-z0-9_-]{0,60}$/.test(it.id)) return null;
    var q = Number(it.qty == null ? 1 : it.qty);
    if (!isFinite(q) || q !== Math.floor(q) || q < 1 || q > MAX_QTY) return null;
    var price = it.price == null ? null : Number(it.price);
    if (price !== null && (!isFinite(price) || price < 0 || price > MAX_MONEY)) return null;
    var nm = text(it.name, 200), sz = text(it.size, 30), co = text(it.color, 50);
    items.push({ id: it.id, qty: q, price: price });
    lines.push(nm + ' (' + it.id + ') ' + sz + (co ? '/' + co : '') + ' x' + q);
  }
  var money = {};
  var keys = ['subtotal', 'discount', 'delivery', 'total'];
  for (var k = 0; k < keys.length; k++) {
    var v = d[keys[k]];
    if (v == null || v === '') { money[keys[k]] = 0; continue; }
    v = Number(v);
    if (!isFinite(v) || v < 0 || v > MAX_MONEY) return null;
    money[keys[k]] = v;
  }
  var payment = d.payment == null ? '' : String(d.payment);
  if (PAYMENTS.indexOf(payment) < 0) return null;
  return {
    mode: d.mode, ref: d.ref, name: name, phone: phone, phoneDigits: phone.replace(/\D/g, ''),
    address: text(d.address, 300), items: items, itemsText: lines.join('\n'),
    subtotal: money.subtotal, discount: money.discount, delivery: money.delivery, total: money.total,
    coupon: text(d.coupon, 40), payment: payment, trx: text(d.trx, 40)
  };
}

/* plain text: strips control characters, trims, limits length */
function text(v, max) {
  if (v == null) return '';
  if (typeof v !== 'string' && typeof v !== 'number') return '';
  return String(v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F\u2028\u2029]/g, '').replace(/\r\n?/g, '\n').trim().slice(0, max);
}
function normPhone(v) {
  if (typeof v !== 'string' && typeof v !== 'number') return '';
  var s = String(v).trim();
  var plus = s.charAt(0) === '+' ? '+' : '';
  var digits = s.replace(/\D/g, '');
  if (digits.length < 10 || digits.length > 15) return '';
  return plus + digits;
}
/* Formula-injection guard: a cell that STARTS with = + - @ (after spaces/tabs) could run as a formula.
   A leading apostrophe makes Sheets keep it as plain text (the apostrophe is not shown). */
function safe(v) {
  var s = String(v == null ? '' : v);
  return /^[\s\u00A0]*[=+\-@]/.test(s) || /^[\t\r]/.test(s) ? "'" + s : s;
}

/* ---- best-effort counters (CacheService is shared and not atomic; this is NOT real rate limiting) ---- */
function overLimit(key, max, seconds) {
  var cache = CacheService.getScriptCache();
  var k = 'c:' + key + ':' + Math.floor(Date.now() / (seconds * 1000));
  var n = Number(cache.get(k) || 0) + 1;
  cache.put(k, String(n), Math.min(21600, seconds * 2));
  return n > max;
}

/* ---- subtotal check against the real product list (only if PRODUCTS_URL is set) ---- */
function verify(r) {
  var base = 'NOT VERIFIED - confirm the amount with the customer before accepting';
  if (!PRODUCTS_URL) return base;
  var prices = productPrices();
  if (!prices) return base + ' (price list unavailable)';
  var sum = 0;
  for (var i = 0; i < r.items.length; i++) {
    var p = prices[r.items[i].id];
    if (p === undefined) return 'CHECK: unknown product id ' + r.items[i].id;
    if (p === 0) return 'CHECK: price on request - confirm manually';
    sum += p * r.items[i].qty;
  }
  if (sum !== r.subtotal) return 'MISMATCH: real subtotal is ' + sum + ' but browser sent ' + r.subtotal;
  return 'Subtotal matches. Discount, delivery and payment are NOT verified.';
}
function productPrices() {
  try {
    var cache = CacheService.getScriptCache(), raw = cache.get('products');
    if (!raw) {
      if (!/^https:\/\//i.test(PRODUCTS_URL)) return null;
      raw = UrlFetchApp.fetch(PRODUCTS_URL, { muteHttpExceptions: false, followRedirects: true }).getContentText();
      if (raw.length < 90000) cache.put('products', raw, 600);
    }
    var i = raw.indexOf('PRODUCTS=');
    if (i < 0) return null;
    i = raw.indexOf('[', i);
    var depth = 0, inStr = false, esc = false, j;
    for (j = i; j < raw.length; j++) {
      var c = raw.charAt(j);
      if (inStr) { if (esc) esc = false; else if (c === '\\') esc = true; else if (c === '"') inStr = false; continue; }
      if (c === '"') { inStr = true; continue; }
      if (c === '[') depth++;
      else if (c === ']') { depth--; if (depth === 0) break; }
    }
    var list = JSON.parse(raw.slice(i, j + 1)), map = {};
    for (var k = 0; k < list.length; k++) map[list[k].id] = Number(list[k].price) || 0;
    return map;
  } catch (x) { return null; }
}
