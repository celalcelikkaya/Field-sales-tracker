/**
 * Sales Dashboard (multi-company sales tracking)
 * Google Apps Script web app backed by a Google Sheet.
 * Works both when bound to a Sheet and as a standalone script.
 *
 *  - Salesmen log in with their PIN and record what they sold.
 *  - Admin logs in with the Admin PIN and sees the main performance dashboard.
 *
 * First run: select `install` in the toolbar and press Run (creates the spreadsheet and all sheets).
 */

const CFG = {
  get TZ() { return tz_(); },   // read from the spreadsheet time zone (set in setup)
  SESSION_SECONDS: 21600,          // 6 hours (CacheService maximum)
  MAX_FAILS: 30,                   // failed logins allowed per 15 minutes
  SHEET: { SALESMEN: 'Salesmen', PRODUCTS: 'Products', SALES: 'Sales', SETTINGS: 'Settings' }
};

/** Developer contact shown in the admin console (Settings tab). Put the real support email in `email`. */
const DEVELOPER = { name: 'Celal Çelikkaya', site: 'https://celalcelikkaya.github.io', email: 'celalmalawi@gmail.com' };

/**
 * LICENSING: every install reports to your license server and gets a 14-day trial.
 * `url` is the deployed LicenseServer web app (already set). Empty = licensing off (development).
 */
const LICENSE = { url: 'https://script.google.com/macros/s/AKfycbxPVT7duxUVx79jOnNH1-llYpyKQWngAttjG_wW0iWA_dKjdsL6T_3ChX-nmtVCbpSHtg/exec', recheckHours: 6 };

const SALES_HEADERS = ['SaleID', 'Timestamp', 'Date', 'SalesmanID', 'Salesman', 'ProductCode', 'Product',
  'Category', 'Qty', 'UnitPrice', 'LineTotal', 'Customer', 'Location', 'Payment', 'Notes', 'Status',
  'VoidedBy', 'VoidedAt', 'VoidReason', 'CustomerID', 'Replaces'];

const VOID_REASONS = ['Wrong quantity', 'Wrong product', 'Wrong customer', 'Wrong payment type',
  'Duplicate entry', 'Customer returned goods', 'Customer cancelled order'];

const PAYMENTS = ['Cash', 'Airtel Money', 'TNM Mpamba', 'Bank', 'Credit'];

/* ======================================================================
 * SPREADSHEET ACCESS (works bound or standalone)
 * ==================================================================== */

function ss_(createIfMissing) {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  if (!createIfMissing) throw new Error('No spreadsheet linked yet - run install() in the Apps Script editor first');
  const ss = SpreadsheetApp.create('Sales Dashboard Data');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" missing - run install() first');
  return sh;
}

/* ======================================================================
 * SETUP & MENU
 * ==================================================================== */

/** Only does something when the script is bound to a Sheet. Safe to ignore otherwise. */
function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('Sales Dashboard')
      .addItem('Install (create sheets)', 'install')
      .addItem('Show dashboard link', 'showLink_')
      .addToUi();
  } catch (e) { /* standalone script: no UI */ }
}

function setup() { install(); }

/**
 * ONE-TIME step for whoever deploys the app: run `install` in the editor.
 * It creates the spreadsheet + empty sheets and prints a SETUP KEY in the
 * Execution log. The first person to open the web app enters that key to
 * start the setup wizard (company, logo, admin PIN, products).
 */
function install() {
  const ss = ss_(true);
  if (ss.getSpreadsheetTimeZone() === 'Etc/GMT' || !ss.getSpreadsheetTimeZone()) ss.setSpreadsheetTimeZone('Africa/Blantyre');

  const setSh = ss.getSheetByName(CFG.SHEET.SETTINGS) || ss.insertSheet(CFG.SHEET.SETTINGS);
  setSh.getRange('B:B').setNumberFormat('@');   // before seeding, so 0000 stays text
  makeSheet_(ss, CFG.SHEET.SETTINGS, ['Key', 'Value'], [
    ['CompanyName', ''], ['Tagline', ''], ['Currency', 'MK'], ['Logo', ''], ['AdminPIN', '0000'], ['SetupDone', 'no']
  ]);
  makeSheet_(ss, CFG.SHEET.SALESMEN, ['SalesmanID', 'Name', 'PIN', 'Phone', 'Area', 'MonthlyTarget', 'Active'], []);
  makeSheet_(ss, CFG.SHEET.PRODUCTS, ['ProductCode', 'Name', 'Category', 'Unit', 'Price', 'Active'], []);
  makeSheet_(ss, CFG.SHEET.SALES, SALES_HEADERS, []);
  ensureSalesHeaders_();
  tab_('Customers', CUSTOMER_HEADERS);
  tab_('Payments', PAYMENT_HEADERS);
  tab_('Allocations', ALLOC_HEADERS);
  ledgerSh_();

  // Keep PINs and dates as plain text (so "0123" keeps its leading zero).
  ss.getSheetByName(CFG.SHEET.SALESMEN).getRange('C:C').setNumberFormat('@');
  ss.getSheetByName(CFG.SHEET.SETTINGS).getRange('B:B').setNumberFormat('@');
  ss.getSheetByName(CFG.SHEET.SALES).getRange('C:C').setNumberFormat('@');
  ss.getSheetByName(CFG.SHEET.SALES).getRange('J:K').setNumberFormat('#,##0');

  const sheet1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('Sayfa1');
  if (sheet1 && sheet1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sheet1);

  const props = PropertiesService.getScriptProperties();
  if (!isSetupDone_()) {
    let key = props.getProperty('SETUP_KEY');
    if (!key) { key = String(Math.floor(100000 + Math.random() * 900000)); props.setProperty('SETUP_KEY', key); }
    Logger.log('SETUP KEY: ' + key + '   (enter it on the first screen of the web app)');
  } else {
    Logger.log('Setup already completed.');
  }
  Logger.log('Spreadsheet: ' + ss.getUrl());
  clearTableCache_();
  const lic = licenseInfo_(); Logger.log('License: ' + lic.status + (lic.daysLeft != null ? ' (' + lic.daysLeft + ' days left)' : ''));
  const ls = licState_(); if (LICENSE.url) Logger.log('License server reached: ' + (ls && !ls.err ? 'YES' : 'NO - ' + (ls && ls.err)));
}

let _tz;
function tz_() {
  if (!_tz) { try { _tz = ss_().getSpreadsheetTimeZone(); } catch (e) {} }
  return _tz || 'Africa/Blantyre';
}
function validTz_(tz) {
  tz = String(tz || '').trim() || 'Africa/Blantyre';
  try { Utilities.formatDate(new Date(), tz, 'yyyy'); } catch (e) { throw new Error('Invalid time zone'); }
  return tz;
}
function validLogo_(v) {
  v = String(v || '');
  if (!v) return '';
  if (!/^data:image\/(png|jpeg|webp);base64,[A-Za-z0-9+\/=]+$/.test(v)) throw new Error('Logo must be a PNG, JPG or WEBP image');
  if (v.length > 45000) throw new Error('Logo is too large. Use a smaller image.');
  return v;
}
/** Admin PIN as text. Sheets can store 0000 as the number 0, so digit strings shorter than 4 are zero-padded back. */
function adminPin_(s) {
  s = s || settings_();
  let v = s.AdminPIN;
  if (v === undefined || v === null) return '';
  v = String(v).trim();
  if (/^\d{1,3}$/.test(v)) v = ('0000' + v).slice(-4);
  return v;
}
function isSetupDone_() {
  try {
    const s = settings_();
    if (s.SetupDone === undefined) return !!adminPin_(s);   // older installs
    return String(s.SetupDone).trim().toLowerCase() === 'yes';
  } catch (e) { return false; }
}
function unitsByCat_(rows) {
  const o = {};
  rows.forEach(r => { o[r.Category] = (o[r.Category] || 0) + r.Qty; });
  return o;
}
function setSetting_(key, value) {
  const sh = sheet_(CFG.SHEET.SETTINGS);
  const values = sh.getDataRange().getValues();
  let row = 0;
  for (let i = 1; i < values.length; i++) if (String(values[i][0]).trim() === key) row = i + 1;
  if (!row) { row = sh.getLastRow() + 1; sh.getRange(row, 1).setValue(key); }
  sh.getRange(row, 2).setNumberFormat('@').setValue(value);
  clearTableCache_();
}

/* ======================================================================
 * FIRST-RUN SETUP (product onboarding) & COMPANY SETTINGS
 * ==================================================================== */

/** Public: tells the screen what to show (not installed / wizard / login) and the branding. */
function getBootState() {
  try { sheet_(CFG.SHEET.SETTINGS); } catch (e) { return { installed: false }; }
  const s = settings_();
  return { installed: true, setupDone: isSetupDone_(), company: s.CompanyName || '', tagline: s.Tagline || '',
    currency: s.Currency || 'MK', logo: s.Logo || '', timezone: tz_(),
    developer: { name: DEVELOPER.name, site: DEVELOPER.site, email: DEVELOPER.email, canEmail: !!DEVELOPER.email },
    license: licenseInfo_() };
}

/**
 * Wizard submit. d = { company, tagline, currency, timezone, logo, pin,
 *   products: [{ name, category, unit, price }] }. Logs the new admin in.
 */
function completeSetup(setupKey, d) {
  d = d || {};
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('FAILS') || 0);
  if (fails >= CFG.MAX_FAILS) throw new Error('Too many wrong attempts. Wait 15 minutes.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let pin;
  try {
    if (isSetupDone_()) throw new Error('Setup is already completed');
    const props = PropertiesService.getScriptProperties();
    const key = props.getProperty('SETUP_KEY');
    if (!key || String(setupKey || '').trim() !== key) {
      cache.put('FAILS', String(fails + 1), 900);
      throw new Error('Wrong setup key. Run install() in the editor and read the Execution log.');
    }

    const company = clean_(d.company, 80);
    if (!company) throw new Error('Company name is required');
    pin = adminPin_() || '0000';   // default Admin PIN is 0000; the admin changes it in Settings
    const tz = validTz_(d.timezone);
    const logo = validLogo_(d.logo);

    const items = (d.products || []).slice(0, 200).map(p => ({
      name: clean_(p.name, 80), category: clean_(p.category, 30), unit: clean_(p.unit, 30), price: Number(p.price)
    })).filter(p => p.name);
    if (!items.length) throw new Error('Add at least one product');
    items.forEach(p => {
      if (!p.category) throw new Error('Category is required for ' + p.name);
      if (!(p.price >= 0) || p.price > 100000000) throw new Error('Enter a valid price for ' + p.name);
    });

    ss_().setSpreadsheetTimeZone(tz); _tz = tz;
    setSetting_('CompanyName', company);
    setSetting_('Tagline', clean_(d.tagline, 80));
    setSetting_('Currency', clean_(d.currency, 8) || 'MK');
    setSetting_('Logo', logo);

    const sh = sheet_(CFG.SHEET.PRODUCTS);
    const rows = items.map((p, i) => ['P' + ('00' + (i + 1)).slice(-3), p.name, p.category, p.unit, p.price, 'Yes']);
    sh.getRange(2, 1, rows.length, 6).setValues(rows);

    setSetting_('AdminPIN', pin);
    setSetting_('SetupDone', 'yes');       // last: marks setup as done
    props.deleteProperty('SETUP_KEY');
  } finally {
    unlock_(lock);
  }
  try { licenseCheck_(true); } catch (e) {}   // report the company name to the license server
  return login(pin);
}

function adminSaveCompany(token, d) {
  session_(token, 'admin');
  d = d || {};
  const company = clean_(d.company, 80);
  if (!company) throw new Error('Company name is required');
  const logo = d.logo === undefined ? undefined : validLogo_(d.logo);   // undefined = keep, '' = remove
  const tz = d.timezone ? validTz_(d.timezone) : null;

  setSetting_('CompanyName', company);
  setSetting_('Tagline', clean_(d.tagline, 80));
  setSetting_('Currency', clean_(d.currency, 8) || 'MK');
  if (logo !== undefined) setSetting_('Logo', logo);
  if (tz) { ss_().setSpreadsheetTimeZone(tz); _tz = tz; }
  return getBootState();
}

/* ======================================================================
 * SETUP HELPERS
 * ==================================================================== */

function makeSheet_(ss, name, headers, seed) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (seed.length) sh.getRange(2, 1, seed.length, headers.length).setValues(seed);
  }
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#0b4f8a').setFontColor('#ffffff');
  return sh;
}

function showLink_() {
  const url = ScriptApp.getService().getUrl();
  const msg = url ? url : 'Not deployed yet. Use Deploy > New deployment > Web app.';
  SpreadsheetApp.getUi().alert('Dashboard link', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

/* ======================================================================
 * WEB APP
 * ==================================================================== */

function doGet() {
  let title = 'Sales Dashboard';
  try { title = String(settings_().CompanyName || '').trim() || title; } catch (e) {}
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ======================================================================
 * AUTH
 * ==================================================================== */

function login(pin) {
  licenseGate_();
  pin = String(pin || '').trim();
  if (!/^\d{4,8}$/.test(pin)) throw new Error('Enter your 4–8 digit PIN');

  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('FAILS') || 0);
  if (fails >= CFG.MAX_FAILS) throw new Error('Too many wrong attempts. Wait 15 minutes.');

  if (!isSetupDone_()) throw new Error('Setup is not completed yet');
  const s = settings_();
  let user = null;
  if (pin === adminPin_(s)) {
    user = { role: 'admin', id: 'ADMIN', name: 'Admin' };
  } else {
    const m = readTable_(CFG.SHEET.SALESMEN)
      .find(r => isYes_(r.Active) && String(r.PIN).trim() === pin);
    if (m) user = { role: 'salesman', id: String(m.SalesmanID).trim(), name: String(m.Name).trim() };
  }

  if (!user) {
    cache.put('FAILS', String(fails + 1), 900);
    throw new Error('Wrong PIN');
  }

  const token = Utilities.getUuid();
  // PIN is kept server-side in the session so that changing a PIN logs that person out.
  cache.put('S_' + token, JSON.stringify(Object.assign({ pin: pin }, user)), CFG.SESSION_SECONDS);
  return { token: token, user: user, company: s.CompanyName || 'My Company', tagline: s.Tagline || '', logo: s.Logo || '',
    currency: s.Currency || 'MK', today: today_() };
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('S_' + token);
  return true;
}

function session_(token, role) {
  licenseGate_();
  const cache = CacheService.getScriptCache();
  const raw = token && cache.get('S_' + token);
  if (!raw) throw new Error('SESSION_EXPIRED');
  const user = JSON.parse(raw);
  if (role && user.role !== role) throw new Error('Not allowed');
  cache.put('S_' + token, raw, CFG.SESSION_SECONDS); // sliding expiry
  // Kick out deactivated salesmen and anyone whose PIN was changed by the admin.
  let valid;
  if (user.role === 'salesman') {
    valid = readTable_(CFG.SHEET.SALESMEN).some(r => String(r.SalesmanID).trim() === user.id &&
      isYes_(r.Active) && String(r.PIN).trim() === user.pin);
  } else {
    valid = adminPin_() === user.pin;
  }
  if (!valid) { cache.remove('S_' + token); throw new Error('SESSION_EXPIRED'); }
  delete user.pin;
  return user;
}

/* ======================================================================
 * SALESMAN API
 * ==================================================================== */

function getSalesmanHome(token) {
  const u = session_(token, 'salesman');
  const today = today_();
  const month = today.slice(0, 7);
  const me = readTable_(CFG.SHEET.SALESMEN).find(r => String(r.SalesmanID).trim() === u.id) || {};
  const mine = salesRows_().filter(r => r.SalesmanID === u.id);
  const live = mine.filter(r => r.Status !== 'VOID');

  const todayLive = live.filter(r => r.Date === today);
  const monthLive = live.filter(r => r.Date.slice(0, 7) === month);

  return {
    user: u,
    today: today,
    products: activeProducts_(),
    payments: PAYMENTS,
    voidReasons: VOID_REASONS,
    stats: {
      todayRevenue: sum_(todayLive, 'LineTotal'),
      todayUnits: sum_(todayLive, 'Qty'),
      todaySales: uniq_(todayLive.map(r => r.SaleID)).length,
      monthRevenue: sum_(monthLive, 'LineTotal'),
      monthTarget: Number(me.MonthlyTarget) || 0
    },
    todaySales: groupSales_(mine.filter(r => r.Date === today)).reverse()
  };
}

function recordSale(token, payload) {
  const u = session_(token, 'salesman');
  payload = payload || {};

  const products = {};
  activeProducts_().forEach(p => products[p.code] = p);

  const items = (payload.items || []).map(it => ({ code: String(it.code), qty: Number(it.qty) }))
    .filter(it => it.qty > 0);
  if (!items.length) throw new Error('Add at least one product');
  items.forEach(it => {
    if (!products[it.code]) throw new Error('Unknown product: ' + it.code);
    if (!Number.isInteger(it.qty) || it.qty > 100000) throw new Error('Invalid quantity');
  });

  const location = clean_(payload.location, 80);
  const notes = clean_(payload.notes, 200);
  const payment = PAYMENTS.indexOf(payload.payment) >= 0 ? payload.payment : 'Cash';
  const replaces = String(payload.replaces || '').trim();

  const cache = CacheService.getScriptCache();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    // Prevent double submission when the network is slow and the button is pressed twice.
    if (payload.clientId) {
      if (cache.get('C_' + payload.clientId)) return getSalesmanHome(token);
      cache.put('C_' + payload.clientId, '1', 21600);   // 6h: offline retries
    }
    const who = actor_(u);

    // Customer: must be a registered customer (by ID, or a new one created on purpose).
    let cust = null;
    const cid = String(payload.customerId || '').trim();
    if (cid) {
      cust = customersAll_().find(c => c.id === cid);
      if (!cust) throw new Error('Customer not found. Choose the customer from the list.');
      if (!cust.active) throw new Error('Customer "' + cust.name + '" is inactive. Ask the admin to reactivate it.');
    } else if (payload.newCustomer && clean_(payload.newCustomer.name, 80)) {
      cust = createCustomer_(who, payload.newCustomer.name, payload.newCustomer.phone, payload.newCustomer.location || location).customer;
    } else {
      const nm = clean_(payload.customer, 80);
      if (!nm) throw new Error('Choose the customer');
      cust = customersAll_().find(c => c.key === key_(nm) && c.active);
      if (!cust) throw new Error('Customer "' + nm + '" is not registered. Add the customer first.');
    }

    ensureSalesHeaders_();
    const sh = sheet_(CFG.SHEET.SALES);

    // Editing = a new sale linked to the old one; the old one is cancelled in the same step.
    let old = null;
    if (replaces) {
      const values = sh.getDataRange().getValues();
      const col = idx_(values[0]);
      const hits = rowsOfSale_(values, col, replaces);
      assertVoidable_(u, values, col, hits, replaces);
      old = { col: col, hits: hits };
    }

    const now = clientTime_(payload.at);
    const saleId = stamp_(now, 'B');
    const date = Utilities.formatDate(now, CFG.TZ, 'yyyy-MM-dd');
    const rows = items.map(it => {
      const p = products[it.code];
      return [saleId, now, date, u.id, u.name, p.code, p.name, p.category, it.qty, p.price,
        it.qty * p.price, cust.name, location, payment, notes, 'OK', '', '', '', cust.id, replaces];
    });
    const start = sh.getLastRow() + 1;
    const need = start + rows.length - 1 - sh.getMaxRows();
    if (need > 0) sh.insertRowsAfter(sh.getMaxRows(), need + 200);
    sh.getRange(start, 3, rows.length, 1).setNumberFormat('@');
    sh.getRange(start, 1, rows.length, SALES_HEADERS.length).setValues(rows);
    const total = rows.reduce((a, r) => a + r[10], 0);
    ledger_(who, 'SALE_CREATED', saleId, [cust.id, replaces], 'Sale ' + total + ' to ' + cust.name + ' (' + payment + ')' + (replaces ? ' - replaces ' + replaces : ''),
      { total: total, payment: payment, items: items });
    if (old) {
      applyVoid_(sh, old.col, old.hits, who, 'Edited - replaced by ' + saleId, new Date());
      ledger_(who, 'SALE_VOIDED', replaces, [saleId, cust.id], 'Sale edited - replaced by ' + saleId, { reason: 'edit', newSaleId: saleId });
    }
  } finally {
    unlock_(lock);
  }
  return getSalesmanHome(token);
}

/* ======================================================================
 * VOID (cancel) A SALE
 *  - Salesman: only his own sale, only on the same day.
 *  - Admin: any sale.
 *  - A reason is always required. Nothing is ever deleted: the rows stay
 *    in the sheet with Status = VOID, who cancelled it, when, and why.
 *  - To correct a mistake: cancel the wrong sale (with reason) and enter it again.
 * ==================================================================== */

function rowsOfSale_(values, col, saleId) {
  const hits = [];
  for (let i = 1; i < values.length; i++) if (String(values[i][col.SaleID]) === saleId) hits.push(i);
  return hits;
}

function assertVoidable_(u, values, col, hits, saleId) {
  if (!hits.length) throw new Error('Sale not found');
  hits.forEach(i => {
    const r = values[i];
    if (String(r[col.Status]).toUpperCase() === 'VOID') throw new Error('This sale is already cancelled');
    if (u.role === 'salesman') {
      if (String(r[col.SalesmanID]).trim() !== u.id) throw new Error('Not your sale');
      if (dateStr_(r[col.Date]) !== today_()) throw new Error('You can only change today\'s sales. Ask the admin.');
    }
  });
  const paid = allocRows_().filter(a => a.saleId === saleId && a.status !== 'VOID');
  if (paid.length) throw new Error('This sale already has payments (' + uniq_(paid.map(a => a.paymentId)).join(', ') + '). Cancel the payment first.');
}

function applyVoid_(sh, col, hits, who, reason, now) {
  hits.forEach(i => {
    sh.getRange(i + 1, col.Status + 1).setValue('VOID');
    sh.getRange(i + 1, col.VoidedBy + 1).setValue(who);
    sh.getRange(i + 1, col.VoidedAt + 1).setValue(now);
    sh.getRange(i + 1, col.VoidReason + 1).setValue(reason);
    sh.getRange(i + 1, 1, 1, SALES_HEADERS.length).setBackground('#fdecec');
  });
}

function voidSale(token, saleId, reason) {
  const u = session_(token);
  saleId = String(saleId || '');
  reason = clean_(reason, 200);
  if (reason.length < 4) throw new Error('Please write the reason for cancelling');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureSalesHeaders_();
    const sh = sheet_(CFG.SHEET.SALES);
    const values = sh.getDataRange().getValues();
    const col = idx_(values[0]);
    const hits = rowsOfSale_(values, col, saleId);
    assertVoidable_(u, values, col, hits, saleId);
    const who = actor_(u);
    applyVoid_(sh, col, hits, who, reason, new Date());
    const row = values[hits[0]];
    const cid = String(row[col.CustomerID] || '').trim() || ((customersAll_().find(c => c.key === key_(row[col.Customer])) || {}).id || '');
    ledger_(who, 'SALE_VOIDED', saleId, [cid], 'Sale cancelled: ' + reason, { reason: reason });
  } finally {
    unlock_(lock);
  }
  return u.role === 'salesman' ? getSalesmanHome(token) : true;
}

/* ======================================================================
 * ADMIN API – manage team, products and admin PIN (no code changes needed)
 *  - Salesmen and products are never deleted (their sales history must
 *    stay intact). Set them Inactive instead.
 * ==================================================================== */

function adminGetSetup(token) {
  session_(token, 'admin');
  const salesmen = readTable_(CFG.SHEET.SALESMEN).filter(m => String(m.SalesmanID).trim()).map(m => ({
    id: String(m.SalesmanID).trim(), name: String(m.Name || ''), pin: String(m.PIN || '').trim(),
    phone: String(m.Phone || ''), area: String(m.Area || ''), target: Number(m.MonthlyTarget) || 0, active: isYes_(m.Active)
  }));
  const products = readTable_(CFG.SHEET.PRODUCTS).filter(p => String(p.ProductCode).trim()).map(p => ({
    code: String(p.ProductCode).trim(), name: String(p.Name || ''), category: String(p.Category || ''),
    unit: String(p.Unit || ''), price: Number(p.Price) || 0, active: isYes_(p.Active)
  }));
  const cats = uniq_(products.map(p => p.category).filter(String));
  return { salesmen: salesmen, products: products, categories: cats };
}

function adminSaveSalesman(token, d) {
  session_(token, 'admin');
  d = d || {};
  const name = clean_(d.name, 60);
  const pin = String(d.pin || '').trim();
  if (!name) throw new Error('Name is required');
  if (!/^\d{4,8}$/.test(pin)) throw new Error('PIN must be 4–8 digits');
  if (pin === adminPin_()) throw new Error('This PIN is the Admin PIN. Choose another.');
  const target = Math.max(0, Math.round(Number(d.target) || 0));

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_(CFG.SHEET.SALESMEN);
    const values = sh.getDataRange().getValues();
    const col = idx_(values[0]);
    const id = String(d.id || '').trim();

    for (let i = 1; i < values.length; i++) {
      const rid = String(values[i][col.SalesmanID]).trim();
      if (rid && rid !== id && String(values[i][col.PIN]).trim() === pin) {
        throw new Error('PIN already used by ' + values[i][col.Name] + '. Each salesman needs a different PIN.');
      }
    }

    const data = { Name: name, PIN: pin, Phone: clean_(d.phone, 30), Area: clean_(d.area, 60),
      MonthlyTarget: target, Active: d.active === false ? 'No' : 'Yes' };

    let rowNum = 0;
    if (id) {
      for (let i = 1; i < values.length; i++) if (String(values[i][col.SalesmanID]).trim() === id) rowNum = i + 1;
      if (!rowNum) throw new Error('Salesman not found');
    } else {
      let max = 0;
      values.slice(1).forEach(r => { const m = String(r[col.SalesmanID]).match(/(\d+)$/); if (m) max = Math.max(max, Number(m[1])); });
      data.SalesmanID = 'S' + ('00' + (max + 1)).slice(-3);
      rowNum = Math.max(sh.getLastRow(), 1) + 1;
    }
    writeRow_(sh, values[0], rowNum, data, ['PIN']);
    const sid = data.SalesmanID || id;
    ledger_('Admin', 'SALESMAN_SAVED', sid, [sid], (id ? 'Updated ' : 'Added ') + name, { name: name, area: data.Area, target: target, active: data.Active });
  } finally {
    unlock_(lock);
  }
  return adminGetSetup(token);
}

function adminSaveProduct(token, d) {
  session_(token, 'admin');
  d = d || {};
  const name = clean_(d.name, 80);
  const category = clean_(d.category, 30);
  const price = Number(d.price);
  if (!name) throw new Error('Product name is required');
  if (!category) throw new Error('Category is required');
  if (!(price >= 0) || price > 100000000) throw new Error('Enter a valid price');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_(CFG.SHEET.PRODUCTS);
    const values = sh.getDataRange().getValues();
    const col = idx_(values[0]);
    const code = String(d.code || '').trim();
    const data = { Name: name, Category: category, Unit: clean_(d.unit, 30), Price: price, Active: d.active === false ? 'No' : 'Yes' };

    let rowNum = 0;
    if (code) {
      for (let i = 1; i < values.length; i++) if (String(values[i][col.ProductCode]).trim() === code) rowNum = i + 1;
      if (!rowNum) throw new Error('Product not found');
    } else {
      let max = 0;
      values.slice(1).forEach(r => { const m = String(r[col.ProductCode]).match(/^P(\d+)$/); if (m) max = Math.max(max, Number(m[1])); });
      data.ProductCode = 'P' + ('00' + (max + 1)).slice(-3);
      rowNum = Math.max(sh.getLastRow(), 1) + 1;
    }
    writeRow_(sh, values[0], rowNum, data, []);
    const pcode = data.ProductCode || code;
    const oldPrice = code ? Number(values[rowNum - 1][col.Price]) : null;
    ledger_('Admin', 'PRODUCT_SAVED', pcode, [pcode], (code ? 'Updated ' : 'Added ') + name + ' @ ' + price,
      { name: name, category: category, unit: data.Unit, price: price, oldPrice: oldPrice, active: data.Active });
  } finally {
    unlock_(lock);
  }
  return adminGetSetup(token);
}

function adminChangePin(token, current, next) {
  session_(token, 'admin');
  current = String(current || '').trim();
  next = String(next || '').trim();
  const s = settings_();
  if (current !== adminPin_(s)) throw new Error('Current Admin PIN is wrong');
  if (!/^\d{4,8}$/.test(next)) throw new Error('New PIN must be 4–8 digits');
  if (next === '0000') throw new Error('Choose a PIN other than 0000');
  if (readTable_(CFG.SHEET.SALESMEN).some(r => String(r.PIN).trim() === next)) throw new Error('A salesman already uses this PIN');

  const sh = sheet_(CFG.SHEET.SETTINGS);
  const values = sh.getDataRange().getValues();
  let row = 0;
  for (let i = 1; i < values.length; i++) if (String(values[i][0]).trim() === 'AdminPIN') row = i + 1;
  if (!row) { row = sh.getLastRow() + 1; sh.getRange(row, 1).setValue('AdminPIN'); }
  sh.getRange(row, 2).setNumberFormat('@').setValue(next);
  clearTableCache_();
  return true; // the current session ends; the admin logs in again with the new PIN
}

/** Writes the given {Header: value} fields into one row, leaving other columns untouched. */
function writeRow_(sh, headers, rowNum, data, textCols) {
  const col = idx_(headers);
  const width = headers.length;
  const existing = rowNum <= sh.getLastRow() ? sh.getRange(rowNum, 1, 1, width).getValues()[0] : new Array(width).fill('');
  Object.keys(data).forEach(k => {
    if (col[k] === undefined) throw new Error('Column "' + k + '" missing in sheet ' + sh.getName());
    existing[col[k]] = data[k];
  });
  textCols.forEach(k => sh.getRange(rowNum, col[k] + 1).setNumberFormat('@'));
  sh.getRange(rowNum, 1, 1, width).setValues([existing]);
}

/* ======================================================================
 * ADMIN API – main dashboard
 * ==================================================================== */

function getAdminData(token, from, to, salesmanId) {
  session_(token, 'admin');
  const today = today_();
  const re = /^\d{4}-\d{2}-\d{2}$/;
  if (!re.test(from || '')) from = today;
  if (!re.test(to || '')) to = today;
  if (from > to) { const t = from; from = to; to = t; }
  salesmanId = String(salesmanId || '');

  const s = settings_();
  const salesmen = readTable_(CFG.SHEET.SALESMEN).filter(r => String(r.SalesmanID).trim());
  const all = salesRows_();
  const live = all.filter(r => r.Status !== 'VOID');
  const month = today.slice(0, 7);

  // ---- Leaderboard (always all salesmen, for the selected range) ----
  const rangeAll = live.filter(r => r.Date >= from && r.Date <= to);
  const rangeVoid = all.filter(r => r.Status === 'VOID' && r.Date >= from && r.Date <= to);
  const board = salesmen.map(m => {
    const id = String(m.SalesmanID).trim();
    const rr = rangeAll.filter(r => r.SalesmanID === id);
    const mine = live.filter(r => r.SalesmanID === id);
    const mtd = sum_(mine.filter(r => r.Date.slice(0, 7) === month), 'LineTotal');
    const target = Number(m.MonthlyTarget) || 0;
    const last = mine.length ? mine[mine.length - 1].Timestamp : '';
    return {
      id: id, name: String(m.Name), area: String(m.Area || ''), active: isYes_(m.Active),
      revenue: sum_(rr, 'LineTotal'), units: sum_(rr, 'Qty'), sales: uniq_(rr.map(r => r.SaleID)).length,
      cats: unitsByCat_(rr),
      voids: uniq_(rangeVoid.filter(r => r.SalesmanID === id).map(r => r.SaleID)).length,
      voidValue: sum_(rangeVoid.filter(r => r.SalesmanID === id), 'LineTotal'),
      todayRevenue: sum_(mine.filter(r => r.Date === today), 'LineTotal'),
      lastSale: last, mtd: mtd, target: target, pct: target ? Math.round(mtd / target * 100) : null
    };
  }).filter(b => b.active || b.revenue > 0)
    .sort((a, b) => b.revenue - a.revenue);

  // ---- Everything below respects the salesman filter ----
  const rows = salesmanId ? rangeAll.filter(r => r.SalesmanID === salesmanId) : rangeAll;
  const revenue = sum_(rows, 'LineTotal');
  const voids = salesmanId ? rangeVoid.filter(r => r.SalesmanID === salesmanId) : rangeVoid;
  const saleIds = uniq_(rows.map(r => r.SaleID));

  // previous period of the same length ("vs previous" badges) and top customers
  const span = dateRange_(from, to, 400).length || 1;
  const pTo = addDays_(from, -1), pFrom = addDays_(from, -span);
  const pRows = live.filter(r => r.Date >= pFrom && r.Date <= pTo && (!salesmanId || r.SalesmanID === salesmanId));
  const prev = { revenue: sum_(pRows, 'LineTotal'), units: sum_(pRows, 'Qty'), sales: uniq_(pRows.map(r => r.SaleID)).length,
    customers: uniq_(pRows.map(r => r.Customer.toLowerCase())).length };
  const byCust = {};
  rows.forEach(r => {
    const k = r.Customer.toLowerCase();
    const c = byCust[k] || (byCust[k] = { name: r.Customer, revenue: 0, ids: {} });
    c.revenue += r.LineTotal; c.ids[r.SaleID] = 1;
  });
  const topCustomers = Object.keys(byCust).map(k => ({ name: byCust[k].name, revenue: byCust[k].revenue, sales: Object.keys(byCust[k].ids).length }))
    .sort((a, b) => b.revenue - a.revenue).slice(0, 8);

  const byProduct = {};
  const byCategory = {};
  const byPayment = {};
  const byDay = {};
  rows.forEach(r => {
    const p = byProduct[r.ProductCode] || (byProduct[r.ProductCode] = { code: r.ProductCode, name: r.Product, category: r.Category, units: 0, revenue: 0 });
    p.units += r.Qty; p.revenue += r.LineTotal;
    const c = byCategory[r.Category] || (byCategory[r.Category] = { units: 0, revenue: 0 });
    c.units += r.Qty; c.revenue += r.LineTotal;
    byPayment[r.Payment] = (byPayment[r.Payment] || 0) + r.LineTotal;
    const d = byDay[r.Date] || (byDay[r.Date] = {});
    d[r.Category] = (d[r.Category] || 0) + r.LineTotal;
  });

  const categories = Object.keys(byCategory).sort();
  const days = dateRange_(from, to, 120);
  const daily = { days: days, series: categories.map(c => ({ name: c, values: days.map(d => (byDay[d] && byDay[d][c]) || 0) })) };

  const recentSrc = (salesmanId ? all.filter(r => r.SalesmanID === salesmanId) : all)
    .filter(r => r.Date >= from && r.Date <= to);
  const recent = groupSales_(recentSrc).reverse().slice(0, 200);

  return {
    company: s.CompanyName || 'My Company',
    currency: s.Currency || 'MK',
    defaultPin: adminPin_(s) === '0000',
    today: today, from: from, to: to, salesmanId: salesmanId,
    salesmen: salesmen.map(m => ({ id: String(m.SalesmanID).trim(), name: String(m.Name) })),
    kpi: {
      revenue: revenue,
      units: sum_(rows, 'Qty'),
      sales: saleIds.length,
      avgSale: saleIds.length ? Math.round(revenue / saleIds.length) : 0,
      credit: byPayment['Credit'] || 0,
      activeSellers: uniq_(rows.map(r => r.SalesmanID)).length,
      customers: uniq_(rows.map(r => r.Customer.toLowerCase())).length,
      voids: uniq_(voids.map(r => r.SaleID)).length,
      voidValue: sum_(voids, 'LineTotal')
    },
    board: board, prev: prev, topCustomers: topCustomers,
    products: Object.keys(byProduct).map(k => byProduct[k]).sort((a, b) => b.revenue - a.revenue),
    categories: categories.map(c => ({ name: c, units: byCategory[c].units, revenue: byCategory[c].revenue })),
    payments: Object.keys(byPayment).map(k => ({ name: k, revenue: byPayment[k] })).sort((a, b) => b.revenue - a.revenue),
    daily: daily,
    recent: recent,
    cancelled: groupSales_(voids).sort((a, b) => (b.voidedAt > a.voidedAt ? 1 : -1)),
    voidReasons: VOID_REASONS,
    updated: Utilities.formatDate(new Date(), CFG.TZ, 'HH:mm')
  };
}

/* ======================================================================
 * HELPERS
 * ==================================================================== */

/* ---- read caching: per-request memo + 60 s shared cache for the small tables ---- */
let MEMO = {};
const CACHED_TABLES = ['Settings', 'Salesmen', 'Products'];
function clearTableCache_() {
  MEMO = {};
  try { CacheService.getScriptCache().removeAll(CACHED_TABLES.map(n => 'T_' + n)); } catch (e) {}
}
/** Releases a write lock and drops cached reads so the next read sees the new data. */
function unlock_(lock) { lock.releaseLock(); clearTableCache_(); }

function readTable_(name) {
  if (MEMO[name]) return MEMO[name];
  const cacheIt = CACHED_TABLES.indexOf(name) >= 0;
  let out = null;
  if (cacheIt) { try { const raw = CacheService.getScriptCache().get('T_' + name); if (raw) out = JSON.parse(raw); } catch (e) {} }
  if (!out) {
    const values = sheet_(name).getDataRange().getValues();
    const headers = values.shift().map(h => String(h).trim());
    out = values.filter(r => r.join('') !== '').map(r => {
      const o = {};
      headers.forEach((h, i) => o[h] = r[i]);
      return o;
    });
    if (cacheIt) { try { CacheService.getScriptCache().put('T_' + name, JSON.stringify(out), 60); } catch (e) {} }
  }
  return (MEMO[name] = out);
}

function settings_() {
  const o = {};
  readTable_(CFG.SHEET.SETTINGS).forEach(r => o[String(r.Key).trim()] = r.Value);
  return o;
}

function activeProducts_() {
  return readTable_(CFG.SHEET.PRODUCTS).filter(p => isYes_(p.Active) && String(p.ProductCode).trim())
    .map(p => ({
      code: String(p.ProductCode).trim(), name: String(p.Name).trim(), category: String(p.Category || 'Other').trim(),
      unit: String(p.Unit || '').trim(), price: Number(p.Price) || 0
    }));
}

function salesRows_() {
  const byKey = {}, byId = {};
  customersAll_().forEach(c => { byKey[c.key] = c; byId[c.id] = c; });
  return readTable_(CFG.SHEET.SALES).map(r => {
  const c = byId[String(r.CustomerID || '').trim()] || byKey[key_(r.Customer)] || null;
  return {
    SaleID: String(r.SaleID),
    Timestamp: r.Timestamp instanceof Date ? Utilities.formatDate(r.Timestamp, CFG.TZ, 'yyyy-MM-dd HH:mm') : String(r.Timestamp),
    Date: dateStr_(r.Date),
    SalesmanID: String(r.SalesmanID).trim(),
    Salesman: String(r.Salesman),
    ProductCode: String(r.ProductCode),
    Product: String(r.Product),
    Category: String(r.Category || 'Other'),
    Qty: Number(r.Qty) || 0,
    UnitPrice: Number(r.UnitPrice) || 0,
    LineTotal: Number(r.LineTotal) || 0,
    CustomerID: c ? c.id : '',
    Customer: c ? c.name : String(r.Customer || ''),
    Location: String(r.Location || ''),
    Payment: String(r.Payment || 'Cash'),
    Notes: String(r.Notes || ''),
    Status: String(r.Status || 'OK').toUpperCase(),
    VoidedBy: String(r.VoidedBy || ''),
    VoidedAt: r.VoidedAt instanceof Date ? Utilities.formatDate(r.VoidedAt, CFG.TZ, 'yyyy-MM-dd HH:mm') : String(r.VoidedAt || ''),
    VoidReason: String(r.VoidReason || ''),
    Replaces: String(r.Replaces || '')
  }; });
}

function groupSales_(rows) {
  const map = {};
  const order = [];
  rows.forEach(r => {
    let g = map[r.SaleID];
    if (!g) {
      g = map[r.SaleID] = {
        saleId: r.SaleID, time: r.Timestamp, date: r.Date, salesman: r.Salesman, salesmanId: r.SalesmanID,
        customerId: r.CustomerID, customer: r.Customer, location: r.Location, payment: r.Payment, notes: r.Notes, status: r.Status,
        voidedBy: r.VoidedBy, voidedAt: r.VoidedAt, voidReason: r.VoidReason, replaces: r.Replaces, total: 0, items: []
      };
      order.push(g);
    }
    g.items.push({ name: r.Product, qty: r.Qty, code: r.ProductCode });
    g.total += r.LineTotal;
  });
  const settle = settle_();
  order.forEach(g => {
    if (g.payment === 'Credit' && g.status !== 'VOID') {
      g.paid = Math.min(g.total, settle[g.saleId] || 0);
      g.remaining = round2_(g.total - g.paid);
      g.payStatus = g.remaining <= 0.005 ? 'PAID' : (g.paid > 0 ? 'PARTIAL' : 'OPEN');
    } else { g.paid = 0; g.remaining = 0; g.payStatus = ''; }
  });
  return order;
}

/** Adds any missing columns (e.g. VoidedBy/VoidedAt/VoidReason) to an older Sales sheet. */
function ensureSalesHeaders_() {
  const sh = sheet_(CFG.SHEET.SALES);
  const width = Math.max(sh.getLastColumn(), 1);
  const have = sh.getRange(1, 1, 1, width).getValues()[0].map(h => String(h).trim());
  SALES_HEADERS.forEach((h, i) => {
    if (have[i] === h) return;
    if (have.indexOf(h) >= 0) throw new Error('Sales sheet columns are out of order – column "' + h + '" moved');
    if (have[i]) throw new Error('Unexpected column "' + have[i] + '" in Sales sheet');
    sh.getRange(1, i + 1).setValue(h).setFontWeight('bold').setBackground('#0b4f8a').setFontColor('#ffffff');
    have[i] = h;
  });
}

function idx_(headers) {
  const o = {};
  headers.forEach((h, i) => o[String(h).trim()] = i);
  return o;
}

function dateStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, CFG.TZ, 'yyyy-MM-dd');
  return String(v || '').trim().slice(0, 10);
}

function dateRange_(from, to, max) {
  const out = [];
  const d = new Date(from + 'T12:00:00Z');
  const end = new Date(to + 'T12:00:00Z');
  while (d <= end && out.length < max) {
    out.push(d.toISOString().slice(0, 10));
    d.setUTCDate(d.getUTCDate() + 1);
  }
  return out;
}

function addDays_(d, n) { const x = new Date(d + 'T12:00:00Z'); x.setUTCDate(x.getUTCDate() + n); return x.toISOString().slice(0, 10); }
function today_() { return Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd'); }
function isYes_(v) { return /^(yes|y|true|1|active)$/i.test(String(v).trim()); }
function sum_(rows, key) { return rows.reduce((a, r) => a + (Number(r[key]) || 0), 0); }
function uniq_(arr) { return Array.from(new Set(arr)); }
function clean_(v, max) { return String(v || '').replace(/[\r\n\t]+/g, ' ').replace(/^[=+\-@]+/, '').trim().slice(0, max); }

/* ======================================================================
 * CUSTOMER REGISTRY, CREDIT INVOICES, PAYMENT ALLOCATION & LEDGER
 *
 *  - Every customer has an ID (C0001...). Sales, payments and allocations store the ID,
 *    so renaming or merging a customer never breaks history. New customers are created
 *    on purpose (with a similar-name check), never by typing a different spelling.
 *  - A credit sale is an invoice: Open -> Partial -> Paid. A payment is ALLOCATED to
 *    specific invoices (oldest first by default), so it is always known which sale it paid.
 *  - Ledger = append-only event log. Each row stores the hash of the previous row
 *    (tamper-evident chain). Search any ID (sale B..., payment R..., customer C...) to see
 *    every linked record and the full timeline.
 * ==================================================================== */

const CUSTOMER_HEADERS = ['CustomerID', 'Name', 'Phone', 'Location', 'Active', 'CreatedAt', 'CreatedBy'];
const PAYMENT_HEADERS = ['PaymentID', 'Timestamp', 'Date', 'Customer', 'Amount', 'Method', 'ReceivedBy', 'Notes',
  'Status', 'VoidedBy', 'VoidReason', 'CustomerID'];
const ALLOC_HEADERS = ['AllocID', 'Timestamp', 'PaymentID', 'SaleID', 'CustomerID', 'Amount', 'Status'];
const LEDGER_HEADERS = ['EventID', 'Time', 'Actor', 'Type', 'RefID', 'LinkedIDs', 'Summary', 'Data', 'PrevHash', 'Hash'];

function key_(n) { return String(n || '').toLowerCase().replace(/\s+/g, ' ').trim(); }
function digits_(v) { return String(v || '').replace(/\D+/g, ''); }
function round2_(n) { return Math.round((Number(n) || 0) * 100) / 100; }
function actor_(u) { return u.role === 'admin' ? 'Admin' : u.name + ' (' + u.id + ')'; }
function stamp_(now, prefix) { return prefix + Utilities.formatDate(now, CFG.TZ, 'yyMMddHHmmss') + Math.floor(Math.random() * 900 + 100); }
function idType_(id) {
  id = String(id);
  return /^B\d/.test(id) ? 'sale' : /^R\d/.test(id) ? 'payment' : /^C\d/.test(id) ? 'customer' : /^E\d/.test(id) ? 'event'
    : /^A\d/.test(id) ? 'allocation' : /^P\d/.test(id) ? 'product' : /^S\d/.test(id) ? 'salesman' : 'id';
}
function sha_(str) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

let HDR_DONE = {};
/** Returns the sheet, creating it and adding any missing header columns (safe for older installs). */
function tab_(name, headers) {
  const ss = ss_();
  const sh = ss.getSheetByName(name) || makeSheet_(ss, name, headers, []);
  if (!HDR_DONE[name]) {
    const have = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(h => String(h).trim());
    headers.forEach((h, i) => {
      if (!have[i]) sh.getRange(1, i + 1).setValue(h).setFontWeight('bold').setBackground('#0b4f8a').setFontColor('#ffffff');
    });
    HDR_DONE[name] = true;
  }
  return sh;
}

function clientTime_(v) {
  const d = v ? new Date(v) : null, now = new Date();
  if (!d || isNaN(d) || d > now || now - d > 7 * 86400000) return now;
  return d;
}

/* ---------------- ledger ---------------- */

function ledgerSh_() {
  const ss = ss_();
  if (!ss.getSheetByName('Ledger')) { const n = ss.insertSheet('Ledger'); n.getRange('A:J').setNumberFormat('@'); }
  return tab_('Ledger', LEDGER_HEADERS);
}

/** Appends one event to the hash chain. Call inside the write lock. Returns the event id. */
function ledger_(actor, type, refId, linked, summary, data) {
  const sh = ledgerSh_();
  const last = sh.getLastRow();
  const prev = last > 1 ? String(sh.getRange(last, 10).getValue()) : 'GENESIS';
  const now = new Date();
  const id = stamp_(now, 'E');
  const time = now.toISOString();
  const links = uniq_((linked || []).filter(Boolean).map(String)).join(',');
  const sum = String(summary || '').slice(0, 300);
  const d = JSON.stringify(data || {});
  const hash = sha_([prev, id, time, actor, type, refId, links, sum, d].join('|'));
  sh.getRange(last + 1, 1, 1, 10).setValues([[id, time, actor, type, String(refId), links, sum, d, prev, hash]]);
  delete MEMO['Ledger'];
  return id;
}

function ledgerRows_() {
  ledgerSh_();
  return readTable_('Ledger').map(r => ({
    id: String(r.EventID), time: String(r.Time), actor: String(r.Actor), type: String(r.Type), ref: String(r.RefID),
    linked: String(r.LinkedIDs || '').split(',').filter(Boolean), summary: String(r.Summary || ''), data: String(r.Data || ''),
    prev: String(r.PrevHash), hash: String(r.Hash)
  }));
}

/** Admin: re-computes the whole chain and reports the first row that does not match. */
function verifyLedger(token) {
  session_(token, 'admin');
  const rows = ledgerRows_();
  let prev = 'GENESIS';
  for (let i = 0; i < rows.length; i++) {
    const e = rows[i];
    const expect = sha_([prev, e.id, e.time, e.actor, e.type, e.ref, e.linked.join(','), e.summary, e.data].join('|'));
    if (e.prev !== prev || e.hash !== expect) return { ok: false, count: rows.length, brokenAt: { row: i + 2, eventId: e.id } };
    prev = e.hash;
  }
  return { ok: true, count: rows.length };
}

/* ---------------- customers ---------------- */

function customersAll_() {
  tab_('Customers', CUSTOMER_HEADERS);
  return readTable_('Customers').filter(c => String(c.Name || '').trim()).map(c => ({
    id: String(c.CustomerID).trim(), name: String(c.Name).trim(), phone: String(c.Phone || '').trim(),
    location: String(c.Location || ''), active: isYes_(c.Active), key: key_(c.Name)
  }));
}
function pubCust_(c) { return { id: c.id, name: c.name, phone: c.phone, location: c.location, active: c.active }; }

function lev_(a, b) {
  const m = a.length, n = b.length;
  if (!m) return n; if (!n) return m;
  let prev = []; for (let j = 0; j <= n; j++) prev[j] = j;
  for (let i = 1; i <= m; i++) {
    const cur = [i];
    for (let j = 1; j <= n; j++) cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    prev = cur;
  }
  return prev[n];
}

function similarCustomers_(name, phone) {
  const k = key_(name), out = [];
  customersAll_().forEach(c => {
    if (!c.active) return;
    const same = c.key === k;
    const near = !same && ((k.length >= 4 && c.key.length >= 4 && (c.key.indexOf(k) >= 0 || k.indexOf(c.key) >= 0)) ||
      (Math.max(k.length, c.key.length) >= 5 && lev_(k, c.key) <= 2));
    const samePhone = digits_(phone).length >= 6 && digits_(c.phone) === digits_(phone);
    if (same || near || samePhone) out.push(Object.assign(pubCust_(c), { reason: same ? 'same name' : samePhone ? 'same phone' : 'similar name' }));
  });
  return out.slice(0, 5);
}

/** Inside the lock. Returns { customer, existed }. An exact (normalized) name match is never duplicated. */
function createCustomer_(actor, name, phone, location) {
  name = clean_(name, 80);
  if (!name) throw new Error('Customer name is required');
  const all = customersAll_();
  const hit = all.find(c => c.key === key_(name));
  if (hit) {
    if (!hit.active) throw new Error('Customer "' + hit.name + '" is inactive. Ask the admin to reactivate it.');
    return { customer: hit, existed: true };
  }
  const sh = tab_('Customers', CUSTOMER_HEADERS);
  let max = 0;
  all.forEach(c => { const m = c.id.match(/(\d+)$/); if (m) max = Math.max(max, Number(m[1])); });
  const id = 'C' + String(max + 1).padStart(4, '0');
  const row = sh.getLastRow() + 1;
  sh.getRange(row, 3).setNumberFormat('@'); sh.getRange(row, 6).setNumberFormat('@');
  sh.getRange(row, 1, 1, 7).setValues([[id, name, clean_(phone, 30), clean_(location, 80), 'Yes', new Date().toISOString(), actor]]);
  delete MEMO['Customers'];
  ledger_(actor, 'CUSTOMER_CREATED', id, [id], 'Customer created: ' + name, { name: name, phone: clean_(phone, 30), location: clean_(location, 80) });
  return { customer: { id: id, name: name, phone: clean_(phone, 30), location: clean_(location, 80), active: true, key: key_(name) }, existed: false };
}

/** Salesman/admin: register a customer. Returns {status:'created'|'exists'|'similar', customer|similar[]}. */
function addCustomer(token, d) {
  const u = session_(token);
  d = d || {};
  const name = clean_(d.name, 80);
  if (!name) throw new Error('Customer name is required');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let res;
  try {
    const exact = customersAll_().find(c => c.key === key_(name));
    if (exact) {
      if (!exact.active) throw new Error('Customer "' + exact.name + '" is inactive. Ask the admin to reactivate it.');
      res = { status: 'exists', customer: pubCust_(exact) };
    } else {
      const sim = similarCustomers_(name, d.phone);
      if (sim.length && !d.confirmNew) res = { status: 'similar', similar: sim };
      else res = { status: 'created', customer: pubCust_(createCustomer_(actor_(u), name, d.phone, d.location).customer) };
    }
  } finally {
    unlock_(lock);
  }
  return res;
}

/** Admin: edit a customer (rename is safe because everything links by ID). */
function adminSaveCustomer(token, d) {
  const u = session_(token, 'admin');
  d = d || {};
  const id = String(d.id || '').trim();
  const name = clean_(d.name, 80);
  if (!id || !name) throw new Error('Name is required');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const all = customersAll_();
    const cur = all.find(c => c.id === id);
    if (!cur) throw new Error('Customer not found');
    const clash = all.find(c => c.id !== id && c.key === key_(name));
    if (clash) throw new Error('Another customer is already called "' + clash.name + '" (' + clash.id + '). Use Merge instead.');
    const sh = tab_('Customers', CUSTOMER_HEADERS);
    const values = sh.getDataRange().getValues();
    const col = idx_(values[0]);
    let rowNum = 0;
    for (let i = 1; i < values.length; i++) if (String(values[i][col.CustomerID]).trim() === id) rowNum = i + 1;
    writeRow_(sh, values[0], rowNum, { Name: name, Phone: clean_(d.phone, 30), Location: clean_(d.location, 80), Active: d.active === false ? 'No' : 'Yes' }, ['Phone']);
    ledger_('Admin', 'CUSTOMER_UPDATED', id, [id], 'Customer updated: ' + cur.name + (cur.name !== name ? ' -> ' + name : ''),
      { before: pubCust_(cur), after: { name: name, phone: clean_(d.phone, 30), location: clean_(d.location, 80), active: d.active !== false } });
  } finally {
    unlock_(lock);
  }
  return adminListCustomers(token);
}

/** Moves every sale, payment and allocation of `fromId` to `intoId` and deactivates `fromId`. */
function adminMergeCustomers(token, fromId, intoId) {
  session_(token, 'admin');
  fromId = String(fromId || '').trim(); intoId = String(intoId || '').trim();
  if (!fromId || !intoId || fromId === intoId) throw new Error('Choose two different customers');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const all = customersAll_();
    const from = all.find(c => c.id === fromId), into = all.find(c => c.id === intoId);
    if (!from || !into) throw new Error('Customer not found');
    ensureSalesHeaders_();
    let moved = 0;
    [[CFG.SHEET.SALES, null], ['Payments', PAYMENT_HEADERS], ['Allocations', ALLOC_HEADERS]].forEach(pair => {
      const sh = pair[1] ? tab_(pair[0], pair[1]) : sheet_(pair[0]);
      const values = sh.getDataRange().getValues();
      if (values.length < 2) return;
      const col = idx_(values[0]);
      const out = [];
      for (let i = 1; i < values.length; i++) {
        const cid = String(values[i][col.CustomerID] || '').trim();
        const nm = col.Customer !== undefined ? key_(values[i][col.Customer]) : '';
        const hit = cid === fromId || (!cid && nm && nm === from.key);
        if (hit) moved++;
        out.push([hit ? intoId : values[i][col.CustomerID]]);
      }
      sh.getRange(2, col.CustomerID + 1, out.length, 1).setValues(out);
    });
    const sh = tab_('Customers', CUSTOMER_HEADERS);
    const values = sh.getDataRange().getValues();
    const col = idx_(values[0]);
    for (let i = 1; i < values.length; i++) if (String(values[i][col.CustomerID]).trim() === fromId) sh.getRange(i + 1, col.Active + 1).setValue('No');
    ledger_('Admin', 'CUSTOMER_MERGED', intoId, [fromId, intoId], 'Merged ' + from.name + ' (' + fromId + ') into ' + into.name + ' (' + intoId + ')', { fromId: fromId, intoId: intoId, records: moved });
  } finally {
    unlock_(lock);
  }
  return adminListCustomers(token);
}

/** Admin: all customers with balance and activity. */
function adminListCustomers(token) {
  session_(token, 'admin');
  const bal = balances_();
  const stats = {};
  groupSales_(salesRows_().filter(r => r.Status !== 'VOID')).forEach(g => {
    const s = stats[g.customerId] || (stats[g.customerId] = { sales: 0, revenue: 0, last: '' });
    s.sales++; s.revenue += g.total; if (g.date > s.last) s.last = g.date;
  });
  return customersAll_().map(c => Object.assign(pubCust_(c), {
    balance: (bal[c.id] || {}).balance || 0, sales: (stats[c.id] || {}).sales || 0,
    revenue: (stats[c.id] || {}).revenue || 0, lastSale: (stats[c.id] || {}).last || ''
  })).sort((a, b) => a.name.localeCompare(b.name));
}

/** One-time helper for older data: registers every customer found in past sales and allocates old payments (oldest invoice first). */
function migrateLegacy() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    ensureSalesHeaders_();
    salesRows_().forEach(r => {
      if (r.Customer && !customersAll_().some(c => c.key === key_(r.Customer))) createCustomer_('System', r.Customer, '', r.Location);
    });
    MEMO = {};
    const done = {};
    allocRows_().forEach(a => { done[a.paymentId] = true; });
    paymentsRows_().filter(p => p.Status !== 'VOID' && !done[p.PaymentID]).sort((a, b) => (a.PaymentID < b.PaymentID ? -1 : a.PaymentID > b.PaymentID ? 1 : 0)).forEach(p => {
      const open = openInvoices_(p.CustomerID);
      let left = p.Amount; const rows = [], now = new Date();
      open.forEach(inv => {
        if (left <= 0.0049) return;
        const take = round2_(Math.min(left, inv.remaining));
        rows.push(['A' + p.PaymentID.slice(1) + '-' + (rows.length + 1), now, p.PaymentID, inv.saleId, p.CustomerID, take, 'OK']);
        left = round2_(left - take);
      });
      if (rows.length) {
        const sh = tab_('Allocations', ALLOC_HEADERS);
        sh.getRange(sh.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
        ledger_('System', 'PAYMENT_ALLOCATED', p.PaymentID, [p.CustomerID].concat(rows.map(r => r[3])), 'Old payment allocated to invoices', { allocations: rows.map(r => ({ saleId: r[3], amount: r[5] })) });
        MEMO = {};
      }
    });
  } finally {
    unlock_(lock);
  }
}

/* ---------------- payments, invoices, balances ---------------- */

function paymentsRows_() {
  tab_('Payments', PAYMENT_HEADERS);
  const byKey = {}, byId = {};
  customersAll_().forEach(c => { byKey[c.key] = c; byId[c.id] = c; });
  return readTable_('Payments').map(r => {
    const c = byId[String(r.CustomerID || '').trim()] || byKey[key_(r.Customer)] || null;
    return {
      PaymentID: String(r.PaymentID),
      Timestamp: r.Timestamp instanceof Date ? Utilities.formatDate(r.Timestamp, CFG.TZ, 'yyyy-MM-dd HH:mm') : String(r.Timestamp),
      Date: dateStr_(r.Date), CustomerID: c ? c.id : '', Customer: c ? c.name : String(r.Customer || ''), Amount: Number(r.Amount) || 0,
      Method: String(r.Method || 'Cash'), ReceivedBy: String(r.ReceivedBy || ''), Notes: String(r.Notes || ''),
      Status: String(r.Status || 'OK').toUpperCase(), VoidedBy: String(r.VoidedBy || ''), VoidReason: String(r.VoidReason || '')
    };
  });
}

function allocRows_() {
  tab_('Allocations', ALLOC_HEADERS);
  return readTable_('Allocations').map(r => ({
    id: String(r.AllocID), paymentId: String(r.PaymentID), saleId: String(r.SaleID), customerId: String(r.CustomerID || ''),
    amount: Number(r.Amount) || 0, status: String(r.Status || 'OK').toUpperCase()
  }));
}

/** { saleId: amount paid } from active allocations (cached per request). */
function settle_() {
  if (MEMO._settle) return MEMO._settle;
  const o = {};
  allocRows_().forEach(a => { if (a.status !== 'VOID') o[a.saleId] = round2_((o[a.saleId] || 0) + a.amount); });
  return (MEMO._settle = o);
}

/** Credit invoices of one customer that still have something to pay, oldest first. */
function openInvoices_(customerId) {
  return groupSales_(salesRows_().filter(r => r.Status !== 'VOID' && r.Payment === 'Credit' && r.CustomerID === customerId))
    .filter(g => g.remaining > 0.005)
    .map((g, i) => ({ g: g, i: i, k: g.saleId.slice(0, 13) }))     // sale ID starts with its date-time, so this is oldest first
    .sort((x, y) => (x.k < y.k ? -1 : x.k > y.k ? 1 : x.i - y.i)).map(o => o.g);
}

/** { customerId: { id, name, owed, paid, balance, lastCredit, lastPayment } } */
function balances_() {
  const b = {};
  const get = (id, name) => b[id] || (b[id] = { id: id, name: name, owed: 0, paid: 0, balance: 0, lastCredit: '', lastPayment: '', openInvoices: 0 });
  groupSales_(salesRows_().filter(r => r.Status !== 'VOID' && r.Payment === 'Credit')).forEach(g => {
    const e = get(g.customerId || 'N:' + key_(g.customer), g.customer);
    e.owed += g.total; e.paid += g.paid;
    if (g.remaining > 0.005) e.openInvoices++;
    if (g.date > e.lastCredit) e.lastCredit = g.date;
  });
  paymentsRows_().forEach(p => {
    if (p.Status === 'VOID') return;
    const e = b[p.CustomerID]; if (e && p.Date > e.lastPayment) e.lastPayment = p.Date;
  });
  Object.keys(b).forEach(k => { b[k].balance = round2_(b[k].owed - b[k].paid); });
  return b;
}

function getCustomers(token) {
  session_(token);
  const bal = balances_();
  return customersAll_().filter(c => c.active).map(c => Object.assign(pubCust_(c), { balance: (bal[c.id] || {}).balance || 0 }))
    .sort((a, b) => a.name.localeCompare(b.name));
}

function invoiceView_(g) {
  return { saleId: g.saleId, time: g.time, date: g.date, total: g.total, paid: g.paid, remaining: g.remaining, payStatus: g.payStatus, items: g.items, salesman: g.salesman };
}

/** Open credit invoices of a customer (used by the payment screen). */
function getOpenInvoices(token, customerId) {
  session_(token);
  return openInvoices_(String(customerId || '').trim()).map(invoiceView_);
}

/** Everything about one customer: sales (with invoice status), payments (with allocations), balance. */
function getCustomerDetail(token, customerId) {
  session_(token);
  const id = String(customerId || '').trim();
  const c = customersAll_().find(x => x.id === id);
  if (!c) throw new Error('Customer not found');
  const b = balances_()[id] || { owed: 0, paid: 0, balance: 0 };
  const allocs = allocRows_().filter(a => a.status !== 'VOID');
  return {
    customer: pubCust_(c), balance: b.balance, owed: b.owed, paid: b.paid,
    sales: groupSales_(salesRows_().filter(r => r.CustomerID === id)).reverse().slice(0, 60),
    payments: paymentsRows_().filter(p => p.CustomerID === id).reverse().slice(0, 60)
      .map(p => Object.assign(p, { allocations: allocs.filter(a => a.paymentId === p.PaymentID).map(a => ({ saleId: a.saleId, amount: a.amount })) }))
  };
}

/**
 * Record money received. p = { customerId, amount, method, notes, allocations?:[{saleId,amount}], clientId, at }.
 * Without `allocations` the payment is applied to the oldest open invoices first.
 */
function recordPayment(token, p) {
  const u = session_(token);
  p = p || {};
  const amount = round2_(p.amount);
  if (!(amount > 0) || amount > 1000000000) throw new Error('Enter a valid amount');
  const method = PAYMENTS.indexOf(p.method) >= 0 && p.method !== 'Credit' ? p.method : 'Cash';
  const cache = CacheService.getScriptCache();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let out;
  try {
    if (p.clientId) {
      if (cache.get('P_' + p.clientId)) return { duplicate: true, balance: 0 };
      cache.put('P_' + p.clientId, '1', 21600);
    }
    const cid = String(p.customerId || '').trim();
    const cust = customersAll_().find(c => cid ? c.id === cid : c.key === key_(p.customer));
    if (!cust) throw new Error('Choose the customer');
    const open = openInvoices_(cust.id);
    const totalOpen = round2_(open.reduce((a, i) => a + i.remaining, 0));
    if (totalOpen <= 0) throw new Error('This customer owes nothing');
    if (amount > totalOpen + 0.001) throw new Error('Amount is more than the balance owed (' + totalOpen + ')');

    const allocs = [];
    if (Array.isArray(p.allocations) && p.allocations.length) {
      let sum = 0; const seen = {};
      p.allocations.forEach(a => {
        const inv = open.find(i => i.saleId === String(a.saleId));
        const amt = round2_(a.amount);
        if (!inv) throw new Error('This invoice is not open: ' + a.saleId);
        if (seen[inv.saleId]) throw new Error('Invoice listed twice: ' + inv.saleId);
        if (!(amt > 0) || amt > inv.remaining + 0.001) throw new Error('Amount is more than the invoice balance for ' + inv.saleId);
        seen[inv.saleId] = 1; sum += amt; allocs.push({ saleId: inv.saleId, amount: amt });
      });
      if (Math.abs(sum - amount) > 0.01) throw new Error('The invoice amounts must add up to the payment amount');
    } else {
      let left = amount;
      open.forEach(inv => {
        if (left <= 0.0049) return;
        const take = round2_(Math.min(left, inv.remaining));
        allocs.push({ saleId: inv.saleId, amount: take });
        left = round2_(left - take);
      });
    }

    const now = clientTime_(p.at);
    const pid = stamp_(now, 'R');
    const sh = tab_('Payments', PAYMENT_HEADERS);
    const row = sh.getLastRow() + 1;
    sh.getRange(row, 3).setNumberFormat('@');
    sh.getRange(row, 1, 1, PAYMENT_HEADERS.length).setValues([[pid, now, Utilities.formatDate(now, CFG.TZ, 'yyyy-MM-dd'),
      cust.name, amount, method, actor_(u), clean_(p.notes, 200), 'OK', '', '', cust.id]]);
    const ash = tab_('Allocations', ALLOC_HEADERS);
    const arows = allocs.map((a, i) => ['A' + pid.slice(1) + '-' + (i + 1), now, pid, a.saleId, cust.id, a.amount, 'OK']);
    ash.getRange(ash.getLastRow() + 1, 1, arows.length, 7).setValues(arows);
    MEMO = {};
    const who = actor_(u);
    ledger_(who, 'PAYMENT_RECEIVED', pid, [cust.id].concat(allocs.map(a => a.saleId)),
      'Received ' + amount + ' from ' + cust.name + ' (' + method + ')', { amount: amount, method: method, allocations: allocs });
    allocs.forEach(a => {
      const inv = open.find(i => i.saleId === a.saleId);
      if (inv && round2_(inv.remaining - a.amount) <= 0.005) ledger_(who, 'SALE_SETTLED', a.saleId, [pid, cust.id], 'Invoice fully paid', { paymentId: pid });
    });
    out = { paymentId: pid, balance: round2_(totalOpen - amount), allocations: allocs };
  } finally {
    unlock_(lock);
  }
  return out;
}

/** Admin only: cancel a wrongly entered payment. Its allocations are released, the invoices re-open. */
function voidPayment(token, paymentId, reason) {
  const u = session_(token, 'admin');
  reason = clean_(reason, 200);
  if (reason.length < 4) throw new Error('Please write the reason for cancelling');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = tab_('Payments', PAYMENT_HEADERS);
    const values = sh.getDataRange().getValues();
    const col = idx_(values[0]);
    let row = 0;
    for (let i = 1; i < values.length; i++) if (String(values[i][col.PaymentID]) === String(paymentId)) row = i + 1;
    if (!row) throw new Error('Payment not found');
    if (String(values[row - 1][col.Status]).toUpperCase() === 'VOID') throw new Error('Already cancelled');
    sh.getRange(row, col.Status + 1).setValue('VOID');
    sh.getRange(row, col.VoidedBy + 1).setValue('Admin');
    sh.getRange(row, col.VoidReason + 1).setValue(reason);
    sh.getRange(row, 1, 1, PAYMENT_HEADERS.length).setBackground('#fdecec');
    const ash = tab_('Allocations', ALLOC_HEADERS);
    const av = ash.getDataRange().getValues();
    const ac = idx_(av[0]);
    const sales = [], cids = [];
    for (let i = 1; i < av.length; i++) {
      if (String(av[i][ac.PaymentID]) === String(paymentId) && String(av[i][ac.Status]).toUpperCase() !== 'VOID') {
        ash.getRange(i + 1, ac.Status + 1).setValue('VOID');
        sales.push(String(av[i][ac.SaleID])); cids.push(String(av[i][ac.CustomerID]));
      }
    }
    ledger_('Admin', 'PAYMENT_VOIDED', String(paymentId), cids.concat(sales), 'Payment cancelled: ' + reason, { reason: reason, releasedInvoices: sales });
  } finally {
    unlock_(lock);
  }
  return true;
}

/** Admin: who owes how much (biggest first), invoice counts and recent payments with their allocations. */
function getReceivables(token) {
  session_(token, 'admin');
  const phones = {};
  customersAll_().forEach(c => { phones[c.id] = c.phone; });
  const bal = balances_();
  const list = Object.keys(bal).map(k => Object.assign({ phone: phones[k] || '' }, bal[k])).filter(e => e.balance > 0.005).sort((a, b) => b.balance - a.balance);
  const allocs = allocRows_().filter(a => a.status !== 'VOID');
  return {
    total: round2_(list.reduce((a, e) => a + e.balance, 0)),
    customers: list,
    payments: paymentsRows_().reverse().slice(0, 100).map(p => Object.assign(p, { allocations: allocs.filter(a => a.paymentId === p.PaymentID).map(a => ({ saleId: a.saleId, amount: a.amount })) }))
  };
}

/* ---------------- trace: follow any ID through the whole web of records ---------------- */

/** Admin: search a sale / payment / customer / event ID (or a customer name). Returns the record, everything linked to it and its timeline. */
function traceId(token, q) {
  session_(token, 'admin');
  q = clean_(q, 80);
  if (!q) throw new Error('Type an ID or a customer name');
  const ql = q.toLowerCase();
  const custs = customersAll_();
  const cname = id => (custs.find(c => c.id === id) || {}).name || id;
  const sales = groupSales_(salesRows_());
  const payments = paymentsRows_();
  const allocs = allocRows_();
  const events = ledgerRows_();
  const link = (id, label) => ({ type: idType_(id), id: id, label: label || id });
  const evFor = id => events.filter(e => e.ref === id || e.linked.indexOf(id) >= 0)
    .map(e => ({ id: e.id, time: e.time, actor: e.actor, type: e.type, summary: e.summary, ref: e.ref }));
  const saleLabel = g => g.time.slice(5) + ' · ' + g.total + ' · ' + (g.status === 'VOID' ? 'VOID' : g.payStatus || g.payment);
  const out = { query: q, found: true, kind: '', title: '', main: {}, linked: [], events: [] };

  const sale = sales.find(g => g.saleId.toLowerCase() === ql);
  const pay = payments.find(p => p.PaymentID.toLowerCase() === ql);
  const cust = custs.find(c => c.id.toLowerCase() === ql) || custs.find(c => c.key === key_(q));
  const ev = events.find(e => e.id.toLowerCase() === ql);
  const al = allocs.find(a => a.id.toLowerCase() === ql);

  if (sale) {
    out.kind = 'sale'; out.title = 'Sale ' + sale.saleId; out.main = sale;
    out.linked.push(Object.assign(link(sale.customerId, cname(sale.customerId) || sale.customer), { rel: 'customer' }));
    if (sale.replaces) out.linked.push(Object.assign(link(sale.replaces, 'Replaces this earlier sale'), { rel: 'replaces' }));
    sales.filter(g => g.replaces === sale.saleId).forEach(g => out.linked.push(Object.assign(link(g.saleId, 'Replaced by this sale'), { rel: 'replaced by' })));
    allocs.filter(a => a.saleId === sale.saleId).forEach(a => {
      const pm = payments.find(p => p.PaymentID === a.paymentId) || {};
      out.linked.push(Object.assign(link(a.paymentId, a.amount + ' paid' + (a.status === 'VOID' ? ' (cancelled)' : '') + ' · ' + (pm.Method || '')), { rel: 'payment' }));
    });
    out.events = evFor(sale.saleId);
  } else if (pay) {
    out.kind = 'payment'; out.title = 'Payment ' + pay.PaymentID; out.main = pay;
    out.linked.push(Object.assign(link(pay.CustomerID, pay.Customer), { rel: 'customer' }));
    allocs.filter(a => a.paymentId === pay.PaymentID).forEach(a => {
      const g = sales.find(x => x.saleId === a.saleId);
      out.linked.push(Object.assign(link(a.saleId, a.amount + ' allocated' + (a.status === 'VOID' ? ' (released)' : '') + (g ? ' · sale of ' + g.total : '')), { rel: 'paid invoice' }));
    });
    out.events = evFor(pay.PaymentID);
  } else if (cust) {
    const b = balances_()[cust.id] || { owed: 0, paid: 0, balance: 0 };
    out.kind = 'customer'; out.title = cust.name; out.main = Object.assign(pubCust_(cust), { balance: b.balance, owed: b.owed, paid: b.paid });
    sales.filter(g => g.customerId === cust.id).reverse().slice(0, 100).forEach(g => out.linked.push(Object.assign(link(g.saleId, saleLabel(g)), { rel: 'sale' })));
    payments.filter(p => p.CustomerID === cust.id).reverse().slice(0, 100).forEach(p => out.linked.push(Object.assign(link(p.PaymentID, p.Timestamp.slice(5) + ' · ' + p.Amount + ' · ' + p.Method + (p.Status === 'VOID' ? ' · VOID' : '')), { rel: 'payment' })));
    out.events = evFor(cust.id);
  } else if (ev) {
    out.kind = 'event'; out.title = 'Event ' + ev.id; out.main = { id: ev.id, time: ev.time, actor: ev.actor, type: ev.type, summary: ev.summary, data: ev.data, prevHash: ev.prev, hash: ev.hash };
    uniq_([ev.ref].concat(ev.linked)).forEach(id => out.linked.push(Object.assign(link(id, id === ev.ref ? 'Subject of this event' : ''), { rel: id === ev.ref ? 'subject' : 'linked' })));
    out.events = [{ id: ev.id, time: ev.time, actor: ev.actor, type: ev.type, summary: ev.summary, ref: ev.ref }];
  } else if (al) {
    out.kind = 'allocation'; out.title = 'Allocation ' + al.id; out.main = al;
    out.linked.push(Object.assign(link(al.paymentId, 'Payment'), { rel: 'payment' }), Object.assign(link(al.saleId, 'Invoice'), { rel: 'sale' }));
    out.events = evFor(al.paymentId);
  } else {
    out.found = false; out.kind = 'none'; out.title = 'No exact match';
    custs.filter(c => c.key.indexOf(ql) >= 0 || c.id.toLowerCase().indexOf(ql) >= 0).slice(0, 10)
      .forEach(c => out.linked.push(Object.assign(link(c.id, c.name), { rel: 'customer' })));
    sales.filter(g => g.saleId.toLowerCase().indexOf(ql) >= 0).slice(0, 10).forEach(g => out.linked.push(Object.assign(link(g.saleId, saleLabel(g)), { rel: 'sale' })));
  }
  return out;
}


/* ======================================================================
 * CONTACT THE DEVELOPER (admin only)
 *  - Sends an email to DEVELOPER.email from the owner's account (needs the
 *    "send email" permission the first time: re-run any function in the editor and approve).
 *  - Max 5 messages per hour.
 * ==================================================================== */

function sendDeveloperMessage(token, d) {
  session_(token, 'admin');
  d = d || {};
  if (!DEVELOPER.email) throw new Error('Developer email is not configured. Visit ' + DEVELOPER.site);
  const subject = clean_(d.subject, 120) || 'Support request';
  const message = String(d.message || '').trim().slice(0, 3000);
  const replyTo = clean_(d.replyTo, 120);
  if (message.length < 5) throw new Error('Please write your message');
  if (replyTo && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(replyTo)) throw new Error('Enter a valid email address');

  const cache = CacheService.getScriptCache();
  const sent = Number(cache.get('MAILS') || 0);
  if (sent >= 5) throw new Error('Too many messages. Try again in an hour.');
  cache.put('MAILS', String(sent + 1), 3600);

  const s = settings_();
  const body = message + '\n\n--\nCompany: ' + (s.CompanyName || '-') + '\nSpreadsheet: ' + ss_().getUrl() +
    '\nWeb app: ' + (ScriptApp.getService().getUrl() || '-');
  const opt = { to: DEVELOPER.email, subject: '[Sales Dashboard] ' + subject, body: body, name: s.CompanyName || 'Sales Dashboard' };
  if (replyTo) opt.replyTo = replyTo;
  MailApp.sendEmail(opt);
  return true;
}

/* ======================================================================
 * LICENSE CLIENT
 *  - Each install has a random INSTANCE_ID (script property, not copied with the code).
 *  - Asks the license server at most every LICENSE.recheckHours hours; if it cannot be
 *    reached the last known state is kept (a first install then gets a local 14-day trial).
 *  - Status trial/active = works. expired/blocked = login and every API call is refused.
 * ==================================================================== */

function instanceId_() {
  const p = PropertiesService.getScriptProperties();
  let id = p.getProperty('INSTANCE_ID');
  if (!id) { id = Utilities.getUuid(); p.setProperty('INSTANCE_ID', id); }
  return id;
}

function licState_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('LIC') || 'null'); } catch (e) { return null; }
}

function licenseCheck_(force) {
  if (!LICENSE.url) return null;
  let st = licState_();
  const now = Date.now();
  if (st && !force && now - st.checkedAt < LICENSE.recheckHours * 3600000) return st;
  try {
    let company = '', email = '';
    try { company = String(settings_().CompanyName || ''); } catch (e) {}
    try { email = Session.getEffectiveUser().getEmail(); } catch (e) {}
    const url = LICENSE.url + (LICENSE.url.indexOf('?') < 0 ? '?' : '&') + 'id=' + encodeURIComponent(instanceId_()) +
      '&company=' + encodeURIComponent(company) + '&url=' + encodeURIComponent(ScriptApp.getService().getUrl() || '') +
      '&email=' + encodeURIComponent(email) + '&v=1';
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    const txt = res.getContentText();
    let j;
    try { j = JSON.parse(txt); } catch (x) { throw new Error('Server reply is not JSON (HTTP ' + res.getResponseCode() + '): ' + txt.slice(0, 120).replace(/\s+/g, ' ')); }
    if (!j || !j.status) throw new Error('Server error: ' + txt.slice(0, 120));
    st = { status: j.status, until: j.until || '', message: j.message || '', checkedAt: now };
  } catch (e) {
    if (!st) st = { status: 'trial', until: Utilities.formatDate(new Date(now + 14 * 86400000), tz_(), 'yyyy-MM-dd'), message: '' };
    st.err = String((e && e.message) || e).slice(0, 200);   // why the license server could not be reached
    st.checkedAt = now - LICENSE.recheckHours * 3600000 + 900000;   // try again in ~15 minutes
  }
  PropertiesService.getScriptProperties().setProperty('LIC', JSON.stringify(st));
  return st;
}

function licenseInfo_() {
  if (!LICENSE.url) return { ok: true, status: 'active', daysLeft: null, message: '' };
  const st = licenseCheck_(false) || {};
  let status = String(st.status || 'trial').toLowerCase();
  const today = today_();
  if ((status === 'trial' || status === 'active') && st.until && st.until < today) status = 'expired';
  const ok = status === 'trial' || status === 'active';
  let daysLeft = null;
  if (status === 'trial' && st.until) daysLeft = Math.max(0, Math.round((new Date(st.until + 'T12:00:00Z') - new Date(today + 'T12:00:00Z')) / 86400000));
  return { ok: ok, status: status, daysLeft: daysLeft, message: String(st.message || '') };
}

function licenseGate_() {
  if (!licenseInfo_().ok) throw new Error('LICENSE_EXPIRED');
}

/** Run from the editor to see exactly what the license server answers. */
function testLicense() {
  const url = LICENSE.url + '?id=' + encodeURIComponent(instanceId_());
  try {
    const r = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    Logger.log('HTTP ' + r.getResponseCode() + ' | ' + r.getContentText().slice(0, 300).replace(/\s+/g, ' '));
  } catch (e) { Logger.log('FETCH ERROR: ' + e); }
  Logger.log('State: ' + JSON.stringify(licenseCheck_(true)));
}

/* ======================================================================
 * RESET DATA (admin)
 *  Clears everything that was entered: sales, payments, invoice allocations, customers and the
 *  audit ledger (optionally also products and salesmen). It NEVER touches company information,
 *  the Admin PIN, or the trial / licence state. A SYSTEM_RESET entry starts the new audit chain.
 *  Needs the Admin PIN and the word RESET.
 * ==================================================================== */

function adminResetData(token, opts) {
  session_(token, 'admin');
  opts = opts || {};
  if (String(opts.pin || '').trim() !== adminPin_()) throw new Error('Admin PIN is wrong');
  if (String(opts.confirm || '').trim().toUpperCase() !== 'RESET') throw new Error('Type RESET to confirm');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  const cleared = {};
  try {
    ensureSalesHeaders_();
    const targets = [
      ['Sales', sheet_(CFG.SHEET.SALES)], ['Payments', tab_('Payments', PAYMENT_HEADERS)], ['Allocations', tab_('Allocations', ALLOC_HEADERS)],
      ['Customers', tab_('Customers', CUSTOMER_HEADERS)], ['Ledger', ledgerSh_()]
    ];
    if (opts.products !== false) targets.push(['Products', sheet_(CFG.SHEET.PRODUCTS)]);
    if (opts.salesmen !== false) targets.push(['Salesmen', sheet_(CFG.SHEET.SALESMEN)]);
    targets.forEach(t => {
      const sh = t[1], last = sh.getLastRow();
      cleared[t[0]] = Math.max(0, last - 1);
      if (last > 1) {
        const rng = sh.getRange(2, 1, last - 1, Math.max(sh.getLastColumn(), 1));
        rng.clearContent();      // keeps the column formats (text columns stay text)
        rng.setBackground(null);
      }
    });
    clearTableCache_();
    ledger_('Admin', 'SYSTEM_RESET', 'RESET', [], 'Data reset by Admin', { cleared: cleared, productsKept: opts.products === false, salesmenKept: opts.salesmen === false });
  } finally {
    unlock_(lock);
  }
  return { ok: true, cleared: cleared };
}
