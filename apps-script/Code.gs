// CODE.GS VERSION 3.1 - 2026-11-14 (one-time-code sign-in for everyone who opens the tracker; code email marked no-reply; paste this whole file, then Deploy > Manage deployments > pencil > New version)
/**
 * Mentari HSSE Tracker — Google Sheets backend
 * ---------------------------------------------
 * Deploy this bound to a Google Sheet with two tabs: "Live" and "Pending".
 *
 * "Live" holds one row per approved month/project (the tracker's source of truth).
 * Header row (columns A-F):
 *   id | month | project | submittedBy | submittedAt | dataJson
 *
 * "Pending" holds one row PER WEEK submitted by a logger — several weeks can
 * accumulate for the same month/project before the manager reviews. When
 * approved, all matching weeks are summed into a single Live row.
 * Header row (columns A-G):
 *   id | month | project | weekStart | submittedBy | submittedAt | dataJson
 *
 * Setup:
 * 1. Create a new Google Sheet. Rename Sheet1 to "Live". Add a second tab named "Pending".
 * 2. Put the matching header row (above) in row 1 of each tab, exact spelling.
 * 3. Extensions > Apps Script. Delete any starter code, paste this whole file in.
 * 4. Update LOGGER_WHITELIST / LOGGER_PASSWORD / MANAGER_EMAIL below (the manager code itself is set with "Forgot code?")
 *    to match the tracker's values exactly.
 * 5. Deploy > New deployment > type "Web app".
 *      - Execute as: Me
 *      - Who has access: Anyone
 *    Click Deploy, authorize when prompted, and copy the Web App URL.
 * 6. Send that URL back — it gets pasted into the tracker as GOOGLE_SHEET_API_URL.
 *
 * Whenever you edit this file in the Apps Script editor, you must create a NEW
 * deployment version (Deploy > Manage deployments > pencil icon > New version)
 * for the changes to actually go live at the same URL.
 */

var LIVE_SHEET = "Live";
var PENDING_SHEET = "Pending";

var LOGGER_WHITELIST = [
  "hocine.boudida@ecooils.com.my",
  "agung.muliawan@ecooils.co.id",
  "ronore.lapat@ecooils.co.id"
];
var LOGGER_PASSWORD = "2026";   // legacy: works for a logger ONLY until their first "Forgot access code?" code is used; then it is switched off for that person
var MANAGER_EMAIL = "iggifernandes@gmail.com";
var MANAGER_RESET_EMAIL = "Ignatius.Ambrosio@ecooils.com.my";   // the "Forgot code?" 6-digit code is sent here
var LEGACY_MANAGER_PASSWORD = "0822";   // works ONLY until the first reset; once a new code is set this is switched off for good


// ---- Manager sign-in: salted-hash code kept in Script Properties; forgotten code is reset by an emailed 6-digit code ----

function _props() { return PropertiesService.getScriptProperties(); }
function _hash(salt, v) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ":" + v);
  return raw.map(function (b) { var h = (b < 0 ? b + 256 : b).toString(16); return h.length < 2 ? "0" + h : h; }).join("");
}
function _mgrLocked() { return Number(CacheService.getScriptCache().get("mgr_fail") || 0) >= 5; }
function _mgrPassOk(pw) {
  var cache = CacheService.getScriptCache();
  if (_mgrLocked()) return false;
  var p = _props(), stored = p.getProperty("MGR_PIN_HASH"), s = String(pw || ""), ok;
  if (stored) ok = _hash(p.getProperty("MGR_PIN_SALT"), s) === stored;
  else ok = s === LEGACY_MANAGER_PASSWORD;
  if (ok) cache.remove("mgr_fail");
  else cache.put("mgr_fail", String(Number(cache.get("mgr_fail") || 0) + 1), 900);
  return ok;
}
function _sixDigits() { return ("000000" + (parseInt(Utilities.getUuid().replace(/-/g, "").slice(0, 8), 16) % 1000000)).slice(-6); }

// ================= v3.0: viewer sign-in (everyone who opens the tracker) =================
// Only @ecooils.com.my / @ecooils.co.id addresses (plus the owner's own) can get a 6-digit code by email.
// Every code request, sign-in and page open is written to the "AccessLog" tab (only people who can edit this Sheet can see it).
// Viewers never see who else opened the tracker. The page itself is encrypted; the key is released only to a signed-in viewer.
var VIEW_DOMAINS = ["ecooils.com.my", "ecooils.co.id"];
var VIEW_EXTRA = ["iggifernandes@gmail.com"];   // owner's own address (remove if not wanted)
var SESSION_DAYS = 14;                          // a device stays signed in this long, then asks for a new code
var CODE_MINUTES = 10;                          // a code is valid this long
var APP_KEY_B64 = "";                           // filled in ONLY in the copy that is pasted into Apps Script (blank in the public repo copy)
var ACCESS_LOG = "AccessLog", ACCESS_SUMMARY = "AccessSummary", BLOCKED_SHEET = "Blocked";

function _normEmail(v) { return String(v || "").trim().toLowerCase(); }
function _allowedEmail(em) {
  if (!/^[a-z0-9._%+'\-]{1,64}@[a-z0-9.\-]{3,80}$/.test(em)) return false;
  if (VIEW_EXTRA.indexOf(em) !== -1) return true;
  return VIEW_DOMAINS.indexOf(em.split("@")[1]) !== -1;
}
function _safeEq(a, b) {
  a = String(a); b = String(b);
  if (a.length !== b.length) return false;
  var d = 0;
  for (var i = 0; i < a.length; i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return d === 0;
}
function _secret() {
  var p = _props(), s = p.getProperty("VIEW_SECRET");
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid(); p.setProperty("VIEW_SECRET", s); }
  return s;
}
function _epoch() { return _props().getProperty("VIEW_EPOCH") || "1"; }
function _sign(payload) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(payload, _secret())); }
function _mkToken(email) {
  var payload = email + "|" + (Date.now() + SESSION_DAYS * 86400000) + "|" + _epoch();
  return Utilities.base64EncodeWebSafe(Utilities.newBlob(payload).getBytes()) + "." + _sign(payload);
}
function _readToken(tok) {   // -> the signed-in email, or "" when the token is missing, tampered, expired, revoked or blocked
  try {
    var parts = String(tok || "").split(".");
    if (parts.length !== 2) return "";
    var payload = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString();
    if (!_safeEq(_sign(payload), parts[1])) return "";
    var f = payload.split("|");
    if (f.length !== 3 || Number(f[1]) < Date.now() || f[2] !== _epoch()) return "";
    if (!_allowedEmail(f[0]) || _blocked(f[0])) return "";
    return f[0];
  } catch (e) { return ""; }
}
function _blocked(email) {
  var cache = CacheService.getScriptCache(), raw = cache.get("blk");
  if (raw === null || raw === undefined) {
    var list = [];
    try {
      var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BLOCKED_SHEET);
      if (sh && sh.getLastRow() > 1) list = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(function (r) { return _normEmail(r[0]); }).filter(String);
    } catch (e) {}
    raw = JSON.stringify(list);
    cache.put("blk", raw, 60);
  }
  return JSON.parse(raw).indexOf(email) !== -1;
}
function _gateOn() { return _props().getProperty("GATE_ON") === "1"; }

// ---- access log (Time | Email | Event | Device / note) ----
function _logSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(ACCESS_LOG);
  if (sh) return sh;
  sh = ss.insertSheet(ACCESS_LOG);
  sh.getRange(1, 1, 1, 4).setValues([["Time", "Email", "Event", "Device / note"]]).setFontWeight("bold");
  sh.setFrozenRows(1);
  sh.getRange("A2:A").setNumberFormat("yyyy-mm-dd hh:mm:ss");
  sh.setColumnWidth(1, 150); sh.setColumnWidth(2, 260); sh.setColumnWidth(3, 120); sh.setColumnWidth(4, 360);
  if (!ss.getSheetByName(BLOCKED_SHEET)) {
    var bl = ss.insertSheet(BLOCKED_SHEET);
    bl.getRange(1, 1).setValue("Email to block (one per row; takes effect within a minute)").setFontWeight("bold");
    bl.setColumnWidth(1, 420);
  }
  if (!ss.getSheetByName(ACCESS_SUMMARY)) {
    var sm = ss.insertSheet(ACCESS_SUMMARY);
    sm.getRange(1, 1, 1, 4).setValues([["Email", "Times opened", "Last access", "First access"]]).setFontWeight("bold");
    sm.getRange("A2").setFormula('=IFERROR(QUERY(AccessLog!A2:D,"select B, count(B), max(A), min(A) where C = \'OPEN\' or C = \'LOGIN_OK\' group by B order by max(A) desc label B \'\', count(B) \'\', max(A) \'\', min(A) \'\'",0),"")');
    sm.getRange("C2:D").setNumberFormat("yyyy-mm-dd hh:mm:ss");
    sm.setColumnWidth(1, 300); sm.setColumnWidth(2, 110); sm.setColumnWidth(3, 160); sm.setColumnWidth(4, 160);
  }
  return sh;
}
function _log(email, event, note) {
  try { _logSheet().appendRow([new Date(), email || "", event, String(note || "").slice(0, 160)]); } catch (e) {}
}
// rejected addresses can be typed by anyone: keep the log from being flooded
function _logRejected(email, note) {
  var cache = CacheService.getScriptCache(), n = Number(cache.get("rej_n") || 0);
  if (n >= 100) return;
  cache.put("rej_n", String(n + 1), 3600);
  _log(String(email || "").slice(0, 80), "REJECTED", note);
}

// ---- sign-in steps ----
function _viewRequest(email, ua) {
  email = _normEmail(email);
  if (!_allowedEmail(email)) { _logRejected(email, "not an allowed address"); return { ok: false, error: "Only @ecooils.com.my and @ecooils.co.id addresses can open this tracker." }; }
  if (!APP_KEY_B64) return { ok: false, error: "The tracker is not fully set up yet (server key missing)." };
  if (_blocked(email)) { _log(email, "BLOCKED", "code request"); return { ok: true }; }   // same reply as a normal send
  var cache = CacheService.getScriptCache(), rk = "otp_rl_" + email, ck = "otp_cd_" + email, n = Number(cache.get(rk) || 0);
  if (cache.get(ck)) return { ok: true, again: true };
  if (n >= 5) { _log(email, "RATE_LIMIT", "too many codes in an hour"); return { ok: false, error: "Too many codes requested. Please try again in an hour." }; }
  if (MailApp.getRemainingDailyQuota() < 3) { _log(email, "MAIL_QUOTA", "daily email limit reached"); return { ok: false, error: "The daily email limit has been reached. Please try again tomorrow." }; }
  var code = _sixDigits(), salt = Utilities.getUuid();
  cache.put("otp_" + email, JSON.stringify({ h: _hash(salt, code), s: salt, e: Date.now() + CODE_MINUTES * 60000, t: 0 }), CODE_MINUTES * 60);
  cache.put(ck, "1", 60);
  cache.put(rk, String(n + 1), 3600);
  try {
    MailApp.sendEmail({
      to: email,
      subject: "HSSE Performance Tracker - your access code " + code,
      body: "Your access code is " + code + "\n\nIt is valid for " + CODE_MINUTES + " minutes and can be used once. Enter it on the tracker sign-in page.\n\nIf you did not ask for this code, ignore this email - nobody can open the tracker without it.\n\nThis is an automatic message. Please do not reply - this mailbox is not monitored.",
      name: "EcoOils HSSE Tracker (no reply)",
      noReply: true
    });
  } catch (mailErr) {
    cache.remove("otp_" + email); cache.remove(ck);
    _log(email, "MAIL_FAIL", String(mailErr).slice(0, 120));
    return { ok: false, error: "The email could not be sent. Please try again." };
  }
  _log(email, "CODE_SENT", ua);
  return { ok: true };
}
function _viewVerify(email, code, ua) {
  email = _normEmail(email);
  var bad = { ok: false, error: "Code expired or incorrect. Request a new one." }, cache = CacheService.getScriptCache(), k = "otp_" + email;
  if (!_allowedEmail(email) || _blocked(email) || !APP_KEY_B64) { _logRejected(email, "sign-in attempt"); return bad; }
  var raw = cache.get(k);
  if (!raw) { _log(email, "LOGIN_FAIL", "no active code"); return bad; }
  var o = JSON.parse(raw);
  if (Date.now() > o.e || o.t >= 5) { cache.remove(k); _log(email, "LOGIN_FAIL", "code expired or used up"); return bad; }
  o.t++;
  if (!_safeEq(_hash(o.s, String(code || "").replace(/\D/g, "")), o.h)) {
    cache.put(k, JSON.stringify(o), Math.max(1, Math.ceil((o.e - Date.now()) / 1000)));
    _log(email, "LOGIN_FAIL", "wrong code (" + o.t + "/5)");
    return bad;
  }
  cache.remove(k);
  _log(email, "LOGIN_OK", ua);
  return { ok: true, token: _mkToken(email), key: APP_KEY_B64 };
}
function _viewOpen(token, ua) {
  var em = _readToken(token);
  if (!em) return { ok: false, auth: false, error: "Sign in required" };
  if (!APP_KEY_B64) return { ok: false, error: "The tracker is not fully set up yet (server key missing)." };
  _log(em, "OPEN", ua);
  return { ok: true, key: APP_KEY_B64 };
}

// ---- Logger access codes: one salted-hash code per logger. "Forgot access code?" emails a NEW code to the manager (MANAGER_RESET_EMAIL), who passes it on. ----
function _loggerPassOk(email, pw) {
  email = String(email || "").trim().toLowerCase();
  if (LOGGER_WHITELIST.indexOf(email) === -1) return false;
  var cache = CacheService.getScriptCache(), fk = "lg_fail_" + email;
  if (Number(cache.get(fk) || 0) >= 5) return false;
  var p = _props(), s = String(pw || ""), cur = p.getProperty("LG_HASH_" + email), nw = p.getProperty("LG_NEW_" + email), ok = false;
  if (cur && _hash(p.getProperty("LG_SALT_" + email), s) === cur) ok = true;
  else if (nw && _hash(p.getProperty("LG_NSALT_" + email), s) === nw) {
    p.setProperties({ ["LG_SALT_" + email]: p.getProperty("LG_NSALT_" + email), ["LG_HASH_" + email]: nw });
    p.deleteProperty("LG_NEW_" + email); p.deleteProperty("LG_NSALT_" + email);
    ok = true;
  } else if (!cur && s === LOGGER_PASSWORD) ok = true;
  if (ok) cache.remove(fk); else cache.put(fk, String(Number(cache.get(fk) || 0) + 1), 900);
  return ok;
}

// ---- Sheet menu (visible only to people who can edit this Sheet): set a code without email. The code is typed into a pop-up, stored only as a salted hash, never in the code or the Sheet. ----
function onOpen() {
  SpreadsheetApp.getUi().createMenu("HSSE Tracker")
    .addItem("Set manager code", "menuSetManagerCode")
    .addItem("Set a logger's access code", "menuSetLoggerCode")
    .addSeparator()
    .addItem("Require sign-in code: turn ON", "menuGateOn")
    .addItem("Require sign-in code: turn OFF", "menuGateOff")
    .addItem("Sign everyone out", "menuSignOutAll")
    .addItem("Open access log", "menuOpenLog")
    .addToUi();
}
function _askCode(ui, title, text) {
  var r = ui.prompt(title, text, ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return null;
  var v = String(r.getResponseText() || "").trim();
  if (!/^\d{4}$/.test(v)) { ui.alert("The code must be exactly 4 digits."); return null; }
  return v;
}
function menuSetManagerCode() {
  var ui = SpreadsheetApp.getUi(), v = _askCode(ui, "Set manager code", "Enter a new 4-digit code:");
  if (!v) return;
  var salt = Utilities.getUuid();
  _props().setProperties({ MGR_PIN_SALT: salt, MGR_PIN_HASH: _hash(salt, v) });
  CacheService.getScriptCache().remove("mgr_fail");
  ui.alert("Manager code updated. Use it to sign in to the tracker.");
}
function menuSetLoggerCode() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt("Set a logger's access code", "Logger email (" + LOGGER_WHITELIST.join(", ") + "):", ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var em = String(r.getResponseText() || "").trim().toLowerCase();
  if (LOGGER_WHITELIST.indexOf(em) === -1) { ui.alert("That email is not on the logger list."); return; }
  var v = _askCode(ui, "Set access code", "Enter a new 4-digit code for " + em + ":");
  if (!v) return;
  var salt = Utilities.getUuid(), o = {};
  o["LG_SALT_" + em] = salt; o["LG_HASH_" + em] = _hash(salt, v);
  var p = _props(); p.setProperties(o); p.deleteProperty("LG_NEW_" + em); p.deleteProperty("LG_NSALT_" + em);
  CacheService.getScriptCache().remove("lg_fail_" + em);
  ui.alert("Access code updated for " + em + ".");
}

function menuGateOn() {
  _props().setProperty("GATE_ON", "1"); _logSheet();
  SpreadsheetApp.getUi().alert("Sign-in code is now REQUIRED for all tracker data.");
}
function menuGateOff() {
  _props().setProperty("GATE_ON", "0");
  SpreadsheetApp.getUi().alert("Sign-in code is now OFF for tracker data (the old, open page works again).");
}
function menuSignOutAll() {
  _props().setProperty("VIEW_EPOCH", String(Number(_epoch()) + 1));
  _log("(all)", "SIGNED_OUT", "everyone signed out from the menu");
  SpreadsheetApp.getUi().alert("Everyone has been signed out. Each person needs a new code next time.");
}
function menuOpenLog() { var sh = _logSheet(); SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sh); }

// Run this ONCE from the editor (Run > authorizeMail > Allow). It also sends a test email to MANAGER_RESET_EMAIL.
function authorizeMail() {
  MailApp.sendEmail(MANAGER_RESET_EMAIL, "Mentari HSSE Tracker - email check", "Code reset emails will be sent to this address.");
}

// v2.1: the Live and Pending tabs (and their header rows) are created automatically the first time they are needed.
var SHEET_HEADERS = {
  "Live": ["id", "month", "project", "submittedBy", "submittedAt", "dataJson"],
  "Pending": ["id", "month", "project", "weekStart", "submittedBy", "submittedAt", "dataJson"]
};
function _sheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(name);
  if (!sh && SHEET_HEADERS[name]) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, SHEET_HEADERS[name].length).setValues([SHEET_HEADERS[name]]);
    sh.setFrozenRows(1);
  }
  return sh;
}
// Optional: run once from the editor to create both tabs straight away.
function setup() { _sheet(LIVE_SHEET); _sheet(PENDING_SHEET); _logSheet(); _secret(); }

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function _tz() { return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(); }
// Sheets may silently turn text like "2026-09" or "2026-09-27" into real dates; always read them back as text.
function _monthStr(v) { return v instanceof Date ? Utilities.formatDate(v, _tz(), "yyyy-MM") : String(v); }
function _dayStr(v) { return v instanceof Date ? Utilities.formatDate(v, _tz(), "yyyy-MM-dd") : String(v); }

// ---- Live sheet helpers (6 columns: id, month, project, submittedBy, submittedAt, dataJson) ----

function _liveRowsToRecords(sheet) {
  var values = sheet.getDataRange().getValues();
  var records = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (!row[0]) continue;
    var data = {};
    try { data = JSON.parse(row[5] || "{}"); } catch (e) { data = {}; }
    records.push({ id: row[0], month: _monthStr(row[1]), project: row[2], submittedBy: row[3], submittedAt: row[4], data: data });
  }
  return records;
}

function _findLiveRowByMonthProject(sheet, month, project) {
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (_monthStr(values[i][1]) === month && values[i][2] === project) return i + 1;
  }
  return -1;
}

// ---- Pending sheet helpers (7 columns: id, month, project, weekStart, submittedBy, submittedAt, dataJson) ----

function _pendingRowsToRecords(sheet) {
  var values = sheet.getDataRange().getValues();
  var records = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (!row[0]) continue;
    var data = {};
    try { data = JSON.parse(row[6] || "{}"); } catch (e) { data = {}; }
    records.push({ id: row[0], month: _monthStr(row[1]), project: row[2], weekStart: _dayStr(row[3]), submittedBy: row[4], submittedAt: row[5], data: data });
  }
  return records;
}

function _findPendingRowByPeriod(sheet, month, project, weekStart) {
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (_monthStr(values[i][1]) === month && values[i][2] === project && _dayStr(values[i][3]) === weekStart) return i + 1;
  }
  return -1;
}

// ---- Merge a list of weekly data objects into one combined month total ----

function _sumDataObjects(list, month, project) {
  var merged = { month: month, project: project, site: {}, offsite: {}, staffPermanent: { site: 0, offsite: 0 }, indicators: {}, total: 0 };
  list.forEach(function (d) {
    Object.keys(d.site || {}).forEach(function (k) {
      merged.site[k] = (merged.site[k] || 0) + (Number(d.site[k]) || 0);
    });
    Object.keys(d.offsite || {}).forEach(function (k) {
      merged.offsite[k] = (merged.offsite[k] || 0) + (Number(d.offsite[k]) || 0);
    });
    merged.staffPermanent.site += Number(d.staffPermanent && d.staffPermanent.site) || 0;
    merged.staffPermanent.offsite += Number(d.staffPermanent && d.staffPermanent.offsite) || 0;
    Object.keys(d.indicators || {}).forEach(function (k) {
      if (!merged.indicators[k]) merged.indicators[k] = { site: 0, offsite: 0 };
      merged.indicators[k].site += Number(d.indicators[k].site) || 0;
      merged.indicators[k].offsite += Number(d.indicators[k].offsite) || 0;
    });
  });
  var totalSite = 0, totalOffsite = 0;
  Object.keys(merged.site).forEach(function (k) { totalSite += merged.site[k]; });
  Object.keys(merged.offsite).forEach(function (k) { totalOffsite += merged.offsite[k]; });
  merged.total = totalSite + totalOffsite + merged.staffPermanent.site + merged.staffPermanent.offsite;
  return merged;
}

function doGet(e) {
  try {
    if (_gateOn()) return _json({ ok: false, auth: false, error: "Sign in required" });
    var live = _sheet(LIVE_SHEET);
    var records = _liveRowsToRecords(live).map(function (r) { return r.data; });
    return _json({ ok: true, months: records });
  } catch (err) {
    return _json({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents || "{}");
    var action = body.action;

    if (action === "viewRequest") return _json(_viewRequest(body.email, body.ua));
    if (action === "viewVerify") return _json(_viewVerify(body.email, body.code, body.ua));
    if (action === "viewOpen") return _json(_viewOpen(body.t, body.ua));
    if (action === "viewSignout") { var so = _readToken(body.t); if (so) _log(so, "SIGNED_OUT", "by the user"); return _json({ ok: true }); }
    // everything below needs a signed-in viewer once the sign-in code is switched on
    if (_gateOn() && !_readToken(body.t)) return _json({ ok: false, auth: false, error: "Sign in required" });

    if (action === "liveMonths") {
      var liveSheet = _sheet(LIVE_SHEET);
      return _json({ ok: true, months: _liveRowsToRecords(liveSheet).map(function (r) { return r.data; }) });
    }

    if (action === "submit") {
      var email = (body.email || "").trim().toLowerCase();
      if (LOGGER_WHITELIST.indexOf(email) === -1 || !_loggerPassOk(email, body.password)) {
        return _json({ ok: false, error: "Invalid credentials" });
      }
      var pending = _sheet(PENDING_SHEET);
      var existingRow = _findPendingRowByPeriod(pending, body.data.month, body.data.project, body.weekStart);
      if (existingRow === -1) {
        var id = "sub-" + new Date().getTime();
        pending.appendRow([id, body.data.month, body.data.project, body.weekStart, email, new Date().toISOString(), JSON.stringify(body.data)]);
        return _json({ ok: true, id: id });
      } else {
        var existingId = pending.getRange(existingRow, 1).getValue();
        pending.getRange(existingRow, 1, 1, 7).setValues([[existingId, body.data.month, body.data.project, body.weekStart, email, new Date().toISOString(), JSON.stringify(body.data)]]);
        return _json({ ok: true, id: existingId });
      }
    }

    if (action === "getDraft") {
      var gEmail = (body.email || "").trim().toLowerCase();
      if (LOGGER_WHITELIST.indexOf(gEmail) === -1 || !_loggerPassOk(gEmail, body.password)) {
        return _json({ ok: false, error: "Invalid credentials" });
      }
      var pendingG = _sheet(PENDING_SHEET);
      var rowIdxG = _findPendingRowByPeriod(pendingG, body.month, body.project, body.weekStart);
      if (rowIdxG === -1) return _json({ ok: true, draft: null });
      var rowG = pendingG.getRange(rowIdxG, 1, 1, 7).getValues()[0];
      var draftData = {};
      try { draftData = JSON.parse(rowG[6] || "{}"); } catch (e) {}
      return _json({ ok: true, draft: draftData, submittedBy: rowG[4], submittedAt: rowG[5] });
    }

    if (action === "listPending") {
      var mEmail = (body.email || "").trim().toLowerCase();
      if (mEmail !== MANAGER_EMAIL || !_mgrPassOk(body.password)) {
        return _json({ ok: false, error: "Invalid credentials" });
      }
      var pendingSheet = _sheet(PENDING_SHEET);
      var raw = _pendingRowsToRecords(pendingSheet);
      var groups = {};
      raw.forEach(function (r) {
        var key = r.month + "::" + r.project;
        if (!groups[key]) groups[key] = { month: r.month, project: r.project, weeks: [] };
        groups[key].weeks.push(r);
      });
      var result = Object.keys(groups).map(function (key) {
        var g = groups[key];
        var merged = _sumDataObjects(g.weeks.map(function (w) { return w.data; }), g.month, g.project);
        var latestAt = g.weeks.reduce(function (a, w) { return w.submittedAt > a ? w.submittedAt : a; }, "");
        return {
          month: g.month,
          project: g.project,
          data: merged,
          submittedAt: latestAt,
          weeks: g.weeks.map(function (w) { return { weekStart: w.weekStart, submittedBy: w.submittedBy, submittedAt: w.submittedAt }; })
        };
      });
      return _json({ ok: true, submissions: result });
    }

    if (action === "upsertLive") {
      var uEmail = (body.email || "").trim().toLowerCase();
      if (uEmail !== MANAGER_EMAIL || !_mgrPassOk(body.password)) {
        return _json({ ok: false, error: "Invalid credentials" });
      }
      var liveSheetU = _sheet(LIVE_SHEET);
      var existingRowU = _findLiveRowByMonthProject(liveSheetU, body.data.month, body.data.project);
      if (existingRowU !== -1 && !body.data.approvedWeeks) {
        try { var prevU = JSON.parse(liveSheetU.getRange(existingRowU, 6).getValue() || "{}"); if (prevU.approvedWeeks) body.data.approvedWeeks = prevU.approvedWeeks; } catch (e) {}
      }
      var newRow = ["live-" + new Date().getTime(), body.data.month, body.data.project, uEmail, new Date().toISOString(), JSON.stringify(body.data)];
      if (existingRowU === -1) {
        liveSheetU.appendRow(newRow);
      } else {
        var oldId = liveSheetU.getRange(existingRowU, 1).getValue();
        newRow[0] = oldId;
        liveSheetU.getRange(existingRowU, 1, 1, 6).setValues([newRow]);
      }
      return _json({ ok: true });
    }

    if (action === "approve" || action === "reject") {
      var mEmail2 = (body.email || "").trim().toLowerCase();
      if (mEmail2 !== MANAGER_EMAIL || !_mgrPassOk(body.password)) {
        return _json({ ok: false, error: "Invalid credentials" });
      }
      var pendingSheet2 = _sheet(PENDING_SHEET);
      var values2 = pendingSheet2.getDataRange().getValues();
      var matchingRows = []; // collected in descending order so deleteRow() indices stay valid
      for (var i = values2.length - 1; i >= 1; i--) {
        if (_monthStr(values2[i][1]) === body.month && values2[i][2] === body.project) matchingRows.push(i + 1);
      }
      if (matchingRows.length === 0) return _json({ ok: false, error: "No matching submissions found" });

      if (action === "approve") {
        var dataList = matchingRows.map(function (rowNum) {
          var row = pendingSheet2.getRange(rowNum, 1, 1, 7).getValues()[0];
          var d = {};
          try { d = JSON.parse(row[6] || "{}"); } catch (e) {}
          return d;
        });
        var weekStarts = matchingRows.map(function (rowNum) { return _dayStr(pendingSheet2.getRange(rowNum, 4).getValue()); });
        var merged = _sumDataObjects(dataList, body.month, body.project);
        var liveSheet = _sheet(LIVE_SHEET);
        var existingLiveRow = _findLiveRowByMonthProject(liveSheet, body.month, body.project);
        var approvedWeeks = weekStarts;
        if (existingLiveRow !== -1) {
          var oldData = {};
          try { oldData = JSON.parse(liveSheet.getRange(existingLiveRow, 6).getValue() || "{}"); } catch (e) {}
          merged = _sumDataObjects([oldData, merged], body.month, body.project);
          approvedWeeks = (oldData.approvedWeeks || []).concat(weekStarts);
        }
        merged.approvedWeeks = approvedWeeks.filter(function (w, i) { return approvedWeeks.indexOf(w) === i; }).sort();
        var newRow2 = ["live-" + new Date().getTime(), body.month, body.project, mEmail2, new Date().toISOString(), JSON.stringify(merged)];
        if (existingLiveRow === -1) {
          liveSheet.appendRow(newRow2);
        } else {
          var oldId2 = liveSheet.getRange(existingLiveRow, 1).getValue();
          newRow2[0] = oldId2;
          liveSheet.getRange(existingLiveRow, 1, 1, 6).setValues([newRow2]);
        }
      }

      matchingRows.forEach(function (rowNum) { pendingSheet2.deleteRow(rowNum); });
      return _json({ ok: true });
    }

    if (action === "loggerCheck") {
      var lce = (body.email || "").trim().toLowerCase();
      return _json({ ok: _loggerPassOk(lce, body.password) });
    }

    if (action === "loggerForgot") {
      var lfe = (body.email || "").trim().toLowerCase();
      if (LOGGER_WHITELIST.indexOf(lfe) !== -1) {
        var pl = _props(), nowL = Date.now();
        if (nowL - Number(pl.getProperty("LG_SENT_" + lfe) || 0) >= 120000) {
          var lcode = ("0000" + (parseInt(Utilities.getUuid().replace(/-/g, "").slice(0, 8), 16) % 10000)).slice(-4), lsalt = Utilities.getUuid(), setObj = {};
          setObj["LG_NSALT_" + lfe] = lsalt; setObj["LG_NEW_" + lfe] = _hash(lsalt, lcode); setObj["LG_SENT_" + lfe] = String(nowL);
          pl.setProperties(setObj);
          try {
            MailApp.sendEmail(MANAGER_RESET_EMAIL, "Mentari HSSE Tracker - new access code for " + lfe,
              lfe + " asked for a new Log month access code.\n\nNew code: " + lcode + "\n\nPass it on to them. It switches on the first time it is used; their old code keeps working until then.\nIf this request was not expected, ignore this email.");
          } catch (mailErr2) {
            return _json({ ok: false, error: "The email could not be sent. Run authorizeMail once in the Apps Script editor." });
          }
        }
      }
      return _json({ ok: true });   // same reply whatever was typed
    }

    if (action === "mgrCheck") {
      var ce = (body.email || "").trim().toLowerCase();
      if (ce !== MANAGER_EMAIL) return _json({ ok: false });
      var okc = _mgrPassOk(body.password);
      return _json(okc ? { ok: true } : { ok: false, locked: _mgrLocked() });
    }

    if (action === "mgrForgot") {
      var fe = (body.email || "").trim().toLowerCase();
      if (fe === MANAGER_EMAIL) {
        var pf = _props(), nowF = Date.now();
        if (nowF - Number(pf.getProperty("MGR_CODE_SENT") || 0) >= 60000) {
          var code = _sixDigits(), csalt = Utilities.getUuid();
          pf.setProperties({ MGR_CODE_SALT: csalt, MGR_CODE_HASH: _hash(csalt, code), MGR_CODE_EXP: String(nowF + 15 * 60000), MGR_CODE_TRIES: "0", MGR_CODE_SENT: String(nowF) });
          try {
            MailApp.sendEmail(MANAGER_RESET_EMAIL, "Mentari HSSE Tracker - verification code",
              "Your verification code is " + code + "\n\nIt is valid for 15 minutes. Enter it on the sign-in screen, then choose a new 4-digit code.\nIf you did not ask for this, ignore this email.");
          } catch (mailErr) {
            return _json({ ok: false, error: "The email could not be sent. Run authorizeMail once in the Apps Script editor." });
          }
        }
      }
      return _json({ ok: true });   // same reply whatever was typed
    }

    if (action === "mgrReset") {
      var re = (body.email || "").trim().toLowerCase(), pr = _props(), bad = "Code expired or incorrect. Request a new one.";
      var tries = Number(pr.getProperty("MGR_CODE_TRIES") || 0);
      if (re !== MANAGER_EMAIL || !pr.getProperty("MGR_CODE_HASH") || Date.now() > Number(pr.getProperty("MGR_CODE_EXP") || 0) || tries >= 5) return _json({ ok: false, error: bad });
      var np = String(body.newPin || "");
      if (!/^\d{4}$/.test(np)) return _json({ ok: false, error: "The new code must be exactly 4 digits." });
      pr.setProperty("MGR_CODE_TRIES", String(tries + 1));
      if (_hash(pr.getProperty("MGR_CODE_SALT"), String(body.code || "").trim()) !== pr.getProperty("MGR_CODE_HASH")) return _json({ ok: false, error: bad });
      var psalt = Utilities.getUuid();
      pr.setProperties({ MGR_PIN_SALT: psalt, MGR_PIN_HASH: _hash(psalt, np) });
      ["MGR_CODE_SALT", "MGR_CODE_HASH", "MGR_CODE_EXP", "MGR_CODE_TRIES"].forEach(function (k) { pr.deleteProperty(k); });
      CacheService.getScriptCache().remove("mgr_fail");
      return _json({ ok: true });
    }

    return _json({ ok: false, error: "Unknown action" });
  } catch (err) {
    return _json({ ok: false, error: String(err) });
  }
}
