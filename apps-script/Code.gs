// CODE.GS VERSION 2.0 - 2026-11-02 (paste this whole file, then Deploy > Manage deployments > New version; then run authorizeMail once)
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
var LOGGER_PASSWORD = "2026";
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

// Run this ONCE from the editor (Run > authorizeMail > Allow). It also sends a test email to MANAGER_RESET_EMAIL.
function authorizeMail() {
  MailApp.sendEmail(MANAGER_RESET_EMAIL, "Mentari HSSE Tracker - email check", "Code reset emails will be sent to this address.");
}

function _sheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

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

    if (action === "submit") {
      var email = (body.email || "").trim().toLowerCase();
      if (LOGGER_WHITELIST.indexOf(email) === -1 || body.password !== LOGGER_PASSWORD) {
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
      if (LOGGER_WHITELIST.indexOf(gEmail) === -1 || body.password !== LOGGER_PASSWORD) {
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
