// ============================================================================
// MASTER CONFIGURATION & IMMUTABLE ADMIN SAFEGUARDS (v5.4.0)
// ============================================================================
var APP_VERSION = "5.4.0";
var SCHEMA_VERSION = "5.4.0";
var PRIMARY_ADMIN_EMAIL = "pitchpulseleague@gmail.com";
var BACKUP_RECOVERY_EMAIL = "ralph@me.com";
var PROTECTED_ADMINS = [
  "pitchpulseleague@gmail.com",
  "ralph@me.com"
];
var SITE_URL = "https://pitchpulseleague.github.io/soccer/";
var MAIL_SENDER_NAME = "Pitch Pulse League";   // the "From" name players see on emails

// ---- Notification volume controls -------------------------------------------
// A consumer Gmail account can email only 100 recipients per day (Workspace: 1,500).
// RSVP confirmations were ~40% of all email traffic, and players already see the
// result instantly in the app, so they go out as push only unless you flip this to true.
var EMAIL_RSVP_CONFIRMATIONS = false;
// Always keep this many recipients in reserve for PIN-reset codes and invitations.
var EMAIL_QUOTA_RESERVE = 10;

// ---- League rules -------------------------------------------------------------
var MAX_GUESTS_PER_PLAYER = 3;
var LEAGUE_TIME_ZONE = "America/Los_Angeles";
var MATCH_DURATION_MINUTES = 75;
var SERIES_WINDOW_DAYS = 28;             // recurring matches appear (and open for RSVP) four weeks ahead, one new day at a time
var PAST_MATCHES_IN_APP = 20;            // finished matches each phone downloads; standings still count every match
var LATE_DROP_WINDOW_MINUTES = 180;     // dropping a YES within 3h of kickoff counts as a late drop
var SMALL_SIDED_MIN_PLAYERS = 10;       // a small-sided match needs at least 5 v 5
var SESSION_DAYS = 180;                  // sliding: renewed automatically while the player keeps using the app
var MAX_SESSIONS_PER_USER = 5;           // phone + laptop + tablet... each device keeps its own login
var PIN_RESET_COOLDOWN_SECONDS = 60;
var PIN_RESET_MAX_PER_DAY = 5;

// League-wide defaults. Admins change these in Admin → Settings (stored in Script Properties).
// The core (In / Maybe / Out, guests, waitlist, directions) is always on. Everything else is a
// switch, and a new league starts with only the core.
var DEFAULT_LEAGUE_SETTINGS = {
  teamAName: "White",
  teamAShirt: "white",
  teamBName: "Dark",
  teamBShirt: "black",
  defaultMin: 14,
  defaultMax: 18,
  maxGuests: 3,
  // Match features (each match keeps the set that was on when it kicked off)
  featTeams: false,
  featScores: false,          // needs Teams
  featMotm: false,            // needs Scores
  featAttendance: false,
  featLateDrops: false,
  trackBall: false,
  trackPinnies: false,
  trackGK: false,
  featArrival: false,
  featCaptains: false,
  featAnnouncements: true,
  featSmallSided: false,      // organizers may start a match short of its minimum as 5 v 5 or more
  // Extras on the card and header
  featWeather: true,
  featCalendar: true,
  featRules: true,
  // Automatic messages
  msgDayBefore: true,         // only when the match is still short of its minimum
  msgMatchDay: true,          // around 8:30 AM; respects channel preferences
  msgGameOn: true,            // push when a match reaches its minimum
  msgConfirm: false,           // push confirming "you're in"
  msgMaybeMorning: true,      // match-morning push to players still on Maybe
  msgTwoHour: false,          // push about two hours before kickoff
  featuresVersion: 1
};
var BOOLEAN_SETTINGS = ["featTeams", "featScores", "featMotm", "featAttendance", "featLateDrops", "trackBall", "trackPinnies", "trackGK",
  "featArrival", "featCaptains", "featAnnouncements", "featSmallSided", "featWeather", "featCalendar", "featRules",
  "msgDayBefore", "msgMatchDay", "msgGameOn", "msgConfirm", "msgMaybeMorning", "msgTwoHour"];
var SHIRT_COLORS = ["white", "black", "gray", "red", "blue", "navy", "green", "yellow", "orange", "purple"];

var DEFAULT_HOW_WE_PLAY_RULES =
  "• Bring both a White shirt and a Dark shirt (or use league pinnies if enabled).\n" +
  "• Turf shoes or cleats allowed; shin guards optional.\n" +
  "• Respectful pickup play — call your own fouls, rotate in goalies fairly, and update your RSVP at least 3 hours before kickoff.";

var SS = SpreadsheetApp.getActiveSpreadsheet();

// Per-request state (Apps Script re-runs this file for every request, so these reset each time).
var __LANG = "en";
var __HAS_LOCK = false;   // true while this request holds the global script lock
var __TBL = {};
var __DEFER_NOTIFY = false;
var __NOTIFY_QUEUE = [];
var __PREF_MAP = null;
var __LANG_MAP = null;
var __PUSH_UNREACHED = {};   // players whose push couldn't be delivered during this request

function isProtectedAdmin(email) {
  var em = String(email || "").toLowerCase().trim();
  return PROTECTED_ADMINS.indexOf(em) !== -1;
}

function isNonPlayingSystemAdmin(email) {
  return String(email || "").toLowerCase().trim() === PRIMARY_ADMIN_EMAIL.toLowerCase().trim();
}

function isValidSixDigitPin(pin) {
  return /^\d{6}$/.test(String(pin || "").trim());
}

function normalizeNotifyPref(val) {
  var v = String(val || "Both").trim();
  if (v === "Push" || v === "Email" || v === "Both") return v;
  return "Both";
}

function normalizeLang(val) {
  return String(val || "").toLowerCase().trim() === "es" ? "es" : "en";
}

function cleanPersonName(val) {
  return String(val || "").replace(/[<>\u0000-\u001F]/g, "").replace(/\s+/g, " ").trim().slice(0, 40);
}

// ============================================================================
// HEADER-BASED TABLE ACCESS (immune to column order — fixes "ghost guests")
// ============================================================================
// Every read/write goes through the header row, so an extra/moved column in the
// Google Sheet can never make a write land in the wrong cell again.
var SCHEMA = {
  "Profiles": ["Email", "PinCode", "FullName", "JoinedDate", "Role", "AccountStatus", "Position", "LateDrops", "CanBringGuests", "AwayUntil", "LateEarlyCount", "NotifyPref", "DefaultTeam", "Lang"],
  "Games": ["GameID", "Title", "LocationName", "Date", "Time", "MaxCap", "MinThreshold", "Status", "SeriesID", "TeamsJSON", "ScoreWhite", "ScoreDark", "ScheduleChangeAt", "CaptainEmail", "MatchFormat", "ScheduleChangeSummary", "AdminActivityJSON", "NotifyStateJSON", "FeaturesJSON"],
  "Signups": ["SignupID", "GameID", "Email", "Status", "Timestamp", "BringBall", "BringPinnies", "PlayGK", "PlusOne", "MotmVote", "AvailabilityNote", "ReconfirmedAt", "AttendanceStatus", "AttendanceRecordedAt", "AttendanceRecordedBy", "LateDropAt"],
  "Locations": ["LocationID", "Name", "Address", "Latitude", "Longitude", "Notes", "LocationStatus"],
  "Invites": ["Email", "Status", "InvitedDate"],
  "Recurring": ["SeriesID", "Title", "LocationName", "DaysOfWeek", "Time", "StartDate", "EndDate", "MaxCap", "MinThreshold", "Status", "UpdatedAt", "GeneratedThrough", "ChangesJSON"],
  "Blasts": ["BlastID", "GameID", "SenderName", "SenderEmail", "Message", "Timestamp", "Hidden"],
  "Notifications": ["JobID", "DedupKey", "CreatedAt", "ExpiresAt", "Priority", "Recipient", "PayloadJSON", "EmailState", "PushState", "EmailAttempts", "PushAttempts", "NextAttemptAt", "LeaseToken", "LeaseUntil", "UpdatedAt", "LastError", "ProviderID"]
};

function T(name) {
  if (__TBL[name]) return __TBL[name];
  var sh = SS.getSheetByName(name);
  if (!sh) throw new Error("Missing sheet '" + name + "'. Run initialSetup() once from the Apps Script editor.");
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  var values = (lastRow > 0 && lastCol > 0) ? sh.getRange(1, 1, lastRow, lastCol).getValues() : [[]];
  var headers = (values[0] || []).map(function(h) { return String(h || "").trim(); });
  var col = {};
  headers.forEach(function(h, i) { if (h && col[h] === undefined) col[h] = i; });
  var t = { name: name, sheet: sh, headers: headers, col: col, rows: values.slice(1) };
  __TBL[name] = t;
  return t;
}

function tHas(t, h) { return t.col[h] !== undefined; }

function tv(t, r, h) {
  var i = t.col[h];
  if (i === undefined) return "";
  var v = t.rows[r][i];
  return v === undefined || v === null ? "" : v;
}

function memoValue(v) {
  if (typeof v === "string" && v.charAt(0) === "'") return v.slice(1);
  return v;
}

function tw(t, r, h, v) {
  var i = t.col[h];
  if (i === undefined) throw new Error("Column '" + h + "' is missing from sheet '" + t.name + "'. Run initialSetup().");
  t.sheet.getRange(r + 2, i + 1).setValue(v);
  t.rows[r][i] = memoValue(v);
}

// Writes several fields of one row with as few sheet calls as possible.
function twMany(t, r, obj) {
  var idxs = [];
  Object.keys(obj).forEach(function(h) {
    var i = t.col[h];
    if (i === undefined) throw new Error("Column '" + h + "' is missing from sheet '" + t.name + "'. Run initialSetup().");
    idxs.push({ i: i, v: obj[h] });
  });
  if (!idxs.length) return;
  idxs.sort(function(a, b) { return a.i - b.i; });
  // one setValues() call per run of adjacent columns
  var run = [idxs[0]];
  var flush = function(list) {
    t.sheet.getRange(r + 2, list[0].i + 1, 1, list.length).setValues([list.map(function(x) { return x.v; })]);
  };
  for (var k = 1; k < idxs.length; k++) {
    if (idxs[k].i === run[run.length - 1].i + 1) run.push(idxs[k]);
    else { flush(run); run = [idxs[k]]; }
  }
  flush(run);
  idxs.forEach(function(x) { t.rows[r][x.i] = memoValue(x.v); });
}

function tadd(t, obj) {
  var row = t.headers.map(function(h) { return (h && obj.hasOwnProperty(h)) ? obj[h] : ""; });
  t.sheet.appendRow(row);
  t.rows.push(row.map(memoValue));
  return t.rows.length - 1;
}

// Appends many rows with one write (far faster than one appendRow per row).
function taddMany(t, objs) {
  if (!objs || !objs.length) return;
  var rows = objs.map(function(obj) { return t.headers.map(function(h) { return (h && obj.hasOwnProperty(h)) ? obj[h] : ""; }); });
  t.sheet.getRange(t.sheet.getLastRow() + 1, 1, rows.length, t.headers.length).setValues(rows);
  rows.forEach(function(row) { t.rows.push(row.map(memoValue)); });
}

function tdel(t, r) {
  t.sheet.deleteRow(r + 2);
  t.rows.splice(r, 1);
}

function tfind(t, pred) {
  for (var r = 0; r < t.rows.length; r++) if (pred(r)) return r;
  return -1;
}

function tobj(t, r) {
  var o = {};
  t.headers.forEach(function(h, i) { if (h) o[h] = t.rows[r][i] === undefined ? "" : t.rows[r][i]; });
  return o;
}

function tobjs(t) {
  var out = [];
  for (var r = 0; r < t.rows.length; r++) {
    var row = t.rows[r];
    var empty = true;
    for (var c = 0; c < row.length; c++) { if (row[c] !== "" && row[c] !== null && row[c] !== undefined) { empty = false; break; } }
    if (!empty) out.push(tobj(t, r));
  }
  return out;
}

function findProfileRow(email) {
  var em = String(email || "").toLowerCase().trim();
  var t = T("Profiles");
  return tfind(t, function(r) { return String(tv(t, r, "Email")).toLowerCase().trim() === em; });
}

function findGameRow(gameId) {
  var t = T("Games");
  var id = String(gameId);
  return tfind(t, function(r) { return String(tv(t, r, "GameID")) === id; });
}

function findSignupRow(gameId, email) {
  var t = T("Signups");
  var id = String(gameId), em = String(email || "").toLowerCase().trim();
  return tfind(t, function(r) { return String(tv(t, r, "GameID")) === id && String(tv(t, r, "Email")).toLowerCase().trim() === em; });
}

function gameRecord(r) {
  var t = T("Games");
  var settings = getLeagueSettings();
  return {
    row: r,
    id: String(tv(t, r, "GameID")),
    title: String(tv(t, r, "Title")),
    location: String(tv(t, r, "LocationName")),
    date: formatDateIso(tv(t, r, "Date")),
    time: formatTimeStr(tv(t, r, "Time")),
    maxCap: Number(tv(t, r, "MaxCap") || settings.defaultMax),
    minThreshold: Number(tv(t, r, "MinThreshold") || settings.defaultMin),
    status: String(tv(t, r, "Status") || "Pending"),
    oldStatus: String(tv(t, r, "Status") || "Pending"),
    captainEmail: String(tv(t, r, "CaptainEmail") || "").toLowerCase().trim(),
    scoreWhite: tv(t, r, "ScoreWhite"),
    scoreDark: tv(t, r, "ScoreDark"),
    scheduleChangeAt: (function() { var v = tv(t, r, "ScheduleChangeAt"); return v ? Number(new Date(v).getTime() || v || 0) : 0; })(),
    features: gameFeaturesAt(t, r)
  };
}

function isoPlusDays(iso, n) {
  var p = String(iso).split("-");
  var d = new Date(Date.UTC(Number(p[0]), Number(p[1]) - 1, Number(p[2]) + n, 12));
  return Utilities.formatDate(d, "UTC", "yyyy-MM-dd");
}

function gameHasScore(gRec) {
  return gRec.scoreWhite !== "" && gRec.scoreWhite !== null && gRec.scoreDark !== "" && gRec.scoreDark !== null;
}

// Single source of truth for "can players still change their RSVP?" (Rule 2B:
// RSVPs stay open during the match and lock once the final score is saved or the day ends).
function gameLockState(gRec) {
  if (gRec.status === "Cancelled") return { locked: true, reason: "cancelled" };
  if (gRec.date < formatDateIso(new Date())) return { locked: true, reason: "past" };
  if ((!gRec.features || gRec.features.scores) && gameHasScore(gRec)) return { locked: true, reason: "scored" };
  return { locked: false, reason: "" };
}

function lockedError(reason) {
  if (reason === "cancelled") return tr("errCancelled");
  if (reason === "scored") return tr("errScored");
  return tr("errPast");
}

// ============================================================================
// LEAGUE SETTINGS (team names & shirts, default min/max, optional gear tracking)
// ============================================================================
var __SETTINGS = null;
function getLeagueSettings() {
  if (__SETTINGS) return __SETTINGS;
  var props = PropertiesService.getScriptProperties();
  var s = {};
  for (var k in DEFAULT_LEAGUE_SETTINGS) s[k] = DEFAULT_LEAGUE_SETTINGS[k];
  var legacyA = props.getProperty("TEAM_A_NAME");
  var legacyB = props.getProperty("TEAM_B_NAME");
  if (legacyA) s.teamAName = legacyA;
  if (legacyB) s.teamBName = legacyB;
  var legacyPinnies = props.getProperty("REQUIRE_PINNIES");
  if (legacyPinnies !== null && legacyPinnies !== undefined) s.trackPinnies = String(legacyPinnies).toLowerCase() !== "false";
  var raw = props.getProperty("LEAGUE_SETTINGS");
  if (raw) {
    try {
      var saved = JSON.parse(raw);
      for (var key in saved) if (DEFAULT_LEAGUE_SETTINGS.hasOwnProperty(key)) s[key] = saved[key];
    } catch (e) {}
  }
  s.defaultMin = Math.max(2, Math.min(60, Number(s.defaultMin) || DEFAULT_LEAGUE_SETTINGS.defaultMin));
  s.defaultMax = Math.max(s.defaultMin, Math.min(60, Number(s.defaultMax) || DEFAULT_LEAGUE_SETTINGS.defaultMax));
  BOOLEAN_SETTINGS.forEach(function(k) { s[k] = s[k] === true || String(s[k]).toLowerCase() === "true"; });
  // A switch that needs another one is off whenever its parent is off.
  if (!s.featTeams) s.featScores = false;
  if (!s.featScores) s.featMotm = false;
  var mg = Number(s.maxGuests);
  s.maxGuests = isFinite(mg) ? Math.max(0, Math.min(MAX_GUESTS_PER_PLAYER, Math.floor(mg))) : MAX_GUESTS_PER_PLAYER;
  if (SHIRT_COLORS.indexOf(s.teamAShirt) === -1) s.teamAShirt = "white";
  if (SHIRT_COLORS.indexOf(s.teamBShirt) === -1) s.teamBShirt = "black";
  s.teamAName = String(s.teamAName || "White").slice(0, 24);
  s.teamBName = String(s.teamBName || "Dark").slice(0, 24);
  __SETTINGS = s;
  return s;
}

function saveLeagueSettings(partial) {
  var cur = getLeagueSettings();
  var next = {};
  for (var k in DEFAULT_LEAGUE_SETTINGS) next[k] = cur[k];
  for (var p in partial) if (DEFAULT_LEAGUE_SETTINGS.hasOwnProperty(p) && partial[p] !== undefined && partial[p] !== null) next[p] = partial[p];
  var props = PropertiesService.getScriptProperties();
  props.setProperty("LEAGUE_SETTINGS", JSON.stringify(next));
  props.setProperty("REQUIRE_PINNIES", next.trackPinnies ? "true" : "false");
  __SETTINGS = null;
  return getLeagueSettings();
}

// Team names are free text; only the untouched defaults are translated for Spanish replies.
var DEFAULT_TEAM_NAMES_ES = { White: "Blanco", Dark: "Oscuro" };
function teamName(side, raw) {
  var s = getLeagueSettings();
  var n = side === "dark" ? s.teamBName : s.teamAName;
  if (!raw && __LANG === "es" && DEFAULT_TEAM_NAMES_ES.hasOwnProperty(n)) return DEFAULT_TEAM_NAMES_ES[n];
  return n;
}

// ============================================================================
// MATCH FEATURES — each match follows the league switches until kickoff, then keeps them
// ============================================================================
// Turning a switch on or off changes every match that hasn't kicked off yet. At kickoff a
// match's set is written into its row (FeaturesJSON), so finished matches never change and
// switching something off only hides it — nothing is deleted.
var FEATURE_FLAGS = ["teams", "scores", "motm", "attendance", "lateDrops", "ball", "pinnies", "gk", "arrival", "captains", "announcements"];
function featuresFromSettings(s) {
  s = s || getLeagueSettings();
  return {
    teams: !!s.featTeams, scores: !!(s.featTeams && s.featScores), motm: !!(s.featTeams && s.featScores && s.featMotm),
    attendance: !!s.featAttendance, lateDrops: !!s.featLateDrops,
    ball: !!s.trackBall, pinnies: !!s.trackPinnies, gk: !!s.trackGK,
    arrival: !!s.featArrival, captains: !!s.featCaptains, announcements: !!s.featAnnouncements
  };
}
function normalizeFeatures(o) {
  var f = {};
  FEATURE_FLAGS.forEach(function(k) { f[k] = !!(o && (o[k] === true || String(o[k]).toLowerCase() === "true")); });
  if (!f.teams) f.scores = false;
  if (!f.scores) f.motm = false;
  return f;
}
function gameFeaturesAt(t, r) {
  if (r !== -1 && r !== undefined && tHas(t, "FeaturesJSON")) {
    var raw = tv(t, r, "FeaturesJSON");
    if (raw) { try { return normalizeFeatures(JSON.parse(String(raw))); } catch (e) {} }
  }
  return featuresFromSettings();
}
function gameFeatures(gameId) {
  var t = T("Games");
  return gameFeaturesAt(t, findGameRow(gameId));
}
// Writes the current switches into every match that has kicked off but isn't locked yet.
// Runs every hour and always just before a switch changes, so a started match never picks
// up a later change.
function freezeStartedMatches(featuresOverride) {
  var t = T("Games");
  if (!tHas(t, "FeaturesJSON")) return 0;
  var json = JSON.stringify(featuresOverride || featuresFromSettings());
  var n = 0;
  for (var r = 0; r < t.rows.length; r++) {
    if (!String(tv(t, r, "GameID") || "")) continue;
    if (tv(t, r, "FeaturesJSON")) continue;
    if (!hasGameStartedOrPassed(formatDateIso(tv(t, r, "Date")), formatTimeStr(tv(t, r, "Time")))) continue;
    tw(t, r, "FeaturesJSON", json);
    n++;
  }
  return n;
}
// One-time move to switches (v5.2). Matches already played keep everything the older
// version showed for them; from now on the league starts with only the core switched on.
function migrateToFeatureSwitches() {
  var props = PropertiesService.getScriptProperties();
  var stored = {};
  try { stored = JSON.parse(props.getProperty("LEAGUE_SETTINGS") || "{}") || {}; } catch (e) { stored = {}; }
  if (Number(stored.featuresVersion || 0) >= 1) return false;
  var legacyOn = function(k) { return !(stored[k] === false || String(stored[k]).toLowerCase() === "false"); };
  var legacyPinnies = props.getProperty("REQUIRE_PINNIES");
  var legacy = { teams: true, scores: true, motm: true, attendance: true, lateDrops: true,
    ball: legacyOn("trackBall"), pinnies: legacyPinnies !== null && legacyPinnies !== undefined ? String(legacyPinnies).toLowerCase() !== "false" : legacyOn("trackPinnies"),
    gk: legacyOn("trackGK"), arrival: true, captains: true, announcements: true };
  __TBL = {};
  freezeStartedMatches(legacy);
  var next = {};
  for (var k in DEFAULT_LEAGUE_SETTINGS) next[k] = DEFAULT_LEAGUE_SETTINGS[k];
  ["teamAName", "teamAShirt", "teamBName", "teamBShirt", "defaultMin", "defaultMax"].forEach(function(key) {
    if (stored[key] !== undefined && stored[key] !== null && stored[key] !== "") next[key] = stored[key];
  });
  var legacyA = props.getProperty("TEAM_A_NAME"), legacyB = props.getProperty("TEAM_B_NAME");
  if (legacyA && stored.teamAName === undefined) next.teamAName = legacyA;
  if (legacyB && stored.teamBName === undefined) next.teamBName = legacyB;
  next.featuresVersion = 1;
  props.setProperty("LEAGUE_SETTINGS", JSON.stringify(next));
  props.setProperty("REQUIRE_PINNIES", next.trackPinnies ? "true" : "false");
  __SETTINGS = null;
  return true;
}

// ============================================================================
// 1-CLICK AUTO SETUP (Nondestructive, Idempotent & Preserves Unrelated Triggers)
// ============================================================================
// Upgrade helpers are deliberately quiet: they add schema and install triggers but do not
// send email/push, reset credentials, repair/delete records, or replace existing settings.
var LEAGUE_OWNED_TRIGGER_HANDLERS = ["send9amGameDayUpdate", "sendDayBeforeNonRsvpReminder", "sendTwoHourReminders", "leagueAutomationDispatcher"];
function pauseLeagueAutomations() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (LEAGUE_OWNED_TRIGGER_HANDLERS.indexOf(trigger.getHandlerFunction()) !== -1) { ScriptApp.deleteTrigger(trigger); removed++; }
  });
  return { paused: true, removedOwnedTriggers: removed };
}
function installProductionTriggers() {
  pauseLeagueAutomations();
  ScriptApp.newTrigger("leagueAutomationDispatcher").timeBased().everyMinutes(5).inTimezone("America/Los_Angeles").create();
  PropertiesService.getScriptProperties().setProperty("AUTOMATION_VERSION", APP_VERSION);
  return { installed: true, handler: "leagueAutomationDispatcher", approximateIntervalMinutes: 5, timeZone: "America/Los_Angeles" };
}
function upgradeToV540() {
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    __HAS_LOCK = true; __TBL = {};
    ensureSchema(true);
  } finally { __HAS_LOCK = false; lock.releaseLock(); }
  var triggers = installProductionTriggers();
  return { version: APP_VERSION, schemaVersion: SCHEMA_VERSION, triggers: triggers, message: "Upgrade ready. No messages were sent by setup; existing records, settings and credentials were preserved." };
}
function initialSetup() { return upgradeToV540(); }
function applyRecommendedLaunchSettings() {
  ensureSchemaIfNeeded();
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    __HAS_LOCK = true; __TBL = {}; __SETTINGS = null;
    freezeStartedMatches();
    var next = saveLeagueSettings({ featRules: true, featLateDrops: false, featAttendance: false, featScores: false, featMotm: false, msgConfirm: false, msgTwoHour: false });
    return { saved: true, settings: next, message: "Recommended launch switches applied. Other settings, existing match history and credentials were preserved. No messages were sent." };
  } finally { __HAS_LOCK = false; lock.releaseLock(); }
}

function notificationEligibleProfile(profiles, email, gameDate) {
  var p = profiles[email];
  return !!(p && String(p.status || "Active").toLowerCase() !== "suspended" && (!p.awayUntil || p.awayUntil < gameDate));
}
function scheduledNotificationKey(kind, g) {
  return kind + "|" + g.GameID + "|" + g.Date + "|" + g.Time + "|" + g.LocationName;
}
function queueScheduledReminders(now) {
  now = now || new Date();
  var today = Utilities.formatDate(now, notificationTz(), "yyyy-MM-dd"), minutes = notificationNowMinutes(now), settings = getLeagueSettings();
  var data = getAppData(), tomorrow = isoPlusDays(today, 1);
  var gameStart = function(g) { return typeof gameInstantMs === "function" ? gameInstantMs(g.Date, g.Time) : Number(g.startAtMs || 0); };
  data.games.forEach(function(g) {
    if (g.Status === "Cancelled") return;
    var confirmed = g.confirmedCount === undefined ? Number(g.yesCount || 0) : Number(g.confirmedCount);
    var start = gameStart(g), cutoff = start - LATE_DROP_WINDOW_MINUTES * 60000;
    var cutoffLabel = start ? Utilities.formatDate(new Date(cutoff), notificationTz(), "h:mm a") : "";
    var eligible = function(em) { return notificationEligibleProfile(data.profiles, em, g.Date); };
    if (g.Date === today && minutes >= 8 * 60 + 30 && minutes < 9 * 60 && now.getTime() < cutoff) {
      if (settings.msgMatchDay) {
        var set = {};
        g.roster.filter(function(r) { return r.status === "Yes" || r.status === "Waitlist"; }).map(function(r) { return r.email; })
          .concat(getGameOrganizersRecipients(g.GameID)).filter(eligible).forEach(function(em) { set[em] = true; });
        var enough = confirmed >= g.MinThreshold;
        notifyMail(Object.keys(set), mailMatchday({ game: g, settings: settings, weather: null }), {
          en: [enough ? "⚽ Enough players for today · " + g.Time : "⚽ Players needed today · " + g.Time,
            g.Title + " · " + confirmed + " confirmed · " + g.LocationName + ". Please update your answer before " + cutoffLabel + " Pacific."],
          es: [enough ? "⚽ Hay suficientes jugadores · " + g.Time : "⚽ Faltan jugadores · " + g.Time,
            g.Title + " · " + confirmed + " confirmados · " + g.LocationName + ". Actualiza tu respuesta antes de las " + cutoffLabel + " (Pacífico)."]
        }, { dedupeKey: scheduledNotificationKey("morning", g), gameId: g.GameID, expiresAtMs: cutoff, priority: 30, pushOnly: true, skipAway: true, expectedGameStatus: g.Status, timeSensitive: true });
      }
      if (settings.msgMaybeMorning) {
        var spots = Math.max(0, Number(g.MaxCap) - Number(g.yesCount));
        g.roster.filter(function(r) { return r.status === "Maybe" && eligible(r.email); }).forEach(function(r) {
          notifyMail([r.email], mailMaybeMorning({ playerEmail: r.email, playerName: r.fullName, game: g, spotsLeft: spots }), {
            en: ["🤔 Playing today?", g.Title + " · " + g.Time + ". Please decide before " + cutoffLabel + " Pacific; open the link to confirm your answer."],
            es: ["🤔 ¿Juegas hoy?", g.Title + " · " + g.Time + ". Decide antes de las " + cutoffLabel + " (Pacífico); abre el enlace para confirmar tu respuesta."]
          }, { dedupeKey: scheduledNotificationKey("maybe", g), gameId: g.GameID, expiresAtMs: cutoff, priority: 30, pushOnly: true, skipAway: true, expectedSignupStatus: "Maybe" });
        });
      }
    }
    if (g.Date === tomorrow && minutes >= 10 * 60 && minutes < 18 * 60 && settings.msgDayBefore && confirmed < g.MinThreshold) {
      var responded = {}; g.roster.forEach(function(r) { responded[r.email] = true; });
      Object.keys(data.profiles).filter(function(em) { return !isNonPlayingSystemAdmin(em) && eligible(em) && !responded[em]; }).forEach(function(em) {
        var p = data.profiles[em], remaining = Math.max(0, Number(g.MaxCap) - Number(g.yesCount));
        notifyMail([em], mailTomorrow({ playerEmail: em, playerName: p.fullName, game: g, weather: null, spotsLeft: remaining, showWeather: false }), {
          en: ["⚽ Playing tomorrow?", g.Title + " · " + g.Time + " · needs " + Math.max(0, g.MinThreshold - confirmed) + " more. Open to answer."],
          es: ["⚽ ¿Juegas mañana?", g.Title + " · " + g.Time + " · faltan " + Math.max(0, g.MinThreshold - confirmed) + ". Abre para responder."]
        }, { dedupeKey: scheduledNotificationKey("day-before", g), gameId: g.GameID, expiresAtMs: now.getTime() + 6 * 3600000, priority: 20, pushOnly: true, skipAway: true, expectedGameStatus: g.Status, timeSensitive: true });
      });
    }
    var until = (start - now.getTime()) / 60000;
    if (g.Date === today && settings.msgTwoHour && g.Status === "Confirmed" && until >= 105 && until <= 135) {
      var recipients = g.roster.filter(function(r) { return r.status === "Yes" && eligible(r.email); }).map(function(r) { return r.email; });
      notifyMail(recipients, mailTwoHour({ game: g }), {
        en: ["⚽ Kickoff in about two hours", g.Title + " · " + g.Time + " · " + g.LocationName],
        es: ["⚽ Empezamos en unas dos horas", g.Title + " · " + g.Time + " · " + g.LocationName]
      }, { dedupeKey: scheduledNotificationKey("two-hour", g), gameId: g.GameID, expiresAtMs: start - 90 * 60000, priority: 20, pushOnly: true, skipAway: true, expectedSignupStatus: "Yes", expectedGameStatus: "Confirmed", timeSensitive: true });
    }
  });
}

function leagueAutomationDispatcher() {
  ensureSchemaIfNeeded();
  extendSeriesIfNewDay();
  var lock = LockService.getScriptLock();
  if (lock.tryLock(5000)) {
    try { __HAS_LOCK = true; __TBL = {}; __NOTIFY_WARNINGS = []; freezeStartedMatches(); queueScheduledReminders(new Date()); }
    finally { __HAS_LOCK = false; lock.releaseLock(); }
  }
  processInviteQueue();
  var result = processNotificationOutbox();
  var props = PropertiesService.getScriptProperties(), day = Utilities.formatDate(new Date(), notificationTz(), "yyyy-MM-dd");
  if (props.getProperty("NOTIFY_PRUNED_ON") !== day) {
    notificationLock(function() { __TBL = {}; housekeeping(); });
    pruneNotificationOutbox();
    props.setProperty("NOTIFY_PRUNED_ON", day);
  }
  return result;
}
// Kept for old editor bookmarks; setup removes their old triggers. These names now share
// the same timing windows and idempotency keys, so manually running them does not add duplicates.
function send9amGameDayUpdate() { return leagueAutomationDispatcher(); }
function sendDayBeforeNonRsvpReminder() { return leagueAutomationDispatcher(); }
function sendTwoHourReminders() { return leagueAutomationDispatcher(); }

// Adds any missing sheets/columns (never deletes or reorders). Runs automatically
// once after you deploy a new version, so you don't have to remember initialSetup().
function ensureSchema(force) {
  var props = PropertiesService.getScriptProperties();
  if (!force && props.getProperty("SCHEMA_VERSION") === SCHEMA_VERSION) return false;
  getAppSecret();
  for (var name in SCHEMA) {
    var sheet = SS.getSheetByName(name);
    if (!sheet) sheet = SS.insertSheet(name);
    var required = SCHEMA[name];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(required);
      sheet.getRange(1, 1, 1, required.length).setFontWeight("bold").setBackground("#00f59b").setFontColor("#070b09");
      sheet.setFrozenRows(1);
    } else {
      var lastCol = Math.max(sheet.getLastColumn(), 1);
      var existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h || "").trim(); });
      for (var c = 0; c < required.length; c++) {
        if (existing.indexOf(required[c]) === -1) {
          var newCol = sheet.getLastColumn() + 1;
          sheet.getRange(1, newCol).setValue(required[c]).setFontWeight("bold").setBackground("#00f59b").setFontColor("#070b09");
          existing.push(required[c]);
        }
      }
    }
    delete __TBL[name];
  }
  props.setProperty("SCHEMA_VERSION", SCHEMA_VERSION);
  return true;
}

// One-time cleanup of data left behind by older versions:
//  • guests/gear/notes still attached to RSVPs that are no longer YES ("ghost guests")
//  • duplicate RSVP rows for the same player and game (keeps the newest)
function repairData() {
  var t = T("Signups");
  var fixedGhosts = 0, removedDupes = 0;
  for (var r = 0; r < t.rows.length; r++) {
    var st = String(tv(t, r, "Status"));
    if (st && st !== "Yes") {
      var dirty = {};
      if (String(tv(t, r, "PlusOne")) !== "") dirty.PlusOne = "";
      if (isTruthy(tv(t, r, "BringBall"))) dirty.BringBall = false;
      if (isTruthy(tv(t, r, "BringPinnies"))) dirty.BringPinnies = false;
      if (isTruthy(tv(t, r, "PlayGK"))) dirty.PlayGK = false;
      if (String(tv(t, r, "AvailabilityNote")) !== "") dirty.AvailabilityNote = "";
      if (Object.keys(dirty).length) { twMany(t, r, dirty); fixedGhosts++; }
    }
  }
  var seen = {};
  for (var d = t.rows.length - 1; d >= 0; d--) {
    var key = String(tv(t, d, "GameID")) + "|" + String(tv(t, d, "Email")).toLowerCase().trim();
    if (key === "|") continue;
    if (seen[key]) { tdel(t, d); removedDupes++; }
    else seen[key] = true;
  }
  return { ghostGuestRowsCleared: fixedGhosts, duplicateRsvpsRemoved: removedDupes };
}

// ============================================================================
// CRYPTOGRAPHIC PIN HASHING, RATE LIMITING & SESSION TOKENS
// ============================================================================
// The secret that signs invite links, magic RSVP links and peppers PIN hashes. It is created
// on first use, so there is never a window where a fixed, guessable fallback key is used.
function getAppSecret() {
  var props = PropertiesService.getScriptProperties();
  var secret = props.getProperty("MAGIC_SECRET");
  if (!secret) { secret = "PPL_" + Utilities.getUuid() + Utilities.getUuid().slice(0, 8); props.setProperty("MAGIC_SECRET", secret); }
  return secret;
}

function computeSaltedPinHash(email, pin, salt) {
  var secret = getAppSecret();
  var payload = String(salt) + "|" + String(email).toLowerCase().trim() + "|" + String(pin).trim() + "|" + secret;
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, payload);
  return Utilities.base64EncodeWebSafe(digest);
}

function createPinHashRecord(email, pin) {
  var salt = Utilities.getUuid().replace(/-/g, "").slice(0, 16);
  var hash = computeSaltedPinHash(email, pin, salt);
  return "HASH$" + salt + "$" + hash;
}

function verifyStoredPinAndMigrateIfLegacy(profileRow, email, storedCellVal, candidatePin) {
  var cleanCandidate = String(candidatePin || "").trim();
  if (!isValidSixDigitPin(cleanCandidate)) return false;

  var rawStored = String(storedCellVal || "").replace(/^'/, "").trim();
  if (rawStored.indexOf("HASH$") === 0) {
    var parts = rawStored.split("$");
    var salt = parts[1] || "";
    var expectedHash = parts[2] || "";
    var actualHash = computeSaltedPinHash(email, cleanCandidate, salt);
    return actualHash === expectedHash && expectedHash.length > 0;
  }
  var paddedStored = rawStored.length < 6 && /^\d+$/.test(rawStored) ? ("000000" + rawStored).slice(-6) : rawStored;
  if (paddedStored === cleanCandidate) {
    tw(T("Profiles"), profileRow, "PinCode", createPinHashRecord(email, cleanCandidate));
    return true;
  }
  return false;
}

// Strictly verifies an Admin's actual 6-digit numeric PIN (rejects session tokens) for unlocking sensitive settings
function verifyActualSixDigitPinForAdmin(email, actualPin) {
  var em = String(email || "").toLowerCase().trim();
  var pin = String(actualPin || "").trim();
  if (!em || !isValidSixDigitPin(pin)) return false;
  var throttle = checkAuthThrottle(em);
  if (!throttle.allowed) return false;
  var t = T("Profiles");
  var r = findProfileRow(em);
  if (r === -1) return false;
  var role = isProtectedAdmin(em) ? "Admin" : String(tv(t, r, "Role") || "Player").trim();
  var status = isProtectedAdmin(em) ? "Active" : String(tv(t, r, "AccountStatus") || "Active").trim();
  if (role.toLowerCase() !== "admin" || status.toLowerCase() === "suspended") return false;
  var ok = verifyStoredPinAndMigrateIfLegacy(r, em, tv(t, r, "PinCode"), pin);
  if (!ok) { recordAuthFailure(em); return false; }
  clearAuthThrottle(em);
  return true;
}

function checkAuthThrottle(email) {
  var key = "THROTTLE_" + String(email || "").toLowerCase().trim();
  var raw = PropertiesService.getScriptProperties().getProperty(key);
  if (!raw) return { allowed: true };
  try {
    var obj = JSON.parse(raw);
    var now = new Date().getTime();
    if (obj.lockUntil && now < obj.lockUntil) {
      var minsLeft = Math.max(1, Math.ceil((obj.lockUntil - now) / 60000));
      return { allowed: false, error: tr("errLocked", { mins: minsLeft }) };
    }
    if (obj.lockUntil && now >= obj.lockUntil) PropertiesService.getScriptProperties().deleteProperty(key);
  } catch (e) {}
  return { allowed: true };
}

function recordAuthFailure(email) {
  var key = "THROTTLE_" + String(email || "").toLowerCase().trim();
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(key);
  var obj = { count: 0, lockUntil: 0 };
  if (raw) { try { obj = JSON.parse(raw); } catch (e) {} }
  obj.count = (obj.count || 0) + 1;
  if (obj.count >= 5) obj.lockUntil = new Date().getTime() + (15 * 60 * 1000);
  props.setProperty(key, JSON.stringify(obj));
}

function clearAuthThrottle(email) {
  PropertiesService.getScriptProperties().deleteProperty("THROTTLE_" + String(email || "").toLowerCase().trim());
}

// Sessions: up to MAX_SESSIONS_PER_USER tokens per player (one per device), sliding expiry.
function readSessions(em) {
  var raw = PropertiesService.getScriptProperties().getProperty("SESSION_" + em);
  if (!raw) return [];
  try {
    var obj = JSON.parse(raw);
    if (obj && obj.token) return [{ t: obj.token, exp: Number(obj.expiresAt || 0) }]; // v4.x single-session format
    if (obj && Array.isArray(obj.tokens)) return obj.tokens;
  } catch (e) {}
  return [];
}

function writeSessions(em, list) {
  var props = PropertiesService.getScriptProperties();
  var now = new Date().getTime();
  var live = list.filter(function(s) { return s && s.t && s.exp > now; }).slice(0, MAX_SESSIONS_PER_USER);
  if (live.length === 0) props.deleteProperty("SESSION_" + em);
  else props.setProperty("SESSION_" + em, JSON.stringify({ tokens: live }));
}

function issueSessionToken(email) {
  var em = String(email || "").toLowerCase().trim();
  var token = "SES_" + Utilities.getUuid().replace(/-/g, "") + Utilities.getUuid().replace(/-/g, "").slice(0, 12);
  var list = readSessions(em);
  list.unshift({ t: token, exp: new Date().getTime() + SESSION_DAYS * 86400000 });
  writeSessions(em, list);
  return token;
}

function verifySessionToken(email, token) {
  var em = String(email || "").toLowerCase().trim();
  var tk = String(token || "").trim();
  if (!em || !tk || tk.indexOf("SES_") !== 0) return false;
  var list = readSessions(em);
  var now = new Date().getTime();
  for (var i = 0; i < list.length; i++) {
    if (list[i].t === tk) {
      var expired = list[i].exp <= now;
      var renew = !expired && list[i].exp - now < 10 * 86400000; // renew quietly while the player stays active
      if (expired || renew) {
        updateSessionsSafely(em, function(cur) {
          if (expired) return cur.filter(function(s) { return s.t !== tk; });
          return cur.map(function(s) { return s.t === tk ? { t: s.t, exp: now + SESSION_DAYS * 86400000 } : s; });
        });
      }
      return !expired;
    }
  }
  return false;
}

// Read-modify-write of a player's session list. The 60-second refresh runs without the global
// lock, so it only touches sessions when it can grab the lock quickly; otherwise a login on
// another device at the same moment could be overwritten. Skipping a renewal is harmless.
function updateSessionsSafely(em, fn) {
  if (__HAS_LOCK) { writeSessions(em, fn(readSessions(em))); return; }
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1500)) return;
  try { writeSessions(em, fn(readSessions(em))); } finally { lock.releaseLock(); }
}

function revokeSessionToken(email, token) {
  var em = String(email || "").toLowerCase().trim();
  writeSessions(em, readSessions(em).filter(function(s) { return s.t !== token; }));
}

function revokeAllSessions(email) {
  PropertiesService.getScriptProperties().deleteProperty("SESSION_" + String(email || "").toLowerCase().trim());
}

// ============================================================================
// MULTI-CHANNEL NOTIFICATION ROUTER (Push Only, Email Only, or Both)
// ============================================================================
function getPlayerNotifyPrefMap() {
  if (__PREF_MAP) return __PREF_MAP;
  var map = {};
  var t = T("Profiles");
  for (var r = 0; r < t.rows.length; r++) {
    var em = String(tv(t, r, "Email") || "").toLowerCase().trim();
    if (em) map[em] = normalizeNotifyPref(tv(t, r, "NotifyPref"));
  }
  __PREF_MAP = map;
  return map;
}

// Each player's app language ("en" / "es"), so emails and pushes reach them in it.
function getPlayerLangMap() {
  if (__LANG_MAP) return __LANG_MAP;
  var map = {};
  var t = T("Profiles");
  var hasLang = tHas(t, "Lang");
  for (var r = 0; r < t.rows.length; r++) {
    var em = String(tv(t, r, "Email") || "").toLowerCase().trim();
    if (em) map[em] = hasLang ? normalizeLang(tv(t, r, "Lang")) : "en";
  }
  __LANG_MAP = map;
  return map;
}

function sendOneSignalPushNotification(targetEmails, title, message, launchUrl, titleEs, messageEs) {
  var props = PropertiesService.getScriptProperties();
  var appId = String(props.getProperty("ONESIGNAL_APP_ID") || "").trim();
  var restKey = String(props.getProperty("ONESIGNAL_REST_KEY") || "").trim();
  if (!appId || !restKey || !targetEmails || targetEmails.length === 0) return { sent: false, reason: "Push not configured" };

  var cleanList = [];
  var seen = {};
  targetEmails.forEach(function(em) {
    var clean = String(em || "").toLowerCase().trim();
    if (clean && !seen[clean]) { seen[clean] = true; cleanList.push(clean); }
  });
  if (cleanList.length === 0) return { sent: false, reason: "No recipients" };

  try {
    var payload = {
      app_id: appId,
      include_aliases: { external_id: cleanList },
      target_channel: "push",
      headings: { en: String(title || "⚽ Pitch Pulse League"), es: String(titleEs || title || "⚽ Pitch Pulse League") },
      contents: { en: String(message || "Match update available on the hub."), es: String(messageEs || message || "Hay novedades en la liga.") },
      url: launchUrl || SITE_URL
    };
    var resp = UrlFetchApp.fetch("https://api.onesignal.com/notifications", {
      method: "post",
      contentType: "application/json; charset=utf-8",
      headers: { "Authorization": "Key " + restKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    if (code < 200 || code >= 300) return { sent: false, count: 0, failed: cleanList, reason: "HTTP " + code + ": " + resp.getContentText() };
    // OneSignal answers 200 even when some players have no subscribed device; it lists them.
    var body = {};
    try { body = JSON.parse(resp.getContentText() || "{}") || {}; } catch (pe) { body = {}; }
    var failed = [];
    var errs = body.errors;
    if (errs && errs.invalid_aliases && Array.isArray(errs.invalid_aliases.external_id)) {
      var bad = {};
      errs.invalid_aliases.external_id.forEach(function(x) { bad[String(x).toLowerCase().trim()] = true; });
      failed = cleanList.filter(function(em) { return bad[em]; });
    } else if (!body.id && errs) {
      failed = cleanList.slice(); // e.g. "All included players are not subscribed"
    }
    return { sent: failed.length < cleanList.length, count: cleanList.length - failed.length, failed: failed };
  } catch (e) {
    return { sent: false, count: 0, failed: cleanList, reason: e.toString() };
  }
}

var __QUOTA = null;
function getEmailQuotaLeft(fresh) {
  if (__QUOTA !== null && !fresh) return __QUOTA;
  try { __QUOTA = MailApp.getRemainingDailyQuota(); } catch (e) { __QUOTA = -1; }
  return __QUOTA;
}

// During a web request notifications are queued and sent AFTER the database lock is
// released, so one slow email never blocks other players' RSVPs.
// Durable notification queue. Rows remain private to the Apps Script owner.
// "accepted" means the provider accepted a request, never that a player read it.
var NOTIFY_MAX_PER_RUN = 40;
var NOTIFY_RUN_BUDGET_MS = 150000;
var NOTIFY_RETENTION_DAYS = 7;
var NOTIFY_MAX_PENDING = 2000;
var __NOTIFY_WARNINGS = [];

function notificationLock(fn) {
  if (__HAS_LOCK) return fn();
  var lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try { delete __TBL.Notifications; return fn(); }
  finally { lock.releaseLock(); }
}

function notificationMillis(v) { return v ? Number(new Date(v).getTime()) || Number(v) || 0 : 0; }
function notificationTz() { return typeof leagueTimeZone === "function" ? leagueTimeZone() : "America/Los_Angeles"; }
function notificationNowMinutes(d) {
  return Number(Utilities.formatDate(d || new Date(), notificationTz(), "H")) * 60 + Number(Utilities.formatDate(d || new Date(), notificationTz(), "m"));
}
function notificationPlain(html) {
  return String(html || "").replace(/<a[^>]*href=['"]([^'"]+)['"][^>]*>([\s\S]*?)<\/a>/gi, "$2 ($1)")
    .replace(/<\/(p|div|li|h[1-6])>|<br\s*\/?\s*>/gi, "\n").replace(/<[^>]*>/g, "")
    .replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&#39;/g, "'")
    .replace(/\n[ \t]+/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
}

// Mail factories attach a small snapshot so queued messages are not sent after a later
// cancellation, restoration, move, or removed match makes them misleading.
function notificationMail(fn, game, opts) {
  var g = game || {};
  fn.notificationOptions = Object.assign({}, opts || {}, { gameId: String(g.id || g.GameID || "") });
  return fn;
}
function notificationCountSummary(game, lang) {
  var es = lang === "es", reserved = Number(game.yesCount || 0), confirmed = game.confirmedCount === undefined ? reserved : Number(game.confirmedCount || 0), pending = Number(game.pendingReconfirmCount || 0);
  return confirmed + (es ? " confirmados" : " confirmed") + (pending ? " · " + reserved + (es ? " lugares reservados · " : " places held · ") + pending + (es ? " por reconfirmar" : " awaiting re-confirmation") : "");
}
function notificationGameSnapshot(gameId) {
  if (!gameId) return null;
  var r = findGameRow(gameId);
  if (r === -1) return { id: String(gameId), missing: true };
  var g = gameRecord(r);
  return { id: String(gameId), date: g.date, time: g.time, location: g.location, status: g.status };
}
function notificationContextCurrent(payload, recipient) {
  var c = payload.context;
  if (c) {
    var cur = notificationGameSnapshot(c.id);
    if (!cur || cur.missing || cur.date !== c.date || cur.time !== c.time || cur.location !== c.location || (cur.status === "Cancelled") !== (c.status === "Cancelled")) return false;
    if (payload.expectedGameStatus && cur.status !== payload.expectedGameStatus) return false;
    if (payload.requireNoAnswer && findSignupRow(c.id, recipient) !== -1) return false;
    if (payload.expectedSignupStatus) {
      var sr = findSignupRow(c.id, recipient);
      if (sr === -1 || String(tv(T("Signups"), sr, "Status")) !== payload.expectedSignupStatus) return false;
    }
  }
  if (!payload.invite) {
    var pr = findProfileRow(recipient);
    if (pr === -1) return false;
    if (!isProtectedAdmin(recipient) && String(tv(T("Profiles"), pr, "AccountStatus")).toLowerCase() === "suspended") return false;
    if (payload.skipAway && c) {
      var away = formatDateIso(tv(T("Profiles"), pr, "AwayUntil"));
      if (away && away >= c.date) return false;
    }
  } else {
    if (findProfileRow(recipient) !== -1) return false;
    var it = T("Invites");
    if (!isProtectedAdmin(recipient) && tfind(it, function(r) { return String(tv(it, r, "Email")).toLowerCase().trim() === recipient; }) === -1) return false;
  }
  return true;
}

function queueNotificationJob(job) {
  var opts = job.opts || {}, now = new Date().getTime();
  var recipients = Array.isArray(job.recipients) ? job.recipients : [job.recipients];
  var seen = {}, list = [];
  recipients.forEach(function(em) { var v = String(em || "").toLowerCase().trim(); if (v && !seen[v]) { seen[v] = true; list.push(v); } });
  if (!list.length) return { queued: true, recipientCount: 0, emailCount: 0, pushCount: 0 };
  var prefMap = getPlayerNotifyPrefMap(), langs = getPlayerLangMap();
  var context = notificationGameSnapshot(opts.gameId || "");
  var expiry = Number(opts.expiresAtMs) || now + (opts.invite ? 7 : 2) * 86400000;
  if (opts.timeSensitive && context && !context.missing && typeof gameInstantMs === "function") expiry = Math.min(expiry, gameInstantMs(context.date, context.time) + 75 * 60000);
  var priority = opts.critical ? 100 : (opts.priority !== undefined ? Number(opts.priority) : 50);
  var eventId = String(opts.dedupeKey || Utilities.getUuid());
  return notificationLock(function() {
    var t = T("Notifications"), existing = {}, pending = 0;
    for (var r = 0; r < t.rows.length; r++) {
      existing[String(tv(t, r, "DedupKey"))] = true;
      if (["pending", "sending", "fallback"].indexOf(String(tv(t, r, "EmailState"))) !== -1 || ["pending", "sending"].indexOf(String(tv(t, r, "PushState"))) !== -1) pending++;
    }
    var newCount = list.filter(function(em) { return !existing[eventId + "|" + em]; }).length;
    if (pending + newCount > NOTIFY_MAX_PENDING) throw new Error("Notification queue is full. The saved match change remains valid; re-send its notice after checking Admin Settings.");
    var rows = [];
    list.forEach(function(em) {
      var dedup = eventId + "|" + em;
      if (existing[dedup]) return;
      var lang = langs[em] === "es" ? "es" : "en", pref = prefMap[em] || "Both";
      var content = opts.mail ? opts.mail(lang) : null;
      var title = lang === "es" ? (opts.pushTitleEs || job.pushTitle || job.subject) : (job.pushTitle || job.subject);
      var body = lang === "es" ? (opts.pushBodyEs || job.pushBody || job.subject) : (job.pushBody || job.subject);
      if (!content || !content.html) content = { subject: title, html: "<p>" + escapeHtmlServer(body) + "</p><p><a href='" + escapeHtmlServer(job.launchUrl || SITE_URL) + "'>Open Pitch Pulse / Abrir Pitch Pulse</a></p>" };
      var payload = { subject: String(content.subject || title || MAIL_SENDER_NAME), html: String(content.html), plain: String(content.text || notificationPlain(content.html)),
        title: String(title || content.subject), body: String(body || content.subject), url: String(job.launchUrl || SITE_URL), replyTo: opts.replyTo || "",
        context: context, invite: !!opts.invite, skipAway: !!opts.skipAway, requireNoAnswer: !!opts.requireNoAnswer, expectedGameStatus: opts.expectedGameStatus || "", expectedSignupStatus: opts.expectedSignupStatus || "", supersedeKey: opts.supersedeKey || "", fallbackOnly: pref === "Push" || !!opts.pushOnly };
      var serialized = JSON.stringify(payload);
      if (serialized.length > 45000) throw new Error("A notification is too long to queue. Shorten the announcement and try again.");
      var emailState = "pending", pushState = "pending";
      if (opts.emailOnly || opts.invite || pref === "Email") pushState = "skipped";
      else if (pref === "Push" || opts.pushOnly) emailState = "fallback";
      rows.push({ JobID: Utilities.getUuid(), DedupKey: dedup, CreatedAt: new Date(now), ExpiresAt: new Date(expiry), Priority: priority, Recipient: em,
        PayloadJSON: serialized, EmailState: emailState, PushState: pushState, EmailAttempts: 0, PushAttempts: 0, NextAttemptAt: new Date(now),
        LeaseToken: "", LeaseUntil: "", UpdatedAt: new Date(now), LastError: "", ProviderID: "" });
    });
    // Expire only unfinished channels of older state-change messages, never re-send an accepted one.
    if (opts.supersedeKey && rows.length) {
      var replacedRecipients = {};
      rows.forEach(function(row) { replacedRecipients[row.Recipient] = true; });
      for (var j = 0; j < t.rows.length; j++) {
        var old = {}; try { old = JSON.parse(String(tv(t, j, "PayloadJSON") || "{}")); } catch (e) {}
        if (old.supersedeKey !== opts.supersedeKey || !replacedRecipients[String(tv(t, j, "Recipient")).toLowerCase().trim()]) continue;
        var fields = {};
        ["EmailState", "PushState"].forEach(function(k) { if (["pending", "fallback"].indexOf(String(tv(t, j, k))) !== -1) fields[k] = "expired"; });
        if (Object.keys(fields).length) { fields.LastError = "Superseded by a newer match update."; fields.UpdatedAt = new Date(now); twMany(t, j, fields); }
      }
    }
    taddMany(t, rows);
    return { queued: true, recipientCount: list.length, queuedCount: rows.length, emailCount: 0, pushCount: 0 };
  });
}

function dispatchLeagueNotification(recipientEmails, subject, htmlBody, pushTitle, pushBody, launchUrl, opts) {
  try { return queueNotificationJob({ recipients: recipientEmails, subject: subject, htmlBody: htmlBody, pushTitle: pushTitle, pushBody: pushBody, launchUrl: launchUrl, opts: opts || {} }); }
  catch (e) {
    var msg = String(e && e.message || e).slice(0, 240);
    __NOTIFY_WARNINGS.push(msg);
    try { PropertiesService.getScriptProperties().setProperty("NOTIFY_QUEUE_ERROR", JSON.stringify({ at: new Date().toISOString(), message: msg })); } catch (ignored) {}
    Logger.log("Notification enqueue failed: " + msg);
    return { queued: false, recipientCount: 0, emailCount: 0, pushCount: 0, mailError: msg };
  }
}

// Compatibility names remain, but web requests only queue. The five-minute dispatcher sends.
function flushNotificationQueue() { __NOTIFY_QUEUE = []; __DEFER_NOTIFY = false; }
function deliverNotification(job) { return queueNotificationJob(job); }

function notificationRowById(id) { var t = T("Notifications"); return tfind(t, function(r) { return String(tv(t, r, "JobID")) === id; }); }
function updateNotificationRow(id, lease, fields) {
  return notificationLock(function() {
    var t = T("Notifications"), r = notificationRowById(id);
    if (r === -1 || (lease && String(tv(t, r, "LeaseToken")) !== lease)) return false;
    fields.UpdatedAt = new Date(); twMany(t, r, fields); return true;
  });
}
function claimNotificationJob() {
  return notificationLock(function() {
    __TBL = {};
    var t = T("Notifications"), now = new Date().getTime(), candidates = [];
    for (var r = 0; r < t.rows.length; r++) {
      var e = String(tv(t, r, "EmailState")), p = String(tv(t, r, "PushState"));
      if (notificationMillis(tv(t, r, "LeaseUntil")) > now) continue;
      // MailApp has no idempotency key. An interrupted email send has an unknown outcome;
      // hold it for human review rather than automatically risk a duplicate.
      var recovered = {};
      if (e === "sending") { recovered.EmailState = e = "review"; recovered.LastError = "Email outcome unknown after interruption. Check the sender mailbox before sending a new notice."; }
      if (p === "sending") recovered.PushState = p = "pending"; // same UUID makes a OneSignal retry safe
      if (Object.keys(recovered).length) twMany(t, r, recovered);
      if (["pending", "fallback"].indexOf(e) === -1 && p !== "pending") continue;
      var payload;
      try { payload = JSON.parse(String(tv(t, r, "PayloadJSON") || "{}")); } catch (bad) { twMany(t, r, { EmailState: "review", PushState: "review", LastError: "Invalid queued message." }); continue; }
      if (notificationMillis(tv(t, r, "ExpiresAt")) <= now || !notificationContextCurrent(payload, String(tv(t, r, "Recipient")))) {
        var expired = { LastError: "Message expired or its match details changed.", UpdatedAt: new Date(now) };
        if (["pending", "fallback"].indexOf(e) !== -1) expired.EmailState = "expired";
        if (p === "pending") expired.PushState = "expired";
        twMany(t, r, expired); continue;
      }
      if (notificationMillis(tv(t, r, "NextAttemptAt")) <= now) candidates.push({ r: r, priority: Number(tv(t, r, "Priority")), created: notificationMillis(tv(t, r, "CreatedAt")) });
    }
    candidates.sort(function(a, b) { return b.priority - a.priority || a.created - b.created; });
    if (!candidates.length) return null;
    var row = candidates[0].r, lease = Utilities.getUuid();
    twMany(t, row, { LeaseToken: lease, LeaseUntil: new Date(now + 7 * 60000) });
    var out = tobj(t, row); out.payload = JSON.parse(String(out.PayloadJSON)); return out;
  });
}

function sendQueuedPush(job) {
  var props = PropertiesService.getScriptProperties(), appId = props.getProperty("ONESIGNAL_APP_ID"), key = props.getProperty("ONESIGNAL_REST_KEY");
  if (!appId || !key) return { state: "skipped", fallback: true, error: "Push is not configured; using email." };
  var payload = job.payload;
  var ttl = Math.max(0, Math.min(259200, Math.floor((notificationMillis(job.ExpiresAt) - new Date().getTime()) / 1000)));
  try {
    var response = UrlFetchApp.fetch("https://api.onesignal.com/notifications?c=push", { method: "post", contentType: "application/json; charset=utf-8", muteHttpExceptions: true,
      headers: { Authorization: "Key " + key }, payload: JSON.stringify({ app_id: appId, include_aliases: { external_id: [job.Recipient] }, target_channel: "push",
        headings: { en: payload.title }, contents: { en: payload.body }, url: payload.url, ttl: ttl, idempotency_key: job.JobID }) });
    var code = response.getResponseCode(), body = {};
    try { body = JSON.parse(response.getContentText() || "{}"); } catch (parseError) {}
    if (code >= 200 && code < 300) {
      var bad = body.errors && body.errors.invalid_aliases && body.errors.invalid_aliases.external_id;
      if (!body.id || (Array.isArray(bad) && bad.indexOf(job.Recipient) !== -1)) return { state: "skipped", fallback: true, error: "No subscribed push device; using email." };
      return { state: "accepted", providerId: String(body.id), error: "" };
    }
    if (code === 429 || code >= 500) {
      var headers = response.getAllHeaders ? response.getAllHeaders() : {}, retry = Number(headers["Retry-After"] || headers["retry-after"] || 0);
      return { state: "pending", fallback: true, retryMs: retry > 0 ? retry * 1000 : 300000, error: "Push provider temporarily unavailable (HTTP " + code + ")." };
    }
    return { state: "review", fallback: true, error: "Push configuration/request rejected (HTTP " + code + "). Check OneSignal settings." };
  } catch (e) { return { state: "pending", fallback: true, retryMs: 300000, error: "Push request outcome unknown; retry will use the same idempotency key." }; }
}

function processClaimedNotification(job) {
  var id = String(job.JobID), lease = String(job.LeaseToken), payload = job.payload;
  var email = String(job.EmailState), push = String(job.PushState), err = "", retryMs = 300000;
  if (push === "pending") {
    var pa = Number(job.PushAttempts || 0) + 1;
    if (!updateNotificationRow(id, lease, { PushState: "sending", PushAttempts: pa })) return;
    var result = sendQueuedPush(job);
    push = result.state; err = result.error || ""; retryMs = Math.max(retryMs, result.retryMs || 0);
    if (pa >= 8 && push === "pending") { push = "review"; err = "Push retry limit reached; inspect OneSignal settings."; }
    if (email === "fallback") email = result.fallback ? "pending" : (push === "accepted" ? "skipped" : "fallback");
    if (!updateNotificationRow(id, lease, { PushState: push, EmailState: email, ProviderID: result.providerId || job.ProviderID || "", LastError: err })) return;
  }
  if (email === "fallback" && ["skipped", "review", "expired"].indexOf(push) !== -1) email = "pending";
  if (email === "pending") {
    // Leave two recipient slots for urgent PIN recovery, even for critical match changes.
    var reserve = Number(job.Priority) >= 100 ? 2 : EMAIL_QUOTA_RESERVE;
    var quota = getEmailQuotaLeft(true);
    if (quota < 0 || quota <= reserve) { err = "Email waiting for sending quota; it will be retried before expiry."; retryMs = Math.max(retryMs, 30 * 60000); }
    else {
      if (!updateNotificationRow(id, lease, { EmailState: "sending", EmailAttempts: Number(job.EmailAttempts || 0) + 1 })) return;
      try {
        var options = { to: job.Recipient, subject: payload.subject, htmlBody: payload.html, body: payload.plain || payload.subject, name: MAIL_SENDER_NAME };
        if (payload.replyTo) options.replyTo = payload.replyTo;
        MailApp.sendEmail(options);
        email = "accepted";
        // A fallback has succeeded. Do not keep chasing an optional push channel.
        if (payload.fallbackOnly && push === "pending") push = "skipped";
      } catch (mailError) {
        email = "review";
        err = "Email send outcome is uncertain. Check the sender mailbox before sending a new notice; automatic retry is paused to avoid duplicates.";
      }
      if (!updateNotificationRow(id, lease, { EmailState: email, PushState: push, LastError: err })) return;
    }
  }
  updateNotificationRow(id, lease, { EmailState: email, PushState: push, NextAttemptAt: new Date(new Date().getTime() + retryMs), LeaseToken: "", LeaseUntil: "", LastError: err });
}

function processNotificationOutbox() {
  var props = PropertiesService.getScriptProperties(), token = Utilities.getUuid(), now = new Date().getTime();
  var owns = notificationLock(function() {
    var old = {}; try { old = JSON.parse(props.getProperty("NOTIFY_RUNNER") || "{}"); } catch (e) {}
    if (Number(old.until || 0) > now) return false;
    props.setProperty("NOTIFY_RUNNER", JSON.stringify({ token: token, until: now + 7 * 60000 })); return true;
  });
  if (!owns) return { busy: true };
  var processed = 0;
  try {
    while (processed < NOTIFY_MAX_PER_RUN && new Date().getTime() - now < NOTIFY_RUN_BUDGET_MS) {
      var job = claimNotificationJob(); if (!job) break;
      processClaimedNotification(job); processed++;
    }
    props.setProperty("NOTIFY_LAST_RUN", new Date().toISOString());
  } finally {
    notificationLock(function() {
      var held = {}; try { held = JSON.parse(props.getProperty("NOTIFY_RUNNER") || "{}"); } catch (e) {}
      if (held.token === token) props.deleteProperty("NOTIFY_RUNNER");
    });
  }
  return { processed: processed };
}

function notificationSummary() {
  var sum = { pending: 0, accepted: 0, needsReview: 0, expired: 0, lastRunAt: "", lastError: "" };
  var t = T("Notifications"), now = new Date().getTime();
  for (var r = 0; r < t.rows.length; r++) {
    if (notificationMillis(tv(t, r, "CreatedAt")) < now - NOTIFY_RETENTION_DAYS * 86400000) continue;
    ["EmailState", "PushState"].forEach(function(k) {
      var state = String(tv(t, r, k));
      if (["pending", "sending", "fallback"].indexOf(state) !== -1) sum.pending++;
      else if (state === "accepted") sum.accepted++;
      else if (state === "review") sum.needsReview++;
      else if (state === "expired") sum.expired++;
    });
    if (String(tv(t, r, "EmailState")) === "review" || String(tv(t, r, "PushState")) === "review") sum.lastError = String(tv(t, r, "LastError") || "");
  }
  var props = PropertiesService.getScriptProperties();
  sum.lastRunAt = props.getProperty("NOTIFY_LAST_RUN") || "";
  try { var q = JSON.parse(props.getProperty("NOTIFY_QUEUE_ERROR") || "null"); if (q && new Date(q.at).getTime() > now - 7 * 86400000) sum.lastError = q.message; } catch (e) {}
  return sum;
}
function adminRetryNotifications(data) {
  var admin = requireAdmin(data); if (!admin) return fail(tr("errAdminOnly"));
  var t = T("Notifications"), n = 0, now = new Date();
  for (var r = 0; r < t.rows.length; r++) {
    if (notificationMillis(tv(t, r, "ExpiresAt")) <= now.getTime() || notificationMillis(tv(t, r, "LeaseUntil")) > now.getTime()) continue;
    var e = String(tv(t, r, "EmailState")), p = String(tv(t, r, "PushState"));
    // Unknown email outcomes deliberately remain review. A push retry keeps its original UUID.
    if (e === "pending" || p === "pending" || p === "review") {
      var fields = { NextAttemptAt: now };
      if (p === "review") { fields.PushState = "pending"; fields.PushAttempts = 0; }
      twMany(t, r, fields); n++;
    }
  }
  return okData(admin, __LANG === "es" ? "Reintentos en cola: " + n + ". Se revisan en la próxima ejecución; los mensajes aceptados no se vuelven a enviar." : "Retry queued for " + n + " message(s), for the next scheduled run. Accepted channels are not sent again; uncertain email outcomes still need review.");
}
function pruneNotificationOutbox() {
  return notificationLock(function() {
    var t = T("Notifications"), cutoff = new Date().getTime() - NOTIFY_RETENTION_DAYS * 86400000, count = 0;
    for (var r = t.rows.length - 1; r >= 0 && count < 100; r--) {
      if (notificationMillis(tv(t, r, "CreatedAt")) < cutoff && notificationMillis(tv(t, r, "LeaseUntil")) <= new Date().getTime()) { tdel(t, r); count++; }
    }
    return count;
  });
}

// Work that must wait until the global lock is released (sending invite emails, for example).
var __AFTER_LOCK_JOBS = [];
function runAfterLockJobs() {
  var jobs = __AFTER_LOCK_JOBS;
  __AFTER_LOCK_JOBS = [];
  jobs.forEach(function(job) { try { job(); } catch (e) { Logger.log("After-lock job failed: " + e); } });
}





// ============================================================================
// API ROUTERS
// ============================================================================
function doGet(e) {
  var params = e && e.parameter ? e.parameter : {};
  __LANG = normalizeLang(params.lang);
  if (params.action === "magicRsvp") {
    try { ensureSchemaIfNeeded(); } catch (x) {}
    return renderMagicEmailConfirmationPage(params);
  }
  return jsonResponse({ success: true, version: APP_VERSION, status: "Pitch Pulse Private League API v" + APP_VERSION + " Active. Authentication required." });
}

// Actions that never write anything important run without the global lock.
var READ_ONLY_ACTIONS = { getLeagueData: true, previewEmailAnswer: true };

function doPost(e) {
  var data = {};
  try {
    data = JSON.parse(e && e.postData ? e.postData.contents : "{}") || {};
  } catch (pe) {
    return jsonResponse({ success: false, error: "Malformed request." });
  }
  __LANG = normalizeLang(data.lang);
  __PREF_MAP = null; __LANG_MAP = null; __NOTIFY_WARNINGS = [];
  var action = String(data.action || "");
  var routes = {
    // Public auth & magic-link confirmation
    register: registerUser,
    login: loginUser,
    logout: logoutUser,
    requestPinReset: requestPinReset,
    resetPin: resetPin,
    confirmMagicRsvp: confirmMagicRsvpWithPin,
    emailAnswer: emailAnswer,
    previewEmailAnswer: previewEmailAnswer,
    // Authenticated player endpoints
    getLeagueData: getLeagueDataAuthenticated,
    updateSignup: updateSignup,
    updateSignupExtras: updateSignupExtras,
    updateProfileDetails: updateProfileDetails,
    voteMotm: voteMotm,
    captainStepDown: captainStepDown,
    // Game-scoped endpoints (assigned match captain OR any admin)
    adminGenerateTeams: adminGenerateTeams,
    adminSwapTeamPlayers: adminSwapTeamPlayers,
    adminMoveTeamPlayer: adminMoveTeamPlayer,
    adminSaveScore: adminSaveScore,
    adminUpdateAttendance: adminUpdateAttendance,
    adminSendGameBlast: adminSendGameBlast,
    adminApproveSmallSided: adminApproveSmallSided,
    adminNudgeUnresponded: adminNudgeUnresponded,
    // Global admin endpoints
    adminSaveGame: adminSaveGame,
    adminCancelGame: adminCancelGame,
    adminRestoreGame: adminRestoreGame,
    adminDeleteGame: adminDeleteGame,
    adminSaveRecurringSeries: adminSaveRecurringSeries,
    adminStopRecurringSeries: adminStopRecurringSeries,
    adminSaveLocation: adminSaveLocation,
    adminToggleLocationStatus: adminToggleLocationStatus,
    adminDeleteLocation: adminDeleteLocation,
    adminSendInvite: adminSendInvite,
    adminSendInvites: adminSendInvites,
    adminDeleteInvite: adminDeleteInvite,
    adminUpdateUserRole: adminUpdateUserRole,
    adminToggleSuspendUser: adminToggleSuspendUser,
    adminToggleGuestPrivilege: adminToggleGuestPrivilege,
    adminSetDefaultTeam: adminSetDefaultTeam,
    adminTogglePinniesSetting: adminTogglePinniesSetting,
    adminSaveLeagueSettings: adminSaveLeagueSettings,
    adminSaveHowWePlay: adminSaveHowWePlay,
    adminDeleteUser: adminDeleteUser,
    adminToggleBlastVisibility: adminToggleBlastVisibility,
    adminMessagePlayer: adminMessagePlayer,
    adminRetryNotifications: adminRetryNotifications,
    adminRenameTeams: adminRenameTeams,
    // PIN-locked push configuration
    adminUnlockPushConfig: adminUnlockPushConfig,
    adminSavePushConfig: adminSavePushConfig
  };
  var handler = routes[action];
  if (!handler) return jsonResponse({ success: false, error: "Unknown action: " + action });

  var lock = null;
  try {
    ensureSchemaIfNeeded();
    if (action !== "previewEmailAnswer") extendSeriesIfNewDay();
    if (!READ_ONLY_ACTIONS[action]) {
      lock = LockService.getScriptLock();
      try {
        lock.waitLock(25000);
      } catch (le) {
        lock = null;
        return jsonResponse({ success: false, busy: true, error: tr("errBusy") });
      }
      __HAS_LOCK = true;
      __TBL = {}; // re-read everything now that we own the lock
    }
    __DEFER_NOTIFY = true;
    var out = handler(data);
    if (lock) { lock.releaseLock(); lock = null; __HAS_LOCK = false; }
    flushNotificationQueue();
    runAfterLockJobs();
    return out;
  } catch (err) {
    return jsonResponse({ success: false, error: String((err && err.message) || err) });
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (x) {} }
    __HAS_LOCK = false;
  }
}

function ensureSchemaIfNeeded() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty("SCHEMA_VERSION") === SCHEMA_VERSION) return;
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    if (props.getProperty("SCHEMA_VERSION") === SCHEMA_VERSION) return; // another request just finished it
    __TBL = {};
    ensureSchema(false);
  } finally {
    lock.releaseLock();
  }
}

function okData(user, message, extra) {
  var appData = getRoleFilteredAppData(user);
  if (user && !user.suspended && user.email) appData.user = publicUser(user);
  if (message) appData.message = message;
  if (extra) for (var k in extra) appData[k] = extra[k];
  if (__NOTIFY_WARNINGS.length) {
    appData.notificationWarning = true;
    var warning = __LANG === "es" ? "El cambio se guardó, pero algunos avisos no pudieron entrar en la cola. Revisa Admin > Ajustes y vuelve a publicar el aviso." : "Your change was saved, but some notifications could not be queued. Check Admin Settings and re-send the notice.";
    appData.message = (appData.message ? appData.message + " " : "") + warning;
  }
  return jsonResponse(appData);
}

function fail(msg, extra) {
  var o = { success: false, error: msg };
  if (extra) for (var k in extra) o[k] = extra[k];
  return jsonResponse(o);
}

// ============================================================================
// PIN-LOCKED ADMIN PUSH CONFIGURATION VAULT
// ============================================================================
function adminUnlockPushConfig(data) {
  var admin = requireAdmin(data);
  if (!admin) return fail(tr("errAdminOnly"));
  if (!verifyActualSixDigitPinForAdmin(admin.email, data.confirmPin)) return fail(tr("errVaultPin"));
  var props = PropertiesService.getScriptProperties();
  var appId = String(props.getProperty("ONESIGNAL_APP_ID") || "").trim();
  var restKey = String(props.getProperty("ONESIGNAL_REST_KEY") || "").trim();
  // The REST key never leaves the server: the browser only learns whether one is saved.
  return jsonResponse({
    success: true,
    oneSignalAppId: appId,
    restKeySet: !!restKey,
    restKeyHint: restKey ? "••••" + restKey.slice(-4) : ""
  });
}

function adminSavePushConfig(data) {
  var admin = requireAdmin(data);
  if (!admin) return fail(tr("errAdminOnly"));
  if (!verifyActualSixDigitPinForAdmin(admin.email, data.confirmPin)) return fail(tr("errVaultPinNoChange"));
  var appId = String(data.oneSignalAppId || "").trim();
  var restKey = String(data.oneSignalRestKey || "").trim();
  var props = PropertiesService.getScriptProperties();
  if (appId) props.setProperty("ONESIGNAL_APP_ID", appId); else props.deleteProperty("ONESIGNAL_APP_ID");
  if (restKey) props.setProperty("ONESIGNAL_REST_KEY", restKey);            // blank = keep the saved key
  if (data.clearRestKey) props.deleteProperty("ONESIGNAL_REST_KEY");
  var active = !!(props.getProperty("ONESIGNAL_APP_ID") && props.getProperty("ONESIGNAL_REST_KEY"));
  return okData(admin, active ? tr("msgPushSaved") : tr("msgPushCleared"));
}

// ============================================================================
// CAPTAIN & ALL-ADMIN FALLBACK ROUTING ENGINE
// ============================================================================
function getAllActiveAdminEmails() {
  var adminSet = {};
  PROTECTED_ADMINS.forEach(function(em) { adminSet[String(em).toLowerCase().trim()] = true; });
  var t = T("Profiles");
  for (var r = 0; r < t.rows.length; r++) {
    var em = String(tv(t, r, "Email") || "").toLowerCase().trim();
    var role = String(tv(t, r, "Role") || "Player").toLowerCase().trim();
    var st = String(tv(t, r, "AccountStatus") || "Active").toLowerCase().trim();
    if (em && role === "admin" && st !== "suspended") adminSet[em] = true;
  }
  return Object.keys(adminSet);
}

function getValidCaptainProfile(captainEmail) {
  var capEm = String(captainEmail || "").toLowerCase().trim();
  if (!capEm || isNonPlayingSystemAdmin(capEm)) return null;
  var r = findProfileRow(capEm);
  if (r === -1) return null;
  var t = T("Profiles");
  if (String(tv(t, r, "AccountStatus") || "Active").toLowerCase().trim() === "suspended") return null;
  return { email: capEm, fullName: String(tv(t, r, "FullName") || capEm) };
}

function getGameOrganizersRecipients(gameId) {
  var r = findGameRow(gameId);
  if (r !== -1 && gameFeaturesAt(T("Games"), r).captains) {
    var validCap = getValidCaptainProfile(tv(T("Games"), r, "CaptainEmail"));
    if (validCap) return [validCap.email];
  }
  return getAllActiveAdminEmails();
}

function verifyGameManagerRequest(data, gameId) {
  var user = verifyUserCredentials(data.requesterEmail, data.requesterPin);
  if (!user || user.suspended || user.rateLimited) return null;
  if (user.isAdmin) return user;
  var r = findGameRow(gameId);
  if (r === -1) return null;
  if (!gameFeaturesAt(T("Games"), r).captains) return null;
  var capEm = String(tv(T("Games"), r, "CaptainEmail") || "").toLowerCase().trim();
  if (capEm && capEm === user.email && getValidCaptainProfile(capEm)) return user;
  return null;
}

function appendGameActivityLog(gameRow, actorName, actionSummary) {
  var t = T("Games");
  var raw = tv(t, gameRow, "AdminActivityJSON");
  var list = [];
  if (raw) { try { list = JSON.parse(String(raw)); } catch (e) { list = []; } }
  list.unshift({ actor: String(actorName || "Admin"), summary: String(actionSummary || ""), timestamp: formatDateTimeStr(new Date()) });
  tw(t, gameRow, "AdminActivityJSON", JSON.stringify(list.slice(0, 15)));
}

// ============================================================================
// MAGIC EMAIL RSVP LANDING PAGE (Safe GET Preview + PIN-Verified POST Submit)
// ============================================================================
function generateMagicToken(email, gameId) {
  var secret = getAppSecret();
  var raw = "RSVP-v5.4|" + String(email).toLowerCase().trim() + "|" + String(gameId) + "|" + secret;
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw);
  return Utilities.base64EncodeWebSafe(digest).slice(0, 22);
}

// Email answer buttons open the app, which previews the choice and asks for confirmation: the signature proves
// the link came from that player's own email, so no login is needed.
function buildMagicRsvpUrl(email, gameId, status) {
  var em = String(email).toLowerCase().trim();
  return SITE_URL + "?rsvp=" + encodeURIComponent(status) +
         "&game=" + encodeURIComponent(gameId) +
         "&email=" + encodeURIComponent(em) +
         "&sig=" + encodeURIComponent(generateMagicToken(em, gameId));
}

// One-tap answer from an email button (no login). Returns just enough to confirm it.
function emailAnswer(data) {
  var email = String(data.email || "").toLowerCase().trim();
  var gameId = String(data.gameId || "");
  var status = String(data.status || "");
  if (["Yes", "Maybe", "No"].indexOf(status) === -1) return fail(tr("errBadStatus"));
  if (!email || !gameId || String(data.sig || "") !== generateMagicToken(email, gameId)) return fail(tr("errBadToken"), { badLink: true });
  var pr = findProfileRow(email);
  if (pr === -1) return fail(tr("errNoProfile"), { badLink: true });
  var pT = T("Profiles");
  if (!isProtectedAdmin(email) && String(tv(pT, pr, "AccountStatus") || "Active").toLowerCase().trim() === "suspended") return fail(tr("errSuspended"));
  if (isNonPlayingSystemAdmin(email)) return fail(tr("errSysAdminRsvp"));
  var gr = findGameRow(gameId);
  if (gr === -1) return fail(tr("rsvpGone"), { gone: true });
  var g = gameRecord(gr);
  var lock = gameLockState(g);
  if (lock.locked) return fail(lockedError(lock.reason), { gone: true });
  var result = processSignupCore(gameId, email, status);
  if (!result.success) return jsonResponse(result);
  var g2 = gameRecord(findGameRow(gameId));
  var yes = 0;
  var sT = T("Signups");
  for (var s2 = 0; s2 < sT.rows.length; s2++) {
    if (String(tv(sT, s2, "GameID")) === gameId && String(tv(sT, s2, "Status")) === "Yes") yes += 1 + parseGuestNamesList(tv(sT, s2, "PlusOne")).length;
  }
  var waitPos = 0;
  if (result.finalStatus === "Waitlist") {
    var wl = [];
    for (var w = 0; w < sT.rows.length; w++) {
      if (String(tv(sT, w, "GameID")) === gameId && String(tv(sT, w, "Status")) === "Waitlist") wl.push({ em: String(tv(sT, w, "Email")).toLowerCase().trim(), ts: tsMs(tv(sT, w, "Timestamp")) });
    }
    wl.sort(function(a, b) { return a.ts - b.ts; });
    waitPos = wl.findIndex(function(x) { return x.em === email; }) + 1;
  }
  var counts = typeof gameCommittedCounts === "function" ? gameCommittedCounts(gameId) : { reserved: yes, confirmed: yes, pending: 0 };
  return jsonResponse({
    success: true,
    finalStatus: result.finalStatus,
    needsReconfirm: typeof signupNeedsReconfirm === "function" && findSignupRow(gameId, email) !== -1 ? !!signupNeedsReconfirm(g2, findSignupRow(gameId, email)) : false,
    message: (result.message || "") + (__NOTIFY_WARNINGS.length ? (__LANG === "es" ? " El cambio se guardó, pero no se pudo poner en cola algún aviso. Avísale al organizador." : " Your answer was saved, but a notice could not be queued. Please tell the organizer.") : ""),
    notificationWarning: __NOTIFY_WARNINGS.length > 0,
    email: email,
    playerName: String(tv(pT, pr, "FullName") || ""),
    game: { GameID: gameId, Title: g2.title, Date: g2.date, Time: g2.time, LocationName: g2.location, yesCount: counts.reserved, confirmedCount: counts.confirmed, pendingReconfirmCount: counts.pending, MaxCap: g2.maxCap, MinThreshold: g2.minThreshold, waitPos: waitPos }
  });
}

function profileLang(email) {
  var r = findProfileRow(email);
  return r === -1 ? "en" : normalizeLang(tv(T("Profiles"), r, "Lang"));
}

// Links in emails sent before the buttons opened the app directly still land here:
// hand the player over to the app, which does the rest.
function renderMagicEmailConfirmationPage(params) {
  var email = String(params.email || "").toLowerCase().trim();
  var gameId = String(params.gameId || params.game || "");
  var status = String(params.status || params.rsvp || "Yes");
  if (["Yes", "Maybe", "No"].indexOf(status) === -1) status = "Yes";
  var supplied = String(params.token || params.sig || "");
  var valid = !!(email && gameId && supplied && supplied === generateMagicToken(email, gameId));
  var url = valid ? buildMagicRsvpUrl(email, gameId, status) : SITE_URL;
  var text = valid ? "Open Pitch Pulse to review and confirm your answer. / Abre Pitch Pulse para revisar y confirmar tu respuesta."
    : "This old or invalid answer link cannot change an RSVP. Open Pitch Pulse and log in, or use a new reminder. / Este enlace antiguo o no válido no puede cambiar una respuesta. Abre Pitch Pulse e inicia sesión o usa un recordatorio nuevo.";
  var html = "<!doctype html><html lang='en'><head><base target='_top'></head><body style='font-family:system-ui;padding:24px;max-width:440px;margin:40px auto;'>" +
    "<h1>Pitch Pulse</h1><p>" + escapeHtmlServer(text) + "</p><p><a href='" + escapeHtmlServer(url) + "' target='_top'>Open Pitch Pulse / Abrir Pitch Pulse</a></p></body></html>";
  return HtmlService.createHtmlOutput(html).setTitle("Pitch Pulse").addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function confirmMagicRsvpWithPin(data) {
  var email = String(data.email || "").toLowerCase().trim();
  var pin = String(data.pin || "").trim();
  var gameId = String(data.gameId || "");
  var status = String(data.status || "Yes");
  if (["Yes", "Maybe", "No"].indexOf(status) === -1) return fail(tr("errBadStatus"));
  if (String(data.token || "") !== generateMagicToken(email, gameId)) return fail(tr("errBadToken"));

  var user = verifyUserCredentials(email, pin, { issueSession: false });
  if (!user) return fail(tr("errBadPin"));
  if (user.suspended) return fail(tr("errSuspended"));
  if (user.rateLimited) return fail(user.error);
  if (isNonPlayingSystemAdmin(email)) return fail(tr("errSysAdminRsvp"));

  var result = processSignupCore(gameId, email, status);
  if (!result.success) return jsonResponse(result);
  var finalSt = result.finalStatus || status;
  var txt = finalSt === "Waitlist" ? tr("msgMagicWaitlist") : tr("msgMagicConfirmed", { status: tr("status" + finalSt) });
  return jsonResponse({ success: true, finalStatus: finalSt, confirmationText: txt });
}

// ============================================================================
// AUTHENTICATION & CREDENTIAL VERIFICATION
function previewEmailAnswer(data) {
  var email = String(data.email || "").toLowerCase().trim(), gameId = String(data.gameId || ""), status = String(data.status || "");
  if (["Yes", "Maybe", "No"].indexOf(status) === -1) return fail(tr("errBadStatus"));
  if (!email || !gameId || !data.sig || String(data.sig) !== generateMagicToken(email, gameId)) return fail(tr("errBadToken"), { badLink: true });
  var pr = findProfileRow(email); if (pr === -1) return fail(tr("errNoProfile"), { badLink: true });
  var pt = T("Profiles");
  if (!isProtectedAdmin(email) && String(tv(pt, pr, "AccountStatus")).toLowerCase() === "suspended") return fail(tr("errSuspended"));
  if (isNonPlayingSystemAdmin(email)) return fail(tr("errSysAdminRsvp"));
  var gr = findGameRow(gameId); if (gr === -1) return fail(tr("rsvpGone"), { gone: true });
  var g = gameRecord(gr), lock = gameLockState(g);
  if (lock.locked) return fail(lockedError(lock.reason), { gone: true });
  var st = T("Signups"), sr = findSignupRow(gameId, email), currentStatus = sr === -1 ? "" : String(tv(st, sr, "Status"));
  var guests = sr === -1 ? 0 : parseGuestNamesList(tv(st, sr, "PlusOne")).length;
  var yes = 0;
  for (var r = 0; r < st.rows.length; r++) if (String(tv(st, r, "GameID")) === gameId && String(tv(st, r, "Status")) === "Yes") yes += 1 + parseGuestNamesList(tv(st, r, "PlusOne")).length;
  var counts = typeof gameCommittedCounts === "function" ? gameCommittedCounts(gameId) : { reserved: yes, confirmed: yes, pending: 0 };
  var needsReconfirm = sr !== -1 && typeof signupNeedsReconfirm === "function" ? !!signupNeedsReconfirm(g, sr) : false;
  var waiting = [];
  for (var w = 0; w < st.rows.length; w++) if (String(tv(st, w, "GameID")) === gameId && String(tv(st, w, "Status")) === "Waitlist") waiting.push({ email: String(tv(st, w, "Email")).toLowerCase().trim(), at: tsMs(tv(st, w, "Timestamp")) });
  waiting.sort(function(a, b) { return a.at - b.at; });
  var waitPos = currentStatus === "Waitlist" ? waiting.findIndex(function(w) { return w.email === email; }) + 1 : 0;
  var start = typeof gameInstantMs === "function" ? gameInstantMs(g.date, g.time) : 0, now = new Date().getTime();
  return jsonResponse({ success: true, email: email, playerName: String(tv(pt, pr, "FullName")), currentStatus: currentStatus, guestCount: guests, needsReconfirm: needsReconfirm, waitPos: waitPos,
    lateDropApplies: !!(g.features.lateDrops && currentStatus === "Yes" && status !== "Yes" && start && now >= start - LATE_DROP_WINDOW_MINUTES * 60000 && now < start + 60 * 60000), isStarted: !!(start && now >= start),
    game: { GameID: gameId, Title: g.title, Date: g.date, Time: g.time, LocationName: g.location, yesCount: counts.reserved, confirmedCount: counts.confirmed,
      pendingReconfirmCount: counts.pending, waitPos: waitPos, MaxCap: g.maxCap, MinThreshold: g.minThreshold, startAtMs: start, endAtMs: start ? start + 75 * 60000 : 0 } });
}
// ============================================================================
function verifyUserCredentials(email, pinOrToken, opts) {
  opts = opts || {};
  var em = String(email || "").toLowerCase().trim();
  var cred = String(pinOrToken || "").trim();
  if (!em || !cred) return null;
  var t = T("Profiles");
  var r = findProfileRow(em);
  if (r === -1) return null;

  var fullName = String(tv(t, r, "FullName") || em);
  var role = String(tv(t, r, "Role") || "Player").trim();
  var accountStatus = String(tv(t, r, "AccountStatus") || "Active").trim();
  var position = String(tv(t, r, "Position") || "MID").toUpperCase().trim() || "MID";
  var cbg = tv(t, r, "CanBringGuests");
  var canBringGuests = !(cbg === false || String(cbg === "" ? "TRUE" : cbg).toLowerCase().trim() === "false");
  var awayUntil = formatDateIso(tv(t, r, "AwayUntil"));
  var notifyPref = normalizeNotifyPref(tv(t, r, "NotifyPref"));

  if (isProtectedAdmin(em)) {
    role = "Admin"; accountStatus = "Active"; canBringGuests = true;
  } else if (accountStatus.toLowerCase() === "suspended") {
    return { suspended: true };
  }

  var activeSessionToken = "";
  if (cred.indexOf("SES_") === 0) {
    if (!verifySessionToken(em, cred)) return null;
    activeSessionToken = cred;
  } else {
    var throttle = checkAuthThrottle(em);
    if (!throttle.allowed) return { rateLimited: true, error: throttle.error };
    if (!verifyStoredPinAndMigrateIfLegacy(r, em, tv(t, r, "PinCode"), cred)) {
      recordAuthFailure(em);
      return null;
    }
    clearAuthThrottle(em);
    if (opts.issueSession !== false) activeSessionToken = issueSessionToken(em);
  }

  return {
    email: em,
    pin: activeSessionToken,
    sessionToken: activeSessionToken,
    fullName: fullName,
    role: role,
    position: position,
    awayUntil: awayUntil || "",
    notifyPref: notifyPref,
    status: accountStatus,
    canBringGuests: canBringGuests,
    defaultTeam: normalizeTeamPref(tv(t, r, "DefaultTeam")),
    lang: normalizeLang(tv(t, r, "Lang")),
    isAdmin: (role.toLowerCase() === "admin"),
    isProtected: isProtectedAdmin(em),
    isSystemAdminOnly: isNonPlayingSystemAdmin(em),
    _row: r
  };
}

function publicUser(u) {
  var o = {};
  for (var k in u) if (k !== "_row") o[k] = u[k];
  return o;
}

function requireAdmin(data) {
  var user = verifyUserCredentials(data.requesterEmail, data.requesterPin);
  if (user && !user.suspended && !user.rateLimited && user.isAdmin) return user;
  return null;
}

function verifyAdminRequest(data) {
  return !!requireAdmin(data);
}

// Remembers the language the player uses, so emails, pushes and magic-link pages match it.
// The 60-second refresh runs without the global lock, so it only writes when it can grab the
// lock quickly (re-reading the row first); otherwise the next refresh saves it.
function rememberLang(user) {
  if (!user || user._row === undefined || user.lang === __LANG || !tHas(T("Profiles"), "Lang")) return;
  var write = function() {
    var r = findProfileRow(user.email);
    if (r !== -1) { tw(T("Profiles"), r, "Lang", __LANG); user.lang = __LANG; }
  };
  if (__HAS_LOCK) { try { write(); } catch (e) {} return; }
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1500)) return;
  try { __TBL = {}; write(); } catch (e) {} finally { lock.releaseLock(); }
}

// Explicit allowlist of what each role may see (nothing raw from the sheet leaks through).
// Phones get the upcoming matches plus the most recent finished ones, so the download stays
// small even with a match every weekday. Standings and stats were already counted over all of them.
function trimGamesForApp(raw) {
  var today = raw.serverToday || formatDateIso(new Date());
  var past = raw.games.filter(function(g) { return g.Date < today; });
  raw.pastMatchesTotal = past.length;
  if (past.length <= PAST_MATCHES_IN_APP) return raw;
  var keep = {};
  past.slice(-PAST_MATCHES_IN_APP).forEach(function(g) { keep[g.GameID] = true; }); // games are sorted oldest → newest
  raw.games = raw.games.filter(function(g) { return g.Date >= today || keep[g.GameID]; });
  return raw;
}

function getRoleFilteredAppData(userObj) {
  var raw = trimGamesForApp(getAppData());
  var isAdminUser = !!(userObj && userObj.isAdmin);
  var viewerEmail = userObj ? String(userObj.email || "").toLowerCase().trim() : "";
  if (isAdminUser) return raw;

  raw.invites = [];
  raw.recurring = [];
  delete raw.system;
  raw.games.forEach(function(g) {
    if (g.blasts && g.blasts.length) g.blasts = g.blasts.filter(function(b) { return !b.hidden; });
    var isGameCap = (g.captainEmail && g.captainEmail === viewerEmail);
    if (!isGameCap) {
      g.adminActivity = [];
      g.unrespondedNames = [];
    }
  });
  Object.keys(raw.profiles).forEach(function(em) {
    if (em === viewerEmail) return;
    var p = raw.profiles[em];
    delete p.notifyPref; delete p.awayUntil; delete p.lateEarlyCount; delete p.lang; delete p.canBringGuests;
  });
  return maskEmailsForPlayer(raw, viewerEmail);
}

// Players never receive anyone else's email address: every other person is identified by a
// stable private id ("p_…") that only the server can map back.
var __PID_CACHE = null;
function playerPid(email) {
  var em = String(email || "").toLowerCase().trim();
  if (!em) return "";
  if (!__PID_CACHE) __PID_CACHE = {};
  if (__PID_CACHE[em]) return __PID_CACHE[em];
  var d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, "PID|" + em + "|" + getAppSecret());
  var pid = "p_" + Utilities.base64EncodeWebSafe(d).replace(/[^A-Za-z0-9]/g, "").slice(0, 12);
  __PID_CACHE[em] = pid;
  return pid;
}
// Turns a private id (or an email) back into the email; guest keys keep their "#name" part.
function resolvePlayerRef(ref) {
  var v = String(ref || "").trim();
  if (v.indexOf("p_") !== 0) return v.toLowerCase();
  var hash = v.indexOf("#");
  var pid = hash === -1 ? v : v.slice(0, hash);
  var rest = hash === -1 ? "" : v.slice(hash);
  var t = T("Profiles");
  for (var r = 0; r < t.rows.length; r++) {
    var em = String(tv(t, r, "Email") || "").toLowerCase().trim();
    if (em && playerPid(em) === pid) return em + rest.toLowerCase();
  }
  return v.toLowerCase();
}
function maskEmailsForPlayer(raw, viewerEmail) {
  var mask = function(em) {
    em = String(em || "").toLowerCase().trim();
    if (!em || em === viewerEmail) return em;
    return playerPid(em);
  };
  var maskKey = function(key) {
    var k = String(key || "");
    var i = k.indexOf("#");
    return i === -1 ? mask(k) : mask(k.slice(0, i)) + k.slice(i);
  };
  var maskEntity = function(e) {
    if (!e) return e;
    if (e.email) e.email = mask(e.email);
    if (e.key) e.key = maskKey(e.key);
    if (e.hostEmail) e.hostEmail = mask(e.hostEmail);
    return e;
  };
  var profiles = {};
  Object.keys(raw.profiles).forEach(function(em) {
    var p = raw.profiles[em];
    var id = mask(em);
    p.email = id;
    profiles[id] = p;
  });
  raw.profiles = profiles;
  raw.games.forEach(function(g) {
    (g.roster || []).forEach(function(r) {
      r.email = mask(r.email);
      if (r.motmVote) r.motmVote = mask(r.motmVote);
    });
    if (g.teams) { (g.teams.white || []).forEach(maskEntity); (g.teams.dark || []).forEach(maskEntity); }
    if (g.captainEmail) g.captainEmail = mask(g.captainEmail);
    if (g.motmWinnerEmail) g.motmWinnerEmail = mask(g.motmWinnerEmail);
    if (g.motmTally) {
      var tally = {};
      Object.keys(g.motmTally).forEach(function(em) { tally[mask(em)] = g.motmTally[em]; });
      g.motmTally = tally;
    }
    (g.blasts || []).forEach(function(b) { delete b.senderEmail; });
  });
  return raw;
}

function getLeagueDataAuthenticated(data) {
  var user = verifyUserCredentials(data.requesterEmail, data.requesterPin);
  if (!user) return fail(tr("errSession"), { authError: true });
  if (user.rateLimited) return fail(user.error, { authError: true });
  if (user.suspended) return fail(tr("errSuspended"), { authError: true });
  rememberLang(user);
  return okData(user, "");
}

function generateInviteToken(email) {
  var secret = getAppSecret();
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, "INVITE|" + String(email).toLowerCase().trim() + "|" + secret);
  return Utilities.base64EncodeWebSafe(digest).slice(0, 20);
}

function buildInviteUrl(email) {
  return SITE_URL + "?invite=" + encodeURIComponent(generateInviteToken(email)) + "&email=" + encodeURIComponent(String(email).toLowerCase().trim());
}

function registerUser(data) {
  var email = String(data.email || "").toLowerCase().trim();
  var pin = String(data.pin || "").trim();
  var fullName = cleanPersonName(data.fullName);
  var position = String(data.position || "MID").toUpperCase().trim();
  if (["GK", "DEF", "MID", "FWD"].indexOf(position) === -1) position = "MID";

  if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || fullName.length < 2 || !isValidSixDigitPin(pin)) return fail(tr("errRegisterFields"));

  var invT = T("Invites");
  var invRow = tfind(invT, function(r) { return String(tv(invT, r, "Email")).toLowerCase().trim() === email; });
  var isProtected = isProtectedAdmin(email);
  if (invRow === -1 && !isProtected) return fail(tr("errInviteRequired", { email: email }));
  // Proof of email ownership: the invite link carries a token only that inbox received.
  if (String(data.inviteToken || "") !== generateInviteToken(email)) return fail(tr("errInviteLink"), { needsInviteLink: true });
  if (findProfileRow(email) !== -1) return fail(tr("errAccountExists"));

  tadd(T("Profiles"), {
    Email: email, PinCode: createPinHashRecord(email, pin), FullName: fullName, JoinedDate: "'" + formatDateIso(new Date()),
    Role: isProtected ? "Admin" : "Player", AccountStatus: "Active", Position: position, LateDrops: 0, CanBringGuests: true,
    AwayUntil: "", LateEarlyCount: 0, NotifyPref: "Both", DefaultTeam: "", Lang: __LANG
  });
  if (invRow !== -1) tw(invT, invRow, "Status", "Registered");
  else tadd(invT, { Email: email, Status: "Registered", InvitedDate: "'" + formatDateIso(new Date()) });

  var user = verifyUserCredentials(email, pin);
  var appData = getRoleFilteredAppData(user);
  return jsonResponse({ success: true, user: publicUser(user), appData: appData });
}

function loginUser(data) {
  var email = String(data.email || "").toLowerCase().trim();
  var pin = String(data.pin || "").trim();
  if (!isValidSixDigitPin(pin)) return fail(tr("errPinFormat"));
  var user = verifyUserCredentials(email, pin);
  if (!user) return fail(tr("errLogin"));
  if (user.rateLimited) return fail(user.error);
  if (user.suspended) return fail(tr("errSuspended"));
  rememberLang(user);
  return jsonResponse({ success: true, user: publicUser(user), appData: getRoleFilteredAppData(user) });
}

function logoutUser(data) {
  var em = String(data.requesterEmail || "").toLowerCase().trim();
  var tk = String(data.requesterPin || "");
  if (em && tk.indexOf("SES_") === 0 && verifySessionToken(em, tk)) revokeSessionToken(em, tk);
  return jsonResponse({ success: true });
}

function updateProfileDetails(data) {
  var user = verifyUserCredentials(data.requesterEmail, data.requesterPin);
  if (!user || user.suspended || user.rateLimited) return fail(tr("errUnauthorized"));

  var newPos = String(data.position || user.position || "MID").toUpperCase().trim();
  if (["GK", "DEF", "MID", "FWD"].indexOf(newPos) === -1) newPos = "MID";
  var newName = data.fullName !== undefined ? cleanPersonName(data.fullName) : user.fullName;
  if (newName.length < 2) newName = user.fullName;
  var newAwayUntil = data.awayUntil !== undefined ? formatDateIso(data.awayUntil) : user.awayUntil;
  var newNotifyPref = data.notifyPref !== undefined ? normalizeNotifyPref(data.notifyPref) : (user.notifyPref || "Both");

  var t = T("Profiles");
  var fields = { FullName: newName, Position: newPos, AwayUntil: newAwayUntil ? "'" + newAwayUntil : "", NotifyPref: newNotifyPref };
  if (tHas(t, "Lang")) fields.Lang = __LANG;
  twMany(t, user._row, fields);

  user.position = newPos; user.fullName = newName; user.awayUntil = newAwayUntil || ""; user.notifyPref = newNotifyPref; user.lang = __LANG;
  return okData(user, tr("msgProfileSaved"));
}

function requestPinReset(data) {
  var inputEmail = String(data.email || "").toLowerCase().trim();
  var useBackupRecovery = !!data.useBackupRecovery;
  var t = T("Profiles");
  var targetProfileEmail = "", foundName = "";

  if (useBackupRecovery) {
    if (!isProtectedAdmin(inputEmail)) return fail(tr("errBackupOnly", { backup: BACKUP_RECOVERY_EMAIL, primary: PRIMARY_ADMIN_EMAIL }));
    var r1 = findProfileRow(inputEmail);
    if (r1 === -1) r1 = findProfileRow(PRIMARY_ADMIN_EMAIL);
    if (r1 === -1) return fail(tr("errPrimaryNotRegistered"));
    targetProfileEmail = String(tv(t, r1, "Email")).toLowerCase().trim();
    foundName = String(tv(t, r1, "FullName") || "Primary System Admin");
  } else {
    var r2 = findProfileRow(inputEmail);
    if (r2 !== -1) { targetProfileEmail = inputEmail; foundName = String(tv(t, r2, "FullName")); }
  }
  if (!targetProfileEmail) return fail(tr("errNoProfile"));

  // Throttle: one code per minute, five per day per account (protects the daily email quota).
  var props = PropertiesService.getScriptProperties();
  var rlKey = "RESETREQ_" + targetProfileEmail;
  var today = formatDateIso(new Date());
  var nowMs = new Date().getTime();
  var rl = { day: today, count: 0, lastAt: 0 };
  try { var prev = JSON.parse(props.getProperty(rlKey) || "null"); if (prev && prev.day === today) rl = prev; } catch (e) {}
  if (rl.lastAt && nowMs - rl.lastAt < PIN_RESET_COOLDOWN_SECONDS * 1000) {
    return fail(tr("errResetCooldown", { secs: Math.ceil((PIN_RESET_COOLDOWN_SECONDS * 1000 - (nowMs - rl.lastAt)) / 1000) }));
  }
  if (rl.count >= PIN_RESET_MAX_PER_DAY) return fail(tr("errResetDaily"));

  var tempCode = String(Math.floor(100000 + Math.random() * 900000));
  props.setProperty("RESET_" + targetProfileEmail, JSON.stringify({ code: tempCode, expiresAt: nowMs + 15 * 60 * 1000, attempts: 0 }));
  var recipientEmail = useBackupRecovery ? BACKUP_RECOVERY_EMAIL : targetProfileEmail;
  var es = __LANG === "es";
  try {
    MailApp.sendEmail({
      to: recipientEmail,
      name: MAIL_SENDER_NAME,
      subject: useBackupRecovery
        ? "🚨 [BACKUP ADMIN RECOVERY] Pitch Pulse League PIN Reset Code"
        : (es ? "⚽ Pitch Pulse: tu código para restablecer el PIN" : "⚽ Pitch Pulse League: Your 6-Digit PIN Reset Code"),
      htmlBody: "<div style='font-family:sans-serif;padding:24px;background:#0d1612;color:#fff;border-radius:12px;'>" +
        "<h2 style='color:#00f59b;margin-top:0;'>" + (useBackupRecovery ? "Emergency Admin PIN Recovery" : (es ? "Restablecer PIN" : "PIN Reset Verification")) + "</h2>" +
        "<p>" + (es ? "Hola " : "Hi ") + escapeHtmlServer(foundName) + ",</p>" +
        "<p>" + (es ? "Se pidió restablecer el PIN de la cuenta: " : "A PIN reset was requested for account: ") + "<b>" + escapeHtmlServer(targetProfileEmail) + "</b></p>" +
        "<p>" + (es ? "Tu código temporal (válido 15 minutos) es:" : "Your temporary 6-digit verification code (valid for 15 minutes) is:") + "</p>" +
        "<div style='font-size:28px;font-weight:bold;letter-spacing:6px;color:#00f59b;padding:12px 0;'>" + tempCode + "</div>" +
        "<p>" + (es ? "Escríbelo en la página de la liga junto con tu nuevo PIN de 6 dígitos. Si no lo pediste, ignora este correo." : "Enter this code on the league website along with your new 6-digit PIN. If you didn't ask for this, you can ignore this email.") + "</p></div>"
    });
  } catch (mailErr) {
    return fail(tr("errResetMail", { err: mailErr.toString() }));
  }
  rl.count++; rl.lastAt = nowMs;
  props.setProperty(rlKey, JSON.stringify(rl));
  return jsonResponse({
    success: true,
    targetEmail: targetProfileEmail,
    message: useBackupRecovery ? tr("msgResetSentBackup", { backup: BACKUP_RECOVERY_EMAIL, email: targetProfileEmail }) : tr("msgResetSent", { email: recipientEmail })
  });
}

function resetPin(data) {
  var email = String(data.email || "").toLowerCase().trim();
  var code = String(data.code || "").trim();
  var newPin = String(data.newPin || "").trim();
  if (!isValidSixDigitPin(newPin)) return fail(tr("errNewPinFormat"));

  var props = PropertiesService.getScriptProperties();
  var rawSaved = props.getProperty("RESET_" + email);
  if (!rawSaved) return fail(tr("errNoResetCode"));
  var resetObj = null;
  try { resetObj = JSON.parse(rawSaved); } catch (e) { resetObj = { code: rawSaved, expiresAt: new Date().getTime() + 60000, attempts: 0 }; }
  if (new Date().getTime() > resetObj.expiresAt) {
    props.deleteProperty("RESET_" + email);
    return fail(tr("errResetExpired"));
  }
  if (resetObj.code !== code) {
    resetObj.attempts = (resetObj.attempts || 0) + 1;
    if (resetObj.attempts >= 5) {
      props.deleteProperty("RESET_" + email);
      return fail(tr("errResetTooMany"));
    }
    props.setProperty("RESET_" + email, JSON.stringify(resetObj));
    return fail(tr("errResetBadCode"));
  }
  var r = findProfileRow(email);
  if (r === -1) return fail(tr("errNoProfile"));
  tw(T("Profiles"), r, "PinCode", createPinHashRecord(email, newPin));
  props.deleteProperty("RESET_" + email);
  clearAuthThrottle(email);
  revokeAllSessions(email); // a new PIN signs every old device out
  return jsonResponse({ success: true, message: tr("msgPinUpdated", { email: email }) });
}

// ============================================================================
// NAMED GUEST PARSER & ROSTER HASH HELPERS
// ============================================================================
function parseGuestNamesList(rawVal) {
  if (rawVal === undefined || rawVal === null || rawVal === false || rawVal === "") return [];
  var str = String(rawVal).trim();
  if (!str || str.toLowerCase() === "false" || str === "0") return [];
  if (str.toLowerCase() === "true" || str === "1") return ["Guest 1"];
  return str.split(",").map(function(s) { return s.trim(); }).filter(Boolean).slice(0, MAX_GUESTS_PER_PLAYER);
}

// What position someone plays in a given match. With goalie tracking on, the per-match
// "Goalie" toggle decides (keepers start switched on and can opt out); with it off,
// everyone simply plays their own position.
function effectivePlayPosition(profilePos, playGkFlag, trackGK) {
  var pos = String(profilePos || "MID").toUpperCase().trim() || "MID";
  if (!trackGK) return pos;
  if (playGkFlag) return "GK";
  return pos === "GK" ? "MID" : pos;
}

// Keepers are down as goalie by default when they join a match (if that match tracks it).
function defaultPlayGkFor(email, gameId) {
  var trackGK = gameId ? gameFeatures(gameId).gk : getLeagueSettings().trackGK;
  if (!trackGK) return false;
  var r = findProfileRow(email);
  return r !== -1 && String(tv(T("Profiles"), r, "Position") || "").toUpperCase().trim() === "GK";
}

// Fingerprint of who is playing (and as what) — used to tell organizers the roster moved on
// since the lineup was last set. Built the same way from the app roster and lineup entities.
function computeRosterSignature(yesRoster) {
  return yesRoster.map(function(r) {
    return r.email + ":" + (r.playPosition || r.position || "MID") + ":" + (r.guestNames || []).join("+");
  }).sort().join("|");
}
function signatureFromEntities(entities) {
  var guestsByHost = {};
  entities.forEach(function(e) { if (e.isGuest) (guestsByHost[e.hostEmail] = guestsByHost[e.hostEmail] || []).push(e.fullName); });
  return entities.filter(function(e) { return !e.isGuest; }).map(function(e) {
    return e.email + ":" + (e.position || "MID") + ":" + (guestsByHost[e.email] || []).join("+");
  }).sort().join("|");
}

// The league plays in California. Device and Apps Script project timezones must not
// move kickoff, the late-drop window, or dates around daylight-saving transitions.
function leagueTimeZone() {
  return LEAGUE_TIME_ZONE;
}
function leagueNowMinutes(now) {
  var parts = Utilities.formatDate(now || new Date(), leagueTimeZone(), "HH:mm").split(":");
  return Number(parts[0]) * 60 + Number(parts[1]);
}
function gameInstantMs(dateStr, timeStr) {
  var date = normalizeDateInput(dateStr), time = normalizeTimeInput(timeStr);
  if (!date || !time) return NaN;
  try { return Utilities.parseDate(date + " " + time, leagueTimeZone(), "yyyy-MM-dd h:mm a").getTime(); }
  catch (e) { return NaN; }
}
function signupNeedsReconfirm(game, signupRow) {
  var sT = T("Signups"), status = String(tv(sT, signupRow, "Status"));
  return (status === "Yes" || status === "Maybe") && game.scheduleChangeAt > 0 &&
    Math.max(tsMs(tv(sT, signupRow, "Timestamp")), tsMs(tv(sT, signupRow, "ReconfirmedAt"))) < game.scheduleChangeAt;
}
// Reserved places remain held after a schedule change. Only acknowledged places
// count toward the minimum needed to announce that the revised match is on.
function gameCommittedCounts(gameId) {
  var gr = findGameRow(gameId);
  if (gr === -1) return { reserved: 0, confirmed: 0, pending: 0, pendingPlayers: 0 };
  var g = gameRecord(gr), sT = T("Signups");
  var result = { reserved: 0, confirmed: 0, pending: 0, pendingPlayers: 0 };
  for (var r = 0; r < sT.rows.length; r++) {
    if (String(tv(sT, r, "GameID")) !== String(gameId) || String(tv(sT, r, "Status")) !== "Yes") continue;
    var spots = 1 + parseGuestNamesList(tv(sT, r, "PlusOne")).length;
    result.reserved += spots;
    if (signupNeedsReconfirm(g, r)) { result.pending += spots; result.pendingPlayers++; }
    else result.confirmed += spots;
  }
  return result;
}
// A name match is a conflict to resolve, never proof that two identities should
// be merged. This also protects legacy guests that later register as members.
function guestReservationForMember(gameId, email) {
  var pr = findProfileRow(email);
  if (pr === -1) return null;
  var name = cleanPersonName(tv(T("Profiles"), pr, "FullName")).toLowerCase();
  if (!name) return null;
  var sT = T("Signups");
  for (var r = 0; r < sT.rows.length; r++) {
    var hostEmail = String(tv(sT, r, "Email")).toLowerCase().trim();
    if (String(tv(sT, r, "GameID")) !== String(gameId) || String(tv(sT, r, "Status")) !== "Yes" || hostEmail === email) continue;
    var names = parseGuestNamesList(tv(sT, r, "PlusOne"));
    if (names.some(function(n) { return cleanPersonName(n).toLowerCase() === name; })) {
      var hp = findProfileRow(hostEmail);
      return { hostName: hp === -1 ? "a player" : String(tv(T("Profiles"), hp, "FullName")), guestName: String(tv(T("Profiles"), pr, "FullName")) };
    }
  }
  return null;
}

function hasGameStartedOrPassed(gameDateStr, gameTimeStr) {
  var start = gameInstantMs(gameDateStr, gameTimeStr);
  return isFinite(start) && new Date().getTime() >= start;
}

function normalizeTeamPref(v) {
  var s = String(v || "").trim().toUpperCase();
  if (s === "A" || s === "WHITE") return "A";
  if (s === "B" || s === "DARK" || s === "BLACK") return "B";
  return "";
}

// ============================================================================
// TEAMS ENGINE (global team names, default home teams, smart balance, auto-sync)
// ============================================================================
function guestKey(hostEmail, guestName) {
  return String(hostEmail).toLowerCase().trim() + "#" + String(guestName).toLowerCase().trim();
}

// Everyone who is currently playing, as lineup entities (players + their named guests).
function lineupEntitiesForGame(gameId) {
  var sT = T("Signups"), pT = T("Profiles");
  var prof = {};
  for (var p = 0; p < pT.rows.length; p++) {
    var pem = String(tv(pT, p, "Email")).toLowerCase().trim();
    if (pem) prof[pem] = { name: String(tv(pT, p, "FullName") || pem), pos: String(tv(pT, p, "Position") || "MID").toUpperCase().trim() || "MID", team: normalizeTeamPref(tv(pT, p, "DefaultTeam")) };
  }
  var rows = [];
  for (var r = 0; r < sT.rows.length; r++) {
    var att = String(tv(sT, r, "AttendanceStatus"));
    if (String(tv(sT, r, "GameID")) === String(gameId) && att !== "NoShow" &&
        (att === "Attended" || String(tv(sT, r, "Status")) === "Yes")) rows.push(r);
  }
  rows.sort(function(a, b) { return tsMs(tv(sT, a, "Timestamp")) - tsMs(tv(sT, b, "Timestamp")); });
  var out = [];
  var trackGK = gameFeatures(gameId).gk;
  rows.forEach(function(r) {
    var em = String(tv(sT, r, "Email")).toLowerCase().trim();
    var pr = prof[em] || { name: em, pos: "MID", team: "" };
    out.push({ key: em, email: em, fullName: pr.name, position: effectivePlayPosition(pr.pos, isTruthy(tv(sT, r, "PlayGK")), trackGK), isGuest: false, hostEmail: "", defaultTeam: pr.team });
    parseGuestNamesList(tv(sT, r, "PlusOne")).forEach(function(gn) {
      out.push({ key: guestKey(em, gn), email: "", fullName: gn, position: "MID", isGuest: true, hostEmail: em, hostName: pr.name, defaultTeam: "" });
    });
  });
  return out;
}

function tsMs(v) {
  if (!v) return 0;
  var n = new Date(v).getTime();
  return isNaN(n) ? 0 : n;
}

// Accepts v4.x lineups ({fullName,email,position} with "host_guest_N" emails) and v5 entities.
function normalizeTeamEntry(x) {
  x = x || {};
  if (x.key) return x;
  var em = String(x.email || "").toLowerCase().trim();
  var gm = em.match(/^(.*)_guest_(\d+)$/);
  if (gm) {
    var nm = String(x.fullName || "").replace(/\s*\((Guest of|Invitado de)[^)]*\)\s*$/i, "").trim() || ("Guest " + (Number(gm[2]) + 1));
    return { key: guestKey(gm[1], nm), email: "", fullName: nm, position: "MID", isGuest: true, hostEmail: gm[1] };
  }
  return { key: em, email: em, fullName: String(x.fullName || em), position: String(x.position || "MID"), isGuest: false, hostEmail: "" };
}

function parseTeamsJson(raw) {
  if (!raw) return null;
  var o = null;
  try { o = JSON.parse(String(raw)); } catch (e) { return null; }
  if (!o || !Array.isArray(o.white) || !Array.isArray(o.dark)) return null;
  o.white = o.white.map(normalizeTeamEntry);
  o.dark = o.dark.map(normalizeTeamEntry);
  return o;
}

function strengthOf(ent, standings) {
  if (ent.isGuest) return 1.0;
  var s = standings[ent.email];
  var ppg = (s && s.games >= 3) ? (s.points / s.games) : 1.2;
  return ppg;
}

// Light-weight standings snapshot used for balancing (points per game from scored matches).
function quickStandings() {
  var gT = T("Games");
  var out = {};
  for (var r = 0; r < gT.rows.length; r++) {
    if (String(tv(gT, r, "Status")) === "Cancelled") continue;
    var sw = tv(gT, r, "ScoreWhite"), sd = tv(gT, r, "ScoreDark");
    if (sw === "" || sd === "") continue;
    if (!gameFeaturesAt(gT, r).scores) continue;
    var teams = parseTeamsJson(tv(gT, r, "TeamsJSON"));
    if (!teams) continue;
    sw = Number(sw); sd = Number(sd);
    [["white", sw, sd], ["dark", sd, sw]].forEach(function(side) {
      teams[side[0]].forEach(function(p) {
        if (p.isGuest || !p.email) return;
        var o = out[p.email] || (out[p.email] = { games: 0, points: 0 });
        o.games++;
        o.points += side[1] > side[2] ? 3 : (side[1] === side[2] ? 1 : 0);
      });
    });
  }
  return out;
}

function sideSum(list, st) { return list.reduce(function(n, e) { return n + strengthOf(e, st); }, 0); }
function gkCount(list) { return list.filter(function(e) { return e.position === "GK"; }).length; }

function pickSide(white, dark, st, ent) {
  if (white.length !== dark.length) return white.length < dark.length ? "white" : "dark";
  if (ent && ent.position === "GK" && gkCount(white) !== gkCount(dark)) return gkCount(white) < gkCount(dark) ? "white" : "dark";
  return sideSum(white, st) <= sideSum(dark, st) ? "white" : "dark";
}

function cleanEntity(e) {
  var o = { key: e.key, email: e.email || "", fullName: e.fullName, position: e.position || "MID", isGuest: !!e.isGuest, hostEmail: e.hostEmail || "" };
  if (e.isGuest && e.hostName) o.hostName = e.hostName;
  return o;
}

// Smart balance: spread positions (GK → DEF → MID → FWD → guests), then strength.
function balanceSmart(entities, st) {
  var buckets = { GK: [], DEF: [], MID: [], FWD: [], GUEST: [] };
  entities.forEach(function(e) {
    var b = e.isGuest ? "GUEST" : (buckets[e.position] ? e.position : "MID");
    buckets[b].push({ e: e, s: strengthOf(e, st) + (Math.random() * 0.3 - 0.15) });
  });
  var white = [], dark = [];
  ["GK", "DEF", "MID", "FWD", "GUEST"].forEach(function(b) {
    buckets[b].sort(function(x, y) { return y.s - x.s; });
    buckets[b].forEach(function(x) {
      var side = pickSide(white, dark, st, x.e);
      (side === "white" ? white : dark).push(cleanEntity(x.e));
    });
  });
  return { white: white, dark: dark };
}

// Default squads: everyone with a home team stays there; unassigned players and guests fill
// the gaps; then the fewest possible moves are made so sides differ by at most one player
// and each side has a goalkeeper when there are two or more.
function balanceDefaults(entities, st) {
  var white = [], dark = [], rest = [];
  entities.forEach(function(e) {
    if (!e.isGuest && e.defaultTeam === "A") white.push(cleanEntity(e));
    else if (!e.isGuest && e.defaultTeam === "B") dark.push(cleanEntity(e));
    else rest.push(e);
  });
  rest.sort(function(a, b) { return strengthOf(b, st) - strengthOf(a, st); });
  rest.forEach(function(e) { (pickSide(white, dark, st, e) === "white" ? white : dark).push(cleanEntity(e)); });
  var moves = 0;
  while (Math.abs(white.length - dark.length) > 1 && moves < 40) {
    var big = white.length > dark.length ? white : dark;
    var small = big === white ? dark : white;
    var target = (sideSum(big, st) - sideSum(small, st)) / 2;
    var bestIdx = -1, bestGap = Infinity;
    for (var i = 0; i < big.length; i++) {
      if (big[i].position === "GK" && gkCount(big) <= 1) continue;
      var gap = Math.abs(strengthOf(big[i], st) - target);
      if (gap < bestGap) { bestGap = gap; bestIdx = i; }
    }
    if (bestIdx === -1) bestIdx = big.length - 1;
    small.push(big.splice(bestIdx, 1)[0]);
    moves++;
  }
  fixGoalkeepers(white, dark, st);
  return { white: white, dark: dark };
}

function fixGoalkeepers(white, dark, st) {
  var gw = gkCount(white), gd = gkCount(dark);
  if ((gw === 0 && gd >= 2) || (gd === 0 && gw >= 2)) {
    var from = gw === 0 ? dark : white, to = gw === 0 ? white : dark;
    var gi = from.findIndex(function(e) { return e.position === "GK"; });
    var gk = from[gi];
    var swapIdx = -1, bestGap = Infinity;
    for (var i = 0; i < to.length; i++) {
      if (to[i].position === "GK") continue;
      var gap = Math.abs(strengthOf(to[i], st) - strengthOf(gk, st));
      if (gap < bestGap) { bestGap = gap; swapIdx = i; }
    }
    from.splice(gi, 1);
    to.push(gk);
    if (swapIdx !== -1) from.push(to.splice(swapIdx, 1)[0]);
  }
}

// Keeps a saved lineup in step with the roster: people who drop out (and their guests)
// leave the lineup; late joiners land on their home team if that keeps it fair, else the
// smaller side. Returns true when the lineup changed.
function syncTeamsForGame(gameRow) {
  var gT = T("Games");
  if (!gameFeaturesAt(gT, gameRow).teams) return false;
  var teams = parseTeamsJson(tv(gT, gameRow, "TeamsJSON"));
  if (!teams) return false;
  var gameId = String(tv(gT, gameRow, "GameID"));
  var entities = lineupEntitiesForGame(gameId);
  var byKey = {};
  entities.forEach(function(e) { byKey[e.key] = e; });
  var changed = false;
  ["white", "dark"].forEach(function(side) {
    var kept = [];
    teams[side].forEach(function(p) {
      var e = byKey[p.key];
      if (!e) { changed = true; return; }
      var fresh = cleanEntity(e);
      if (fresh.fullName !== p.fullName || fresh.position !== p.position) changed = true;
      kept.push(fresh);
    });
    teams[side] = kept;
  });
  var placed = {};
  teams.white.concat(teams.dark).forEach(function(p) { placed[p.key] = true; });
  var st = null;
  entities.forEach(function(e) {
    if (placed[e.key]) return;
    st = st || quickStandings();
    var pref = e.isGuest ? "" : (e.defaultTeam === "A" ? "white" : (e.defaultTeam === "B" ? "dark" : ""));
    var side = pickSide(teams.white, teams.dark, st, e);
    if (pref) {
      var other = pref === "white" ? "dark" : "white";
      if (teams[pref].length + 1 - teams[other].length <= 1) side = pref;
    }
    teams[side].push(cleanEntity(e));
    placed[e.key] = true;
    changed = true;
  });
  if (changed) tw(gT, gameRow, "TeamsJSON", JSON.stringify(teams));
  return changed;
}

// ============================================================================
// CORE DATA ENGINE (Includes Season W/D/L Standings, Rivalry Stats & Push Status)
// ============================================================================
function getAppData() {
  var settings = getLeagueSettings();
  var gT = T("Games"), sT = T("Signups"), lT = T("Locations"), pT = T("Profiles"), iT = T("Invites"), rT = T("Recurring");
  var bT = null;
  try { bT = T("Blasts"); } catch (e) { bT = null; }

  var props = PropertiesService.getScriptProperties();
  // Empty means "the standard rules": each phone shows them in its own language.
  var rawRules = String(props.getProperty("HOW_WE_PLAY_RULES") || "").trim();
  var howWePlayRules = rawRules === DEFAULT_HOW_WE_PLAY_RULES.trim() ? "" : rawRules;
  var oneSignalAppId = String(props.getProperty("ONESIGNAL_APP_ID") || "").trim();
  var pushConfigured = !!(oneSignalAppId && props.getProperty("ONESIGNAL_REST_KEY"));
  var now = new Date();
  var todayStr = formatDateIso(now);

  var games = [];
  var gameMap = {};
  var usedLocationsInHistory = {};
  for (var gr = 0; gr < gT.rows.length; gr++) {
    var gid = String(tv(gT, gr, "GameID") || "");
    if (!gid) continue;
    var rawW = tv(gT, gr, "ScoreWhite"), rawD = tv(gT, gr, "ScoreDark");
    var g = {
      GameID: gid,
      Title: String(tv(gT, gr, "Title")),
      LocationName: String(tv(gT, gr, "LocationName")),
      Date: formatDateIso(tv(gT, gr, "Date")),
      Time: formatTimeStr(tv(gT, gr, "Time")),
      MaxCap: Number(tv(gT, gr, "MaxCap") || settings.defaultMax),
      MinThreshold: Number(tv(gT, gr, "MinThreshold") || settings.defaultMin),
      Status: String(tv(gT, gr, "Status") || "Pending"),
      SeriesID: String(tv(gT, gr, "SeriesID") || ""),
      ScoreWhite: (rawW !== "" && rawW !== null) ? Number(rawW) : "",
      ScoreDark: (rawD !== "" && rawD !== null) ? Number(rawD) : "",
      ScheduleChangeAt: 0,
      ScheduleChangeSummary: String(tv(gT, gr, "ScheduleChangeSummary") || ""),
      MatchFormat: String(tv(gT, gr, "MatchFormat") || "Standard"),
      CaptainEmail: String(tv(gT, gr, "CaptainEmail") || "").toLowerCase().trim()
    };
    var sca = tv(gT, gr, "ScheduleChangeAt");
    g.ScheduleChangeAt = sca ? Number(new Date(sca).getTime() || sca || 0) : 0;
    g.startAtMs = gameInstantMs(g.Date, g.Time);
    g.endAtMs = g.startAtMs + MATCH_DURATION_MINUTES * 60000;
    g.hasStartedOrPassed = isFinite(g.startAtMs) && now.getTime() >= g.startAtMs;
    g.hasEnded = isFinite(g.endAtMs) && now.getTime() >= g.endAtMs;
    g.isPast = g.Date < todayStr;
    g.isToday = g.Date === todayStr;
    g.features = gameFeaturesAt(gT, gr);
    if (!g.features.scores) { g.ScoreWhite = ""; g.ScoreDark = ""; }
    g.hasScore = g.ScoreWhite !== "" && g.ScoreDark !== "";
    var lock = g.Status === "Cancelled" ? { locked: true, reason: "cancelled" } : (g.isPast ? { locked: true, reason: "past" } : (g.hasScore ? { locked: true, reason: "scored" } : { locked: false, reason: "" }));
    g.rsvpLocked = lock.locked;
    g.lockReason = lock.reason;
    g.teams = g.features.teams ? parseTeamsJson(tv(gT, gr, "TeamsJSON")) : null;
    try {
      var act = tv(gT, gr, "AdminActivityJSON");
      g.adminActivity = act ? JSON.parse(String(act)) : [];
    } catch (e2) { g.adminActivity = []; }
    games.push(g);
    gameMap[gid] = g;
    if (g.LocationName) usedLocationsInHistory[g.LocationName.toLowerCase().trim()] = true;
  }

  var signups = tobjs(sT);
  var profiles = tobjs(pT);

  var playersWithHistory = {}, completedYesGamesByPlayer = {}, rsvpGamesByPlayer = {}, verifiedGamesByPlayer = {}, outcomeByPlayer = {}, seenOutcomes = {};
  signups.forEach(function(s) {
    var em = String(s.Email || "").toLowerCase().trim(), g = gameMap[String(s.GameID)];
    if (!em || !g) return;
    if (g.hasStartedOrPassed || s.Status === "Yes" || s.Status === "Maybe" || s.Status === "Waitlist") playersWithHistory[em] = true;
    if (!g.hasStartedOrPassed || g.Status === "Cancelled") return;
    var key = String(s.GameID) + "|" + em;
    if (seenOutcomes[key]) return;
    seenOutcomes[key] = true;
    var attendance = String(s.AttendanceStatus || ""), attended = attendance === "Attended", absent = attendance === "NoShow";
    if (s.Status === "Yes") rsvpGamesByPlayer[em] = (rsvpGamesByPlayer[em] || 0) + 1;
    if (attended) verifiedGamesByPlayer[em] = (verifiedGamesByPlayer[em] || 0) + 1;
    // Keep the historic participation total for existing standings. Explicit
    // attendance takes precedence; unverified RSVP history remains separately visible.
    if (attended || (s.Status === "Yes" && !absent)) completedYesGamesByPlayer[em] = (completedYesGamesByPlayer[em] || 0) + 1;
    var late = !!s.LateDropAt;
    if (!attended && !absent && !late) return;
    var out = outcomeByPlayer[em] || (outcomeByPlayer[em] = { success: 0, sample: 0, noShows: 0, lateDrops: 0 });
    out.sample++;
    // One match, one outcome. Recording Attended corrects a mistaken NoShow and
    // supersedes a previous late drop; NoShow + late drop is still only one failure.
    if (attended) out.success++;
    else if (absent) out.noShows++;
    else if (late) out.lateDrops++;
  });

  var profileMap = {};
  profiles.forEach(function(p) {
    var em = String(p.Email || "").toLowerCase().trim();
    if (!em) return;
    var role = String(p.Role || "Player").trim();
    var accountStatus = String(p.AccountStatus || "Active").trim();
    var position = String(p.Position || "MID").toUpperCase().trim() || "MID";
    var legacyLateDrops = Number(p.LateDrops || 0);
    var recorded = outcomeByPlayer[em] || { success: 0, sample: 0, noShows: 0, lateDrops: 0 };
    var lateDrops = recorded.lateDrops;
    var canBringGuests = !(p.CanBringGuests === false || String(p.CanBringGuests === "" ? "TRUE" : p.CanBringGuests).toLowerCase().trim() === "false");
    var protectedAdmin = isProtectedAdmin(em);
    if (protectedAdmin) { role = "Admin"; accountStatus = "Active"; canBringGuests = true; }
    var playedCount = completedYesGamesByPlayer[em] || 0;
    var totalAttempts = recorded.sample;
    profileMap[em] = {
      email: em,
      fullName: String(p.FullName || em),
      joinedDate: formatDateIso(p.JoinedDate),
      role: role,
      position: position,
      gamesPlayed: playedCount,
      verifiedGamesPlayed: verifiedGamesByPlayer[em] || 0,
      rsvpGamesCount: rsvpGamesByPlayer[em] || 0,
      reliabilitySampleCount: totalAttempts,
      noShows: recorded.noShows,
      legacyLateDrops: legacyLateDrops,
      wins: 0, draws: 0, losses: 0, points: 0,
      lateDrops: lateDrops,
      lateEarlyCount: Number(p.LateEarlyCount || 0),
      awayUntil: formatDateIso(p.AwayUntil) || "",
      notifyPref: normalizeNotifyPref(p.NotifyPref),
      lang: normalizeLang(p.Lang),
      defaultTeam: normalizeTeamPref(p.DefaultTeam),
      reliability: totalAttempts > 0 ? Math.max(0, Math.round((recorded.success / totalAttempts) * 100)) : null,
      motmCount: 0,
      status: accountStatus,
      canBringGuests: canBringGuests,
      isProtected: protectedAdmin,
      isSystemAdminOnly: isNonPlayingSystemAdmin(em),
      hasHistory: !!playersWithHistory[em]
    };
  });

  var activePlayingMemberEmails = Object.keys(profileMap).filter(function(em) {
    var pr = profileMap[em];
    return !pr.isSystemAdminOnly && String(pr.status || "Active").toLowerCase() !== "suspended";
  });

  var formattedInvites = tobjs(iT).map(function(inv) {
    var em = String(inv.Email || "").toLowerCase().trim();
    var isReg = !!profileMap[em];
    return { Email: em, Status: isReg ? "Registered" : (inv.Status || "Pending"), InvitedDate: formatDateIso(inv.InvitedDate), FullName: isReg ? profileMap[em].fullName : "" };
  });

  var locations = tobjs(lT).map(function(l) {
    return {
      LocationID: String(l.LocationID || ""), Name: String(l.Name || ""), Address: String(l.Address || ""),
      Latitude: l.Latitude === "" ? "" : String(l.Latitude), Longitude: l.Longitude === "" ? "" : String(l.Longitude),
      Notes: String(l.Notes || ""), LocationStatus: String(l.LocationStatus || "Active").trim(),
      hasHistory: !!usedLocationsInHistory[String(l.Name || "").toLowerCase().trim()]
    };
  });
  var locationMap = {};
  locations.forEach(function(l) {
    locationMap[l.Name] = { id: l.LocationID, name: l.Name, address: l.Address, lat: l.Latitude, lng: l.Longitude, notes: l.Notes, status: l.LocationStatus, hasHistory: l.hasHistory };
  });

  var blastsByGame = {};
  (bT ? tobjs(bT) : []).forEach(function(b) {
    var gId = String(b.GameID || "");
    if (!blastsByGame[gId]) blastsByGame[gId] = [];
    blastsByGame[gId].push({
      blastId: String(b.BlastID), gameId: gId, senderName: String(b.SenderName || "League Admin"), senderEmail: String(b.SenderEmail || ""),
      message: String(b.Message || ""), timestamp: formatDateTimeStr(b.Timestamp), hidden: String(b.Hidden || "").toLowerCase() === "true" || b.Hidden === true
    });
  });

  var signupsByGameId = {};
  signups.forEach(function(s) {
    var gId = String(s.GameID);
    (signupsByGameId[gId] = signupsByGameId[gId] || []).push(s);
  });

  var rivalryStats = { whiteWins: 0, darkWins: 0, draws: 0, whiteGoals: 0, darkGoals: 0, totalMatchesScored: 0 };

  games.forEach(function(g) {
    var f = g.features;
    var capProf = (f.captains && g.CaptainEmail && profileMap[g.CaptainEmail] && String(profileMap[g.CaptainEmail].status).toLowerCase() !== "suspended" && !profileMap[g.CaptainEmail].isSystemAdminOnly) ? profileMap[g.CaptainEmail] : null;
    g.captainEmail = capProf ? capProf.email : "";
    g.captainName = capProf ? capProf.fullName : "";

    var gameSignups = (signupsByGameId[g.GameID] || []).slice();
    gameSignups.sort(function(a, b) { return tsMs(a.Timestamp) - tsMs(b.Timestamp); });
    var yesSignups = gameSignups.filter(function(s) { return s.Status === "Yes"; });
    var totalGuests = 0;
    yesSignups.forEach(function(s) { totalGuests += parseGuestNamesList(s.PlusOne).length; });
    var yesCount = yesSignups.length + totalGuests;

    var pendingSpots = 0, pendingPlayers = 0;
    yesSignups.forEach(function(s) {
      var answered = Math.max(tsMs(s.Timestamp), tsMs(s.ReconfirmedAt));
      if (g.ScheduleChangeAt > 0 && answered < g.ScheduleChangeAt) {
        pendingSpots += 1 + parseGuestNamesList(s.PlusOne).length;
        pendingPlayers++;
      }
    });
    g.confirmedCount = yesCount - pendingSpots;
    g.pendingReconfirmCount = pendingSpots;
    g.pendingReconfirmPlayers = pendingPlayers;
    g.attendanceRecorded = gameSignups.some(function(s) { return s.AttendanceStatus === "Attended" || s.AttendanceStatus === "NoShow"; });
    g.attendanceRecordedCount = gameSignups.filter(function(s) { return s.AttendanceStatus === "Attended" || s.AttendanceStatus === "NoShow"; }).length;
    g.verifiedAttendeeCount = gameSignups.filter(function(s) { return s.AttendanceStatus === "Attended"; }).length;
    if (g.Status !== "Cancelled") g.Status = (g.confirmedCount >= g.MinThreshold) ? "Confirmed" : "Pending";
    g.yesCount = yesCount;
    g.yesPlayerCount = yesSignups.length;
    g.plusOneCount = totalGuests;
    g.waitlistCount = gameSignups.filter(function(s) { return s.Status === "Waitlist"; }).length;
    g.maybeCount = gameSignups.filter(function(s) { return s.Status === "Maybe"; }).length;
    g.noCount = gameSignups.filter(function(s) { return s.Status === "No"; }).length;

    var motmTally = {}, topMotmEmail = "", topMotmVotes = 0, respondedMap = {};
    var playedSet = {};   // only people who actually played can vote or be voted for
    gameSignups.forEach(function(ys) {
      if (ys.AttendanceStatus === "Attended" || (ys.AttendanceStatus !== "NoShow" && ys.Status === "Yes")) playedSet[String(ys.Email).toLowerCase().trim()] = true;
    });
    g.roster = gameSignups.map(function(s) {
      var em = String(s.Email).toLowerCase().trim();
      if (s.Status === "Yes" || s.Status === "Maybe" || s.Status === "No" || s.Status === "Waitlist") respondedMap[em] = true;
      var prof = profileMap[em];
      var voteTarget = f.motm && (g.hasScore || g.hasEnded) ? String(s.MotmVote || "").toLowerCase().trim() : "";
      if (playedSet[em] && voteTarget && playedSet[voteTarget] && profileMap[voteTarget] && voteTarget !== em) {
        motmTally[voteTarget] = (motmTally[voteTarget] || 0) + 1;
        if (motmTally[voteTarget] > topMotmVotes) { topMotmVotes = motmTally[voteTarget]; topMotmEmail = voteTarget; }
      }
      var guestList = s.Status === "Yes" ? parseGuestNamesList(s.PlusOne) : [];
      var answeredAt = Math.max(tsMs(s.Timestamp), tsMs(s.ReconfirmedAt));
      var needsReconfirm = g.ScheduleChangeAt > 0 && (s.Status === "Yes" || s.Status === "Maybe") && answeredAt < g.ScheduleChangeAt;
      var isYes = s.Status === "Yes";
      var gkFlag = isYes && f.gk && isTruthy(s.PlayGK);
      return {
        email: em,
        fullName: prof ? prof.fullName : em,
        position: prof ? prof.position : "MID",
        status: s.Status,
        timestamp: formatDateTimeStr(s.Timestamp),
        bringBall: isYes && f.ball && isTruthy(s.BringBall),
        bringPinnies: isYes && f.pinnies && isTruthy(s.BringPinnies),
        playGK: gkFlag,
        playPosition: effectivePlayPosition(prof ? prof.position : "MID", gkFlag, f.gk),
        guestNames: guestList,
        guestCount: guestList.length,
        plusOne: guestList.length > 0,
        availabilityNote: (isYes && f.arrival) ? String(s.AvailabilityNote || "") : "",
        needsReconfirm: needsReconfirm,
        attendanceStatus: String(s.AttendanceStatus || ""),
        attendanceRecordedAt: formatDateTimeStr(s.AttendanceRecordedAt),
        lateDropRecorded: !!s.LateDropAt,
        motmVote: voteTarget
      };
    });

    var unrespondedNames = [];
    activePlayingMemberEmails.forEach(function(em) {
      if (!respondedMap[em]) {
        var pr = profileMap[em];
        if (!pr.awayUntil || String(pr.awayUntil) < g.Date) unrespondedNames.push(pr.fullName);
      }
    });
    g.unrespondedCount = unrespondedNames.length;
    g.unrespondedNames = unrespondedNames;

    if (g.teams) {
      var currentSig = computeRosterSignature(g.roster.filter(function(r) { return r.status === "Yes"; }));
      g.teamsStale = !!(g.teams.rosterSignature && g.teams.rosterSignature !== currentSig);
    } else {
      g.teamsStale = false;
    }

    if (g.Status !== "Cancelled" && g.hasScore) {
      var sw = Number(g.ScoreWhite), sd = Number(g.ScoreDark);
      rivalryStats.totalMatchesScored++;
      rivalryStats.whiteGoals += sw;
      rivalryStats.darkGoals += sd;
      if (sw > sd) rivalryStats.whiteWins++; else if (sd > sw) rivalryStats.darkWins++; else rivalryStats.draws++;
      if (g.teams) {
        [["white", sw, sd], ["dark", sd, sw]].forEach(function(side) {
          g.teams[side[0]].forEach(function(tp) {
            var em2 = String(tp.email || "").toLowerCase().trim();
            if (!tp.isGuest && profileMap[em2]) {
              if (side[1] > side[2]) { profileMap[em2].wins++; profileMap[em2].points += 3; }
              else if (side[1] === side[2]) { profileMap[em2].draws++; profileMap[em2].points += 1; }
              else profileMap[em2].losses++;
            }
          });
        });
      }
    }

    g.motmTally = motmTally;
    g.motmWinnerEmail = topMotmEmail;
    g.motmWinnerName = (topMotmEmail && profileMap[topMotmEmail]) ? profileMap[topMotmEmail].fullName : "";
    g.motmWinnerVotes = topMotmVotes;
    if ((g.hasEnded || g.hasScore) && g.Status !== "Cancelled" && topMotmEmail && profileMap[topMotmEmail]) profileMap[topMotmEmail].motmCount++;

    g.blasts = f.announcements ? (blastsByGame[g.GameID] || []) : [];
    g.locationDetails = locationMap[g.LocationName] || { name: g.LocationName, address: g.LocationName, lat: "", lng: "", notes: "", status: "Active" };
  });

  games.sort(function(a, b) {
    var d = String(a.Date).localeCompare(String(b.Date));
    return d !== 0 ? d : parseTimeMinutes(a.Time) - parseTimeMinutes(b.Time);
  });

  var recurring = tobjs(rT).map(function(r) {
    return {
      SeriesID: String(r.SeriesID || ""), Title: String(r.Title || ""), LocationName: String(r.LocationName || ""), DaysOfWeek: String(r.DaysOfWeek || ""),
      Time: formatTimeStr(r.Time), StartDate: formatDateIso(r.StartDate), EndDate: formatDateIso(r.EndDate),
      MaxCap: Number(r.MaxCap || settings.defaultMax), MinThreshold: Number(r.MinThreshold || settings.defaultMin), Status: String(r.Status || "Active"), UpdatedAt: formatDateIso(r.UpdatedAt),
      scheduledChanges: (function() { try { return JSON.parse(String(r.ChangesJSON || "[]")).filter(function(c) { return c.effectiveDate > todayStr; }); } catch (e) { return []; } })()
    };
  });

  return {
    success: true,
    version: APP_VERSION,
    serverTime: formatDateTimeStr(now),
    serverToday: todayStr,
    serverNowMs: now.getTime(),
    leagueTimeZone: leagueTimeZone(),
    matchDurationMinutes: MATCH_DURATION_MINUTES,
    settings: settings,
    requirePinnies: settings.trackPinnies,
    howWePlayRules: howWePlayRules,
    oneSignalAppId: oneSignalAppId,
    pushConfigured: pushConfigured,
    rivalryStats: rivalryStats,
    games: games,
    locations: locations,
    profiles: profileMap,
    invites: formattedInvites,
    recurring: recurring,
    system: { emailQuotaLeft: getEmailQuotaLeft(), pushConfigured: pushConfigured, notifications: notificationSummary() }
  };
}

// ============================================================================
// UNIFIED WAITLIST RECONCILIATION & AUTO-PROMOTION HELPER
// ============================================================================
// Promotes waitlisted players (first come, first served) while there is room — but only
// for matches that haven't kicked off yet. Fixing attendance on a past or live match
// never pulls someone off the waitlist.
function reconcileGameWaitlistAndStatus(gameId) {
  var gr = findGameRow(gameId);
  if (gr === -1) return { promotedEmails: [], statusFlipped: false };
  var gameObj = gameRecord(gr);
  if (gameObj.status === "Cancelled") return { gameObj: gameObj, promotedEmails: [], statusFlipped: false };

  var pT = T("Profiles");
  var isSuspended = {};
  for (var p = 0; p < pT.rows.length; p++) {
    var pEm = String(tv(pT, p, "Email")).toLowerCase().trim();
    if (!isProtectedAdmin(pEm) && String(tv(pT, p, "AccountStatus") || "Active").toLowerCase().trim() === "suspended") isSuspended[pEm] = true;
  }

  var sT = T("Signups");
  var headcount = 0;
  var waitlist = [];
  for (var w = 0; w < sT.rows.length; w++) {
    if (String(tv(sT, w, "GameID")) !== String(gameId)) continue;
    var st = String(tv(sT, w, "Status"));
    var em = String(tv(sT, w, "Email")).toLowerCase().trim();
    if (st === "Yes") headcount += 1 + parseGuestNamesList(tv(sT, w, "PlusOne")).length;
    else if (st === "Waitlist" && !isSuspended[em]) waitlist.push({ row: w, email: em, ts: tsMs(tv(sT, w, "Timestamp")) });
  }
  waitlist.sort(function(a, b) { return a.ts - b.ts; });

  var promotedEmails = [];
  var canPromote = !hasGameStartedOrPassed(gameObj.date, gameObj.time);
  while (canPromote && headcount < gameObj.maxCap && waitlist.length > 0) {
    var next = waitlist.shift();
    if (guestReservationForMember(gameId, next.email)) continue; // never count a legacy guest twice
    twMany(sT, next.row, { Status: "Yes", Timestamp: new Date(), PlusOne: "", BringBall: false, BringPinnies: false, PlayGK: defaultPlayGkFor(next.email, gameId), AvailabilityNote: "" });
    headcount++;
    promotedEmails.push(next.email);
  }

  var committed = gameCommittedCounts(gameId);
  var newGameStatus = (committed.confirmed >= gameObj.minThreshold) ? "Confirmed" : "Pending";
  if (gameObj.oldStatus !== newGameStatus) tw(T("Games"), gr, "Status", newGameStatus);
  if (promotedEmails.length) syncTeamsForGame(gr);

  return {
    gameObj: gameObj,
    recalculatedYes: headcount,
    recalculatedConfirmed: committed.confirmed,
    pendingReconfirmCount: committed.pending,
    newGameStatus: newGameStatus,
    statusFlipped: (gameObj.oldStatus !== newGameStatus) && gameObj.date >= formatDateIso(new Date()),
    justConfirmed: (gameObj.oldStatus !== "Confirmed" && newGameStatus === "Confirmed"),
    promotedEmails: promotedEmails
  };
}

function sendWaitlistPromotionEmails(gameObj, promotedEmails) {
  if (!gameObj || !promotedEmails || !promotedEmails.length) return;
  notifyMail(promotedEmails, mailWaitlistPromoted(gameObj), {
    en: ["🎉 You're off the waitlist!", "A spot opened in " + gameObj.title + " (" + mailWhen(gameObj.date, gameObj.time, "en") + ") — you're in!"],
    es: ["🎉 ¡Saliste de la lista de espera!", "Se liberó un lugar en " + gameObj.title + " (" + mailWhen(gameObj.date, gameObj.time, "es") + "). ¡Ya estás confirmado!"]
  });
}

// ============================================================================
// RSVP ENGINE WITH LOCKSERVICE & GEAR-BRINGER DROPOUT PROTECTION
// ============================================================================
function updateSignup(data) {
  var user = verifyUserCredentials(data.requesterEmail || data.email, data.requesterPin);
  if (!user || user.suspended || user.rateLimited) return fail(tr("errAuth"));
  if (isNonPlayingSystemAdmin(user.email)) return fail(tr("errSysAdminRsvp"));
  var result = processSignupCore(String(data.gameId), user.email, String(data.status));
  if (!result.success) return jsonResponse(result);
  var extra = { finalStatus: result.finalStatus };
  if (result.justConfirmed) extra.justConfirmed = true;
  if (result.captainLeft) extra.captainLeft = true;
  return okData(user, result.message, extra);
}

function processSignupCore(gameId, email, requestedStatus) {
  if (["Yes", "Maybe", "No", "Waitlist"].indexOf(requestedStatus) === -1) return { success: false, error: tr("errBadStatus") };
  var settings = getLeagueSettings();
  var gr = findGameRow(gameId);
  if (gr === -1) return { success: false, error: tr("errGameNotFound") };
  var gameObj = gameRecord(gr);
  var lock = gameLockState(gameObj);
  if (lock.locked) return { success: false, error: lockedError(lock.reason), locked: lock.reason };
  var feat = gameObj.features;

  var sT = T("Signups");
  var myRow = findSignupRow(gameId, email);
  var oldStatus = myRow === -1 ? "" : String(tv(sT, myRow, "Status"));
  if ((requestedStatus === "Yes" || requestedStatus === "Waitlist") && oldStatus !== "Yes") {
    var guestConflict = guestReservationForMember(gameId, email);
    if (guestConflict) return { success: false, error: tr("errMemberAlreadyGuest", { host: guestConflict.hostName }) };
  }
  var hadBall = myRow !== -1 && isTruthy(tv(sT, myRow, "BringBall"));
  var hadPinnies = myRow !== -1 && isTruthy(tv(sT, myRow, "BringPinnies"));
  var hadGK = myRow !== -1 && isTruthy(tv(sT, myRow, "PlayGK"));
  var hadGuests = myRow !== -1 ? parseGuestNamesList(tv(sT, myRow, "PlusOne")).length : 0;

  var othersHeadcount = 0;
  for (var j = 0; j < sT.rows.length; j++) {
    if (j === myRow) continue;
    if (String(tv(sT, j, "GameID")) === String(gameId) && String(tv(sT, j, "Status")) === "Yes") {
      othersHeadcount += 1 + parseGuestNamesList(tv(sT, j, "PlusOne")).length;
    }
  }

  var finalStatus = requestedStatus === "Waitlist" ? "Yes" : requestedStatus;
  var customMsg = "";
  if (finalStatus === "Yes" && oldStatus !== "Yes" && othersHeadcount >= gameObj.maxCap) {
    finalStatus = "Waitlist";
    customMsg = tr("msgWaitlisted", { cap: gameObj.maxCap });
  }
  if (oldStatus === finalStatus && (oldStatus === "Yes" || oldStatus === "Maybe")) {
    // Same answer again. After a time/place change this is the player confirming they still
    // stand by it, so record that without touching their place in line.
    var reMsg = "";
    if (myRow !== -1 && gameObj.scheduleChangeAt > 0 && Math.max(tsMs(tv(sT, myRow, "Timestamp")), tsMs(tv(sT, myRow, "ReconfirmedAt"))) < gameObj.scheduleChangeAt) {
      if (tHas(sT, "ReconfirmedAt")) tw(sT, myRow, "ReconfirmedAt", new Date());
      reMsg = tr(finalStatus === "Yes" ? "msgReconfirmedIn" : "msgReconfirmedMaybe");
    }
    var reRecon = reMsg ? reconcileGameWaitlistAndStatus(gameId) : null;
    if (reRecon && reRecon.statusFlipped) notifyStatusFlip(reRecon.gameObj, reRecon.newGameStatus, reRecon.recalculatedConfirmed);
    return { success: true, finalStatus: finalStatus, message: reMsg, justConfirmed: !!(reRecon && reRecon.justConfirmed) };
  }

  // Late drop: leaving a YES within the window before kickoff (or during the match)
  if (feat.lateDrops && oldStatus === "Yes" && finalStatus !== "Yes" && gameObj.date === formatDateIso(new Date())) {
    var now = new Date();
    var minsToKick = (gameInstantMs(gameObj.date, gameObj.time) - now.getTime()) / 60000;
    if (minsToKick <= LATE_DROP_WINDOW_MINUTES && minsToKick > -60) recordIdempotentLateDrop(gameId, email);
  }

  var fields = { Status: finalStatus };
  if (!(oldStatus === "Waitlist" && finalStatus === "Waitlist")) fields.Timestamp = new Date();
  var startsAsGK = finalStatus === "Yes" && oldStatus !== "Yes" && defaultPlayGkFor(email, gameId);
  if (finalStatus === "Yes" && oldStatus !== "Yes") fields.PlayGK = startsAsGK;
  if (finalStatus !== "Yes") {
    // Leaving the YES list releases everything attached to the spot: guests, gear, timing note.
    fields.BringBall = false; fields.BringPinnies = false; fields.PlayGK = false; fields.PlusOne = ""; fields.AvailabilityNote = "";
  }
  if (myRow !== -1) {
    twMany(sT, myRow, fields);
  } else {
    tadd(sT, { SignupID: "S_" + new Date().getTime() + "_" + Math.floor(Math.random() * 1000), GameID: gameId, Email: email, Status: finalStatus, Timestamp: new Date(),
      BringBall: false, BringPinnies: false, PlayGK: startsAsGK, PlusOne: "", MotmVote: "", AvailabilityNote: "" });
  }

  var recon = reconcileGameWaitlistAndStatus(gameId);
  if (oldStatus === "Yes" || finalStatus === "Yes") syncTeamsForGame(gr);

  var playerName = email;
  var pr = findProfileRow(email);
  if (pr !== -1) playerName = String(tv(T("Profiles"), pr, "FullName") || email);

  // Gear / goalkeeper coverage alert (only for the items the league actually tracks)
  if (oldStatus === "Yes" && finalStatus !== "Yes") {
    var leftGK = hadGK;
    var watch = (hadBall && feat.ball) || (hadPinnies && feat.pinnies) || (leftGK && feat.gk);
    if (watch) {
      var balls = 0, pinnies = 0, gks = 0;
      for (var a = 0; a < sT.rows.length; a++) {
        if (String(tv(sT, a, "GameID")) === String(gameId) && String(tv(sT, a, "Status")) === "Yes") {
          if (isTruthy(tv(sT, a, "BringBall"))) balls++;
          if (isTruthy(tv(sT, a, "BringPinnies"))) pinnies++;
          if (isTruthy(tv(sT, a, "PlayGK"))) gks++;
        }
      }
      var lackBall = feat.ball && hadBall && balls === 0;
      var lackPinnies = feat.pinnies && hadPinnies && pinnies === 0;
      var lackGK = feat.gk && leftGK && gks === 0;
      if (lackBall || lackPinnies || lackGK) {
        var missingEn = [], missingEs = [];
        if (lackBall) { missingEn.push("⚽ ball"); missingEs.push("⚽ balón"); }
        if (lackPinnies) { missingEn.push("🦺 pinnies"); missingEs.push("🦺 chalecos"); }
        if (lackGK) { missingEn.push("🧤 goalkeeper"); missingEs.push("🧤 portero"); }
        appendGameActivityLog(gr, playerName, "⚠️ Gear/GK alert: left the match with no " + missingEn.join(", "));
        notifyMail(getGameOrganizersRecipients(gameId), mailGearAlert({ playerName: playerName, status: finalStatus, g: gameObj, ball: lackBall, pinnies: lackPinnies, gk: lackGK }), {
          en: ["⚠️ Gear / GK alert", playerName + " dropped out of " + gameObj.title + " — now missing: " + missingEn.join(", ")],
          es: ["⚠️ Alerta de equipo / portero", playerName + " se bajó de " + gameObj.title + " — ahora falta: " + missingEs.join(", ")]
        });
      }
    }
  }

  if (recon && recon.promotedEmails && recon.promotedEmails.length > 0) sendWaitlistPromotionEmails(gameObj, recon.promotedEmails);

  if (finalStatus === "Yes" && oldStatus !== "Yes" && settings.msgConfirm) {
    notifyMail([email], mailRsvpConfirmed({ g: gameObj, yes: recon ? recon.recalculatedYes : "", min: gameObj.minThreshold, max: gameObj.maxCap }), {
      en: ["✅ You're in: " + gameObj.title, gameObj.title + " · " + mailWhen(gameObj.date, gameObj.time, "en") + " (" + gameObj.location + ")."],
      es: ["✅ Confirmado: " + gameObj.title, gameObj.title + " · " + mailWhen(gameObj.date, gameObj.time, "es") + " (" + gameObj.location + ")."]
    }, { pushOnly: !EMAIL_RSVP_CONFIRMATIONS });
  }

  if (recon && recon.statusFlipped) notifyStatusFlip(gameObj, recon.newGameStatus, recon.recalculatedYes);

  var promotedMsg = recon && recon.promotedEmails.length ? tr("msgPromoted", { n: recon.promotedEmails.length }) : "";
  var guestsMsg = (oldStatus === "Yes" && hadGuests > 0) ? tr("msgGuestsReleased", { n: hadGuests }) : "";
  return {
    success: true,
    finalStatus: finalStatus,
    justConfirmed: recon ? recon.justConfirmed : false,
    captainLeft: !!(feat.captains && gameObj.captainEmail === email && finalStatus !== "Yes"),
    message: [customMsg, guestsMsg, promotedMsg].filter(Boolean).join(" ")
  };
}

function bumpProfileCounter(email, header) {
  var r = findProfileRow(email);
  if (r === -1) return;
  var t = T("Profiles");
  tw(t, r, header, Number(tv(t, r, header) || 0) + 1);
}

function recordIdempotentLateDrop(gameId, email) {
  var r = findSignupRow(gameId, email);
  if (r === -1) return;
  var sT = T("Signups");
  if (tv(sT, r, "LateDropAt")) return;
  // Old aggregate counts stay untouched and are shown separately as legacy history.
  // A legacy marker identifies the same event: do not create a second new penalty.
  var key = "LATEDROP_" + String(gameId) + "_" + String(email).toLowerCase().trim();
  if (PropertiesService.getScriptProperties().getProperty(key)) return;
  tw(sT, r, "LateDropAt", new Date());
}

function recordIdempotentLateEarlyNoteCount(gameId, email) {
  var key = "LATEEARLY_" + String(gameId) + "_" + String(email).toLowerCase().trim();
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty(key)) return;
  props.setProperty(key, "true");
  bumpProfileCounter(email, "LateEarlyCount");
}

// ============================================================================
// EQUIPMENT, NAMED GUESTS & RUNNING LATE / LEAVING EARLY NOTES
// ============================================================================
function updateSignupExtras(data) {
  var user = verifyUserCredentials(data.requesterEmail, data.requesterPin);
  if (!user || user.suspended || user.rateLimited) return fail(tr("errUnauthorized"));
  var settings = getLeagueSettings();
  var gameId = String(data.gameId || "");
  var field = String(data.field || "");
  var val = data.value;

  var gr = findGameRow(gameId);
  if (gr === -1) return fail(tr("errGameNotFound"));
  var gameObj = gameRecord(gr);
  var lock = gameLockState(gameObj);
  if (lock.locked) return fail(lockedError(lock.reason), { locked: lock.reason });

  var sT = T("Signups");
  var myRow = findSignupRow(gameId, user.email);
  if (myRow === -1 || String(tv(sT, myRow, "Status")) !== "Yes") return fail(tr("errMustBeYes"));

  var msg = "";
  var feat = gameObj.features;
  if (field === "availabilityNote") {
    if (!feat.arrival) return fail(tr("errFeatureOff"));
    var cleanNote = String(val || "").replace(/[<>]/g, "").trim().slice(0, 60);
    tw(sT, myRow, "AvailabilityNote", cleanNote);
    if (cleanNote) {
      recordIdempotentLateEarlyNoteCount(gameId, user.email);
      notifyMail(getGameOrganizersRecipients(gameId), mailTimingAlert({ playerName: user.fullName, playerEmail: user.email, note: cleanNote, g: gameObj }), {
        en: ["🕒 " + user.fullName + ": " + timingWord(cleanNote, "en"), gameObj.title + " · " + mailWhen(gameObj.date, gameObj.time, "en")],
        es: ["🕒 " + user.fullName + ": " + timingWord(cleanNote, "es"), gameObj.title + " · " + mailWhen(gameObj.date, gameObj.time, "es")]
      });
      msg = tr("msgNoteSaved");
    } else {
      msg = tr("msgNoteCleared");
    }
  } else if (field === "guestNames" || field === "plusOne") {
    var previousGuests = parseGuestNamesList(tv(sT, myRow, "PlusOne"));
    var previousNames = {};
    previousGuests.forEach(function(n) { previousNames[cleanPersonName(n).toLowerCase()] = true; });
    var maxG = settings.maxGuests;
    var rawList = Array.isArray(val) ? val : (typeof val === "boolean" ? (val ? ["Guest 1"] : []) : String(val || "").split(","));
    var cleaned = [];
    var seenOwn = {};
    for (var n = 0; n < rawList.length; n++) {
      var gName = cleanPersonName(String(rawList[n] || "").replace(/,/g, " "));
      if (!gName) continue;
      var lower = gName.toLowerCase();
      if (seenOwn[lower]) return fail(tr("errGuestDuplicateOwn", { name: gName }));
      seenOwn[lower] = true;
      cleaned.push(gName);
    }
    var onlyRemoving = cleaned.every(function(n) { return !!previousNames[n.toLowerCase()]; });
    // Removing an existing reservation is always allowed, including after guest
    // privileges are revoked or the league lowers its limit below the old count.
    if (!onlyRemoving) {
      if (!user.canBringGuests && !user.isProtected) return fail(tr("errGuestsDisabled"));
      if (cleaned.length > 0 && maxG === 0) return fail(tr("errGuestsLeagueOff"));
      if (cleaned.length > maxG) return fail(tr("errGuestMax", { max: maxG }));
    }

    var pT = T("Profiles");
    var emailToName = {};
    for (var pr = 0; pr < pT.rows.length; pr++) emailToName[String(tv(pT, pr, "Email")).toLowerCase().trim()] = String(tv(pT, pr, "FullName") || "").trim();
    var otherYes = 0, guestOwner = {}, memberNames = {};
    Object.keys(emailToName).forEach(function(em) {
      var memberName = cleanPersonName(emailToName[em]);
      if (memberName) memberNames[memberName.toLowerCase()] = memberName;
    });
    for (var k = 0; k < sT.rows.length; k++) {
      if (String(tv(sT, k, "GameID")) === gameId && String(tv(sT, k, "Status")) === "Yes") {
        var rowEm = String(tv(sT, k, "Email")).toLowerCase().trim();
        var host = emailToName[rowEm] || rowEm;
        memberNames[host.toLowerCase()] = host;
        if (k !== myRow) {
          var og = parseGuestNamesList(tv(sT, k, "PlusOne"));
          otherYes += 1 + og.length;
          og.forEach(function(x) { guestOwner[x.toLowerCase()] = host; });
        }
      }
    }
    for (var c = 0; c < cleaned.length; c++) {
      var cl = cleaned[c].toLowerCase();
      if (previousNames[cl]) continue; // an old conflict must not prevent releasing another spot
      if (guestOwner[cl]) return fail(tr("errGuestTaken", { name: cleaned[c], host: guestOwner[cl] }));
      if (memberNames[cl]) return fail(tr("errGuestIsMember", { name: cleaned[c] }));
    }
    if (cleaned.length > previousGuests.length && otherYes + 1 + cleaned.length > gameObj.maxCap) {
      return fail(tr("errGuestCapacity", { n: cleaned.length, left: Math.max(0, gameObj.maxCap - (otherYes + 1)), cap: gameObj.maxCap }));
    }
    tw(sT, myRow, "PlusOne", cleaned.join(", "));
    var recon = reconcileGameWaitlistAndStatus(gameId);
    syncTeamsForGame(gr);
    if (recon && recon.promotedEmails && recon.promotedEmails.length > 0) sendWaitlistPromotionEmails(recon.gameObj, recon.promotedEmails);
    if (recon && recon.statusFlipped && recon.gameObj) notifyStatusFlip(recon.gameObj, recon.newGameStatus, recon.recalculatedYes);
    msg = cleaned.length > 0 ? tr("msgGuestsSaved", { n: cleaned.length, names: cleaned.join(", ") }) : tr("msgGuestsCleared");
  } else {
    var colMap = { bringBall: "BringBall", bringPinnies: "BringPinnies", playGK: "PlayGK" };
    var trackMap = { bringBall: feat.ball, bringPinnies: feat.pinnies, playGK: feat.gk };
    if (!colMap[field]) return fail(tr("errBadGear"));
    if (!trackMap[field]) return fail(tr("errGearOff"));
    tw(sT, myRow, colMap[field], !!val);
    if (field === "playGK") syncTeamsForGame(gr);
  }
  return okData(user, msg);
}

// ============================================================================
// CAPTAIN & ADMIN OPERATIONAL ACTIONS
// ============================================================================
function requireManagedGame(data, opts) {
  opts = opts || {};
  var gameId = String(data.gameId || "");
  var manager = verifyGameManagerRequest(data, gameId);
  if (!manager) return { error: tr("errManagerOnly") };
  var gr = findGameRow(gameId);
  if (gr === -1) return { error: tr("errGameNotFound") };
  var g = gameRecord(gr);
  if (!opts.allowCancelled && g.status === "Cancelled") return { error: tr("errCancelled") };
  return { manager: manager, gr: gr, g: g, gameId: gameId };
}

function adminNudgeUnresponded(data) {
  var ctx = requireManagedGame(data);
  if (ctx.error) return fail(ctx.error);
  if (hasGameStartedOrPassed(ctx.g.date, ctx.g.time)) return fail(tr("errAfterKickoff"));
  var props = PropertiesService.getScriptProperties();
  var cooldownKey = "NUDGE_COOLDOWN_" + ctx.gameId;
  var lastNudgeMs = Number(props.getProperty(cooldownKey) || 0);
  var nowMs = new Date().getTime();
  if (lastNudgeMs > 0 && (nowMs - lastNudgeMs) < (30 * 60 * 1000) && !data.force) {
    return fail(tr("errNudgeCooldown", { mins: Math.max(1, Math.ceil(((30 * 60 * 1000) - (nowMs - lastNudgeMs)) / 60000)) }));
  }

  var appData = getAppData();
  var game = appData.games.filter(function(g) { return String(g.GameID) === ctx.gameId; })[0];
  var responded = (game.roster || []).map(function(r) { return r.email; });
  var missing = Object.keys(appData.profiles).filter(function(em) {
    var pr = appData.profiles[em];
    if (isNonPlayingSystemAdmin(em)) return false;
    if (String(pr.status || "Active").toLowerCase() === "suspended") return false;
    if (pr.awayUntil && String(pr.awayUntil) >= game.Date) return false;
    return responded.indexOf(em) === -1;
  });
  if (missing.length === 0) return fail(tr("errNudgeNobody"));

  var showWeather = getLeagueSettings().featWeather;
  var weather = showWeather ? fetchBackendWeather(game.locationDetails.lat, game.locationDetails.lng, game.Date, game.Time) : null;
  missing.forEach(function(playerEmail) {
    var prof = appData.profiles[playerEmail];
    notifyMail([playerEmail], mailNudge({ playerEmail: playerEmail, playerName: prof.fullName, managerName: ctx.manager.fullName, game: game, weather: weather, showWeather: showWeather }), {
      en: ["⚽ Are you playing? " + game.Title, game.Title + " · " + mailWhen(game.Date, game.Time, "en") + ". Open to confirm your answer."],
      es: ["⚽ ¿Juegas? " + game.Title, game.Title + " · " + mailWhen(game.Date, game.Time, "es") + ". Abre para confirmar tu respuesta."]
    }, { pushOnly: true, dedupeKey: "nudge:" + ctx.gameId + ":" + nowMs });
  });
  props.setProperty(cooldownKey, String(nowMs));
  appendGameActivityLog(ctx.gr, ctx.manager.fullName, "Queued reminder for " + missing.length + " unresponded player(s)");
  return okData(ctx.manager, tr("msgNudged", { n: missing.length }));
}

function captainStepDown(data) {
  var user = verifyUserCredentials(data.requesterEmail, data.requesterPin);
  if (!user || user.suspended || user.rateLimited) return fail(tr("errUnauthorized"));
  var gr = findGameRow(data.gameId);
  if (gr === -1) return fail(tr("errGameNotFound"));
  var gT = T("Games");
  var curCap = String(tv(gT, gr, "CaptainEmail") || "").toLowerCase().trim();
  if (!(curCap && curCap === user.email) && !user.isAdmin) return fail(tr("errNotCaptain"));
  var gameTitle = String(tv(gT, gr, "Title"));
  tw(gT, gr, "CaptainEmail", "");
  appendGameActivityLog(gr, user.fullName, "Captain stepped down (Reverted to Managed by System Admins)");
  notifyMail(getAllActiveAdminEmails(), mailCaptainStepDown({ playerName: user.fullName, title: gameTitle }), {
    en: ["👑 Captain change: " + gameTitle, user.fullName + " stepped down as captain. The admins run this match now."],
    es: ["👑 Cambio de capitán: " + gameTitle, user.fullName + " dejó la capitanía. Ahora lo llevan los admins."]
  });
  return okData(user, tr("msgSteppedDown"));
}

// Small-sided: with the switch on, an organizer can start a match that's short of its minimum,
// as long as at least 10 are in (5 v 5). Off by default: matches go ahead at their minimum.
function adminApproveSmallSided(data) {
  var ctx = requireManagedGame(data);
  if (ctx.error) return fail(ctx.error);
  if (!getLeagueSettings().featSmallSided) return fail(tr("errFeatureOff"));
  var lock = gameLockState(ctx.g);
  if (lock.locked) return fail(lockedError(lock.reason));
  var formatLabel = String(data.matchFormat || "Small-sided (5v5+)").trim().slice(0, 40);
  var currentYes = gameCommittedCounts(ctx.gameId).confirmed;
  if (currentYes < SMALL_SIDED_MIN_PLAYERS) return fail(tr("errSmallSidedMin", { n: SMALL_SIDED_MIN_PLAYERS }));
  var newThreshold = Math.max(SMALL_SIDED_MIN_PLAYERS, currentYes);
  twMany(T("Games"), ctx.gr, { MinThreshold: newThreshold, Status: "Confirmed", MatchFormat: formatLabel });
  appendGameActivityLog(ctx.gr, ctx.manager.fullName, "Approved " + formatLabel + " (Min threshold set to " + newThreshold + ")");
  var gObj = gameRecord(ctx.gr);
  notifyStatusFlip(gObj, "Confirmed", currentYes);
  return okData(ctx.manager, tr("msgSmallSided", { format: formatLabel, n: currentYes }));
}

function adminUpdateAttendance(data) {
  var ctx = requireManagedGame(data);
  if (ctx.error) return fail(ctx.error);
  if (!ctx.g.features.attendance) return fail(tr("errFeatureOff"));
  if (ctx.g.date > formatDateIso(new Date())) return fail(tr("errAttendanceFuture"));
  var targetEmail = resolvePlayerRef(data.targetEmail), newStatus = String(data.newStatus || "Yes");
  if (!targetEmail || ["Yes", "No"].indexOf(newStatus) === -1) return fail(tr("errAttendanceArgs"));
  if (isNonPlayingSystemAdmin(targetEmail)) return fail(tr("errSysAdminRsvp"));
  var pr = findProfileRow(targetEmail);
  if (pr === -1) return fail(tr("errNoProfile"));
  if (newStatus === "No" && !hasGameStartedOrPassed(ctx.g.date, ctx.g.time)) return fail(tr("errNoShowBeforeKickoff"));
  if (newStatus === "Yes") {
    var conflict = guestReservationForMember(ctx.gameId, targetEmail);
    if (conflict) return fail(tr("errMemberAlreadyGuest", { host: conflict.hostName }));
  }
  var playerName = String(tv(T("Profiles"), pr, "FullName") || targetEmail);
  var sT = T("Signups"), row = findSignupRow(ctx.gameId, targetEmail);
  var oldAttendance = row === -1 ? "" : String(tv(sT, row, "AttendanceStatus"));
  var wasYes = row !== -1 && String(tv(sT, row, "Status")) === "Yes";
  var attendance = newStatus === "Yes" ? "Attended" : "NoShow";
  if (oldAttendance === attendance && ((newStatus === "Yes" && wasYes) || (newStatus === "No" && !wasYes)) && !data.clearPenalty) {
    return okData(ctx.manager, tr(newStatus === "Yes" ? "msgAttended" : "msgNoShow", { name: playerName }));
  }
  if (newStatus === "No" && !wasYes && oldAttendance !== "NoShow") return fail(tr("errNoShowNotConfirmed"));
  var now = new Date();
  var fields = { Status: newStatus, AttendanceStatus: attendance, AttendanceRecordedAt: now, AttendanceRecordedBy: ctx.manager.email || "" };
  if (!(wasYes && newStatus === "Yes")) fields.Timestamp = now;
  if (newStatus === "Yes") {
    fields.ReconfirmedAt = now;
    if (!wasYes) fields.PlayGK = defaultPlayGkFor(targetEmail, ctx.gameId);
    if (data.clearPenalty === true) fields.LateDropAt = "";
  } else {
    fields.PlusOne = ""; fields.BringBall = false; fields.BringPinnies = false; fields.PlayGK = false; fields.AvailabilityNote = "";
  }
  if (row !== -1) twMany(sT, row, fields);
  else {
    var fresh = { SignupID: "S_" + now.getTime() + "_" + Math.floor(Math.random() * 1000), GameID: ctx.gameId, Email: targetEmail, Timestamp: now, BringBall: false, BringPinnies: false, PlayGK: false, PlusOne: "", MotmVote: "", AvailabilityNote: "" };
    Object.keys(fields).forEach(function(k) { fresh[k] = fields[k]; });
    tadd(sT, fresh);
  }
  appendGameActivityLog(ctx.gr, ctx.manager.fullName, "Attendance: Marked " + playerName + " as " + (newStatus === "Yes" ? "Attended" : "No-Show"));
  var recon = reconcileGameWaitlistAndStatus(ctx.gameId);
  syncTeamsForGame(ctx.gr);
  if (recon && recon.promotedEmails && recon.promotedEmails.length) sendWaitlistPromotionEmails(recon.gameObj, recon.promotedEmails);
  if (newStatus === "No" && oldAttendance !== "NoShow") {
    notifyMail([targetEmail], null, {
      en: ["📋 Attendance: " + ctx.g.title, "You were marked absent for " + ctx.g.title + " (" + mailDate(ctx.g.date, "en") + "). Ask an organizer if this needs correcting."],
      es: ["📋 Asistencia: " + ctx.g.title, "Te marcaron ausente en " + ctx.g.title + " (" + mailDate(ctx.g.date, "es") + "). Pide al organizador que lo corrija si hace falta."]
    }, { pushOnly: true, gameId: ctx.gameId, expectedSignupStatus: "No" });
  }
  return okData(ctx.manager, tr(newStatus === "Yes" ? "msgAttended" : "msgNoShow", { name: playerName }));
}

function voteMotm(data) {
  var user = verifyUserCredentials(data.requesterEmail, data.requesterPin);
  if (!user || user.suspended || user.rateLimited) return fail(tr("errUnauthorized"));
  var gameId = String(data.gameId || "");
  var nomineeEmail = resolvePlayerRef(data.nomineeEmail);
  if (!gameId || !nomineeEmail) return fail(tr("errMotmPick"));
  if (nomineeEmail === user.email) return fail(tr("errMotmSelf"));
  var gr = findGameRow(gameId);
  if (gr === -1) return fail(tr("errGameNotFound"));
  var g = gameRecord(gr);
  if (g.status === "Cancelled") return fail(tr("errCancelled"));
  if (!g.features.motm) return fail(tr("errFeatureOff"));
  var scheduledEnd = gameInstantMs(g.date, g.time) + MATCH_DURATION_MINUTES * 60000;
  if (!gameHasScore(g) && new Date().getTime() < scheduledEnd) return fail(tr("errMotmLocked"));
  var sT = T("Signups");
  var eligible = function(r) {
    if (r === -1) return false;
    var attendance = String(tv(sT, r, "AttendanceStatus"));
    return attendance === "Attended" || (attendance !== "NoShow" && String(tv(sT, r, "Status")) === "Yes");
  };
  var voterRow = findSignupRow(gameId, user.email);
  if (!eligible(voterRow)) return fail(tr("errMotmVoter"));
  var nomRow = findSignupRow(gameId, nomineeEmail);
  if (!eligible(nomRow)) return fail(tr("errMotmNominee"));
  tw(sT, voterRow, "MotmVote", nomineeEmail);
  return okData(user, tr("msgMotmSaved"));
}

function adminGenerateTeams(data) {
  var ctx = requireManagedGame(data);
  if (ctx.error) return fail(ctx.error);
  if (!ctx.g.features.teams) return fail(tr("errFeatureOff"));
  var strategy = String(data.strategy || "smart") === "defaults" ? "defaults" : "smart";
  var entities = lineupEntitiesForGame(ctx.gameId);
  if (entities.length < 2) return fail(tr("errTeamsMin"));
  var st = quickStandings();
  var res = strategy === "defaults" ? balanceDefaults(entities, st) : balanceSmart(entities, st);

  var teamsObj = {
    white: res.white,
    dark: res.dark,
    strategy: strategy,
    rosterSignature: signatureFromEntities(entities),
    generatedAt: formatDateTimeStr(new Date())
  };
  tw(T("Games"), ctx.gr, "TeamsJSON", JSON.stringify(teamsObj));
  appendGameActivityLog(ctx.gr, ctx.manager.fullName, (strategy === "defaults" ? "Set lineups from default teams" : "Smart-balanced lineups") + " (" + (res.white.length + res.dark.length) + " players)");
  return okData(ctx.manager, tr(strategy === "defaults" ? "msgTeamsDefaults" : "msgTeamsSmart", { a: teamName("white"), b: teamName("dark"), na: res.white.length, nb: res.dark.length }));
}

function adminMoveTeamPlayer(data) {
  var ctx = requireManagedGame(data);
  if (ctx.error) return fail(ctx.error);
  if (!ctx.g.features.teams) return fail(tr("errFeatureOff"));
  var gT = T("Games");
  var teams = parseTeamsJson(tv(gT, ctx.gr, "TeamsJSON"));
  if (!teams) return fail(tr("errNoTeams"));
  var key = resolvePlayerRef(data.key);
  var toSide = data.toSide === "dark" ? "dark" : "white";
  var fromSide = toSide === "white" ? "dark" : "white";
  var idx = teams[fromSide].findIndex(function(p) { return String(p.key).toLowerCase() === key; });
  if (idx === -1) return fail(tr("errNotInLineup"));
  var p = teams[fromSide].splice(idx, 1)[0];
  teams[toSide].push(p);
  teams.strategy = "manual";
  teams.generatedAt = formatDateTimeStr(new Date());
  teams.rosterSignature = signatureFromEntities(lineupEntitiesForGame(ctx.gameId)); // organizer has reviewed this lineup
  tw(gT, ctx.gr, "TeamsJSON", JSON.stringify(teams));
  appendGameActivityLog(ctx.gr, ctx.manager.fullName, "Moved " + p.fullName + " to " + teamName(toSide, true));
  return okData(ctx.manager, tr("msgMoved", { name: p.fullName, team: teamName(toSide) }));
}

// Kept for older cached pages: swap one player from each side by index.
function adminSwapTeamPlayers(data) {
  var ctx = requireManagedGame(data);
  if (ctx.error) return fail(ctx.error);
  if (!ctx.g.features.teams) return fail(tr("errFeatureOff"));
  var gT = T("Games");
  var teams = parseTeamsJson(tv(gT, ctx.gr, "TeamsJSON"));
  if (!teams) return fail(tr("errNoTeams"));
  var wi = Number(data.whiteIdx), di = Number(data.darkIdx);
  if (isNaN(wi) || wi < 0 || wi >= teams.white.length || isNaN(di) || di < 0 || di >= teams.dark.length) return fail(tr("errSwapPick"));
  var a = teams.white[wi], b = teams.dark[di];
  teams.white[wi] = b; teams.dark[di] = a;
  teams.strategy = "manual";
  teams.generatedAt = formatDateTimeStr(new Date());
  teams.rosterSignature = signatureFromEntities(lineupEntitiesForGame(ctx.gameId));
  tw(gT, ctx.gr, "TeamsJSON", JSON.stringify(teams));
  appendGameActivityLog(ctx.gr, ctx.manager.fullName, "Swapped " + a.fullName + " ↔ " + b.fullName);
  return okData(ctx.manager, tr("msgSwapped", { a: a.fullName, b: b.fullName }));
}

function adminSaveScore(data) {
  var ctx = requireManagedGame(data);
  if (ctx.error) return fail(ctx.error);
  if (!ctx.g.features.scores) return fail(tr("errFeatureOff"));
  if (!hasGameStartedOrPassed(ctx.g.date, ctx.g.time)) return fail(tr("errScoreBeforeKickoff"));
  var gT = T("Games");
  if (data.clear) {
    twMany(gT, ctx.gr, { ScoreWhite: "", ScoreDark: "" });
    appendGameActivityLog(ctx.gr, ctx.manager.fullName, "Cleared final score");
    return okData(ctx.manager, tr("msgScoreCleared"));
  }
  var sw = Number(data.scoreWhite), sd = Number(data.scoreDark);
  if (!isFinite(sw) || !isFinite(sd) || sw < 0 || sd < 0 || sw > 99 || sd > 99 || Math.floor(sw) !== sw || Math.floor(sd) !== sd) return fail(tr("errScoreValue"));
  twMany(gT, ctx.gr, { ScoreWhite: sw, ScoreDark: sd });
  appendGameActivityLog(ctx.gr, ctx.manager.fullName, "Recorded Final Score: " + teamName("white", true) + " " + sw + " - " + sd + " " + teamName("dark", true));
  return okData(ctx.manager, tr("msgScoreSaved", { a: teamName("white"), sa: sw, sb: sd, b: teamName("dark") }));
}

// Status-change notices are de-duplicated per match: players get one "Game On" email,
// organizers get "dropped below minimum" alerts (at most every 2 hours). Before this,
// one indecisive player tapping in/out could email the whole roster on every tap.
function notifyStatusFlip(gameObj, newGameStatus, yesCount) {
  var gr = findGameRow(gameObj.id);
  if (gr === -1) return;
  var g = gameRecord(gr);
  if (g.status === "Cancelled" || hasGameStartedOrPassed(g.date, g.time)) return;
  var gT = T("Games");
  var state = {};
  try { state = JSON.parse(String(tv(gT, gr, "NotifyStateJSON") || "{}")) || {}; } catch (e) { state = {}; }
  var nowMs = new Date().getTime();

  var appData = getAppData();
  var game = appData.games.filter(function(x) { return String(x.GameID) === String(gameObj.id); })[0];
  if (!game) return;
  var isConfirmed = (newGameStatus === "Confirmed");
  yesCount = game.confirmedCount === undefined ? yesCount : game.confirmedCount;
  var recipients;
  if (isConfirmed) {
    if (state.confirmedNotified) return;
    if (!getLeagueSettings().msgGameOn) return;
    var rsvpList = game.roster.filter(function(r) { return r.status === "Yes" || r.status === "Maybe"; }).map(function(r) { return r.email; });
    recipients = rsvpList.concat(getGameOrganizersRecipients(game.GameID));
    state.confirmedNotified = true;
  } else {
    if (state.droppedOrgAt && nowMs - state.droppedOrgAt < 2 * 3600000) return;
    recipients = getGameOrganizersRecipients(game.GameID);
    state.droppedOrgAt = nowMs;
  }

  var whenEn = mailWhen(game.Date, game.Time, "en"), whenEs = mailWhen(game.Date, game.Time, "es");
  var queued = notifyMail(recipients, mailStatusFlip({ game: game, confirmed: isConfirmed, yesCount: yesCount }), {
    en: [isConfirmed ? "🟢 Game on: " + game.Title : "⚠️ " + game.Title + " needs players",
         game.Title + " · " + whenEn + " — " + yesCount + "/" + game.MaxCap + " in" + (isConfirmed ? "." : " (minimum " + game.MinThreshold + ").")],
    es: [isConfirmed ? "🟢 ¡Hay partido! " + game.Title : "⚠️ Faltan jugadores: " + game.Title,
         game.Title + " · " + whenEs + " — " + yesCount + "/" + game.MaxCap + " confirmados" + (isConfirmed ? "." : " (mínimo " + game.MinThreshold + ").")]
  }, { pushOnly: isConfirmed });
  if (queued.queued && tHas(gT, "NotifyStateJSON")) tw(gT, gr, "NotifyStateJSON", JSON.stringify(state));
}

// ============================================================================
// ADMIN & CAPTAIN BLASTS, SETTINGS, RECURRING, GAMES & PLAYERS
// ============================================================================
function adminToggleGuestPrivilege(data) {
  var admin = requireAdmin(data);
  if (!admin) return fail(tr("errAdminOnly"));
  var targetEmail = String(data.targetEmail || "").toLowerCase().trim();
  if (isProtectedAdmin(targetEmail)) return fail(tr("errProtectedGuests"));
  var r = findProfileRow(targetEmail);
  if (r === -1) return fail(tr("errNoProfile"));
  var canBring = !!data.canBring;
  tw(T("Profiles"), r, "CanBringGuests", canBring);
  return okData(admin, tr(canBring ? "msgGuestsOn" : "msgGuestsOff", { email: targetEmail }));
}

function adminSetDefaultTeam(data) {
  var admin = requireAdmin(data);
  if (!admin) return fail(tr("errAdminOnly"));
  var targetEmail = String(data.targetEmail || "").toLowerCase().trim();
  var r = findProfileRow(targetEmail);
  if (r === -1) return fail(tr("errNoProfile"));
  var team = normalizeTeamPref(data.team);
  var t = T("Profiles");
  tw(t, r, "DefaultTeam", team);
  var name = String(tv(t, r, "FullName") || targetEmail);
  return okData(admin, team ? tr("msgDefaultTeamSet", { name: name, team: teamName(team === "A" ? "white" : "dark") }) : tr("msgDefaultTeamCleared", { name: name }));
}

function adminSaveLeagueSettings(data) {
  var admin = requireAdmin(data);
  if (!admin) return fail(tr("errAdminOnly"));
  var cur = getLeagueSettings();
  var next = {};
  if (data.teamAName !== undefined) next.teamAName = cleanPersonName(data.teamAName).slice(0, 24);
  if (data.teamBName !== undefined) next.teamBName = cleanPersonName(data.teamBName).slice(0, 24);
  if (next.teamAName === "") return fail(tr("errTeamNameEmpty"));
  if (next.teamBName === "") return fail(tr("errTeamNameEmpty"));
  var a = next.teamAName !== undefined ? next.teamAName : cur.teamAName;
  var b = next.teamBName !== undefined ? next.teamBName : cur.teamBName;
  if (a.toLowerCase() === b.toLowerCase()) return fail(tr("errTeamNamesSame"));
  if (data.teamAShirt !== undefined) { if (SHIRT_COLORS.indexOf(data.teamAShirt) === -1) return fail(tr("errShirt")); next.teamAShirt = data.teamAShirt; }
  if (data.teamBShirt !== undefined) { if (SHIRT_COLORS.indexOf(data.teamBShirt) === -1) return fail(tr("errShirt")); next.teamBShirt = data.teamBShirt; }
  var mn = data.defaultMin !== undefined ? Number(data.defaultMin) : cur.defaultMin;
  var mx = data.defaultMax !== undefined ? Number(data.defaultMax) : cur.defaultMax;
  if (!isFinite(mn) || !isFinite(mx) || mn < 2 || mx > 60 || mn > mx || Math.floor(mn) !== mn || Math.floor(mx) !== mx) return fail(tr("errMinMax"));
  next.defaultMin = mn; next.defaultMax = mx;
  BOOLEAN_SETTINGS.forEach(function(k) { if (data[k] !== undefined) next[k] = data[k] === true || String(data[k]).toLowerCase() === "true"; });
  if (data.maxGuests !== undefined) {
    var mg = Number(data.maxGuests);
    if (!isFinite(mg) || mg < 0 || mg > MAX_GUESTS_PER_PLAYER || Math.floor(mg) !== mg) return fail(tr("errMaxGuests", { max: MAX_GUESTS_PER_PLAYER }));
    next.maxGuests = mg;
  }
  // A switch that needs another is turned off with it (and stays off until turned on again).
  var teamsOn = next.featTeams !== undefined ? next.featTeams : cur.featTeams;
  var scoresOn = next.featScores !== undefined ? next.featScores : cur.featScores;
  if (!teamsOn) { next.featScores = false; scoresOn = false; }
  if (!scoresOn) next.featMotm = false;
  // Matches that already kicked off keep the switches they started with.
  var before = JSON.stringify(featuresFromSettings(cur));
  var preview = {};
  for (var k in cur) preview[k] = cur[k];
  for (var k2 in next) preview[k2] = next[k2];
  if (JSON.stringify(featuresFromSettings(preview)) !== before) freezeStartedMatches(featuresFromSettings(cur));
  saveLeagueSettings(next);
  return okData(admin, tr("msgSettingsSaved"));
}

// Older pages call these two; both now feed the single League Settings store.
function adminTogglePinniesSetting(data) {
  var admin = requireAdmin(data);
  if (!admin) return fail(tr("errAdminOnly"));
  var cur = getLeagueSettings();
  if (!!cur.trackPinnies !== !!data.requirePinnies) freezeStartedMatches(featuresFromSettings(cur));
  saveLeagueSettings({ trackPinnies: !!data.requirePinnies });
  return okData(admin, tr("msgSettingsSaved"));
}

function adminRenameTeams(data) {
  var gameId = String(data.gameId || "");
  var user = gameId ? verifyGameManagerRequest(data, gameId) : requireAdmin(data);
  if (!user || !user.isAdmin) return fail(tr("errAdminOnly"));
  var a = cleanPersonName(data.teamAName || data.whiteName || data.teamWhiteName || "").slice(0, 24);
  var b = cleanPersonName(data.teamBName || data.darkName || data.teamDarkName || "").slice(0, 24);
  if (!a || !b) return fail(tr("errTeamNameEmpty"));
  if (a.toLowerCase() === b.toLowerCase()) return fail(tr("errTeamNamesSame"));
  saveLeagueSettings({ teamAName: a, teamBName: b });
  return okData(user, tr("msgSettingsSaved"));
}

function adminSaveHowWePlay(data) {
  var admin = requireAdmin(data);
  if (!admin) return fail(tr("errAdminOnly"));
  // Empty (or the untouched standard text) means "use the standard rules in each player's language".
  var rulesText = String(data.howWePlayRules || "").trim().slice(0, 2000);
  var props = PropertiesService.getScriptProperties();
  if (!rulesText || rulesText === DEFAULT_HOW_WE_PLAY_RULES.trim()) props.deleteProperty("HOW_WE_PLAY_RULES");
  else props.setProperty("HOW_WE_PLAY_RULES", rulesText);
  return okData(admin, tr("msgRulesSaved"));
}

function adminSendGameBlast(data) {
  var ctx = requireManagedGame(data, { allowCancelled: true });
  if (ctx.error) return fail(ctx.error);
  if (!ctx.g.features.announcements) return fail(tr("errFeatureOff"));
  var alsoEmail = data.alsoEmail === true || String(data.alsoEmail).toLowerCase() === "true";
  var message = String(data.message || "").trim().slice(0, 1000);
  if (!message) return fail(tr("errBlastEmpty"));
  var senderName = String(ctx.manager.fullName || "Match Organizer").trim();
  tadd(T("Blasts"), { BlastID: "BL_" + new Date().getTime(), GameID: ctx.gameId, SenderName: senderName, SenderEmail: ctx.manager.email, Message: message, Timestamp: new Date(), Hidden: false });
  appendGameActivityLog(ctx.gr, senderName, "Sent Match Blast: " + message.slice(0, 45));

  var sT = T("Signups");
  var recipients = [];
  for (var r = 0; r < sT.rows.length; r++) {
    if (String(tv(sT, r, "GameID")) !== ctx.gameId) continue;
    var st = String(tv(sT, r, "Status"));
    if (st === "Yes" || st === "Maybe" || st === "Waitlist") recipients.push(String(tv(sT, r, "Email")).toLowerCase().trim());
  }
  var g = ctx.g;
  if (recipients.length > 0) {
    notifyMail(recipients, mailBlast({ senderName: senderName, message: message, g: g }), {
      en: ["📢 " + g.title, senderName + ": " + message.slice(0, 120)],
      es: ["📢 Aviso: " + g.title, senderName + ": " + message.slice(0, 120)]
    }, { pushOnly: !alsoEmail });
  }
  return okData(ctx.manager, tr(alsoEmail ? "msgBlastSentEmail" : "msgBlastSent", { n: recipients.length }));
}

// An admin writes to one player. It goes by push and email, following that player's
// notification setting, and replies to the email go to the admin who wrote it.
function adminMessagePlayer(data) {
  var admin = requireAdmin(data);
  if (!admin) return fail(tr("errAdminOnly"));
  var target = resolvePlayerRef(data.targetEmail);
  if (!target || target === admin.email || isNonPlayingSystemAdmin(target)) return fail(tr("errDirectTarget"));
  var pr = findProfileRow(target);
  if (pr === -1) return fail(tr("errNoProfile"));
  var message = String(data.message || "").trim().slice(0, 1000);
  if (!message) return fail(tr("errBlastEmpty"));
  var playerName = String(tv(T("Profiles"), pr, "FullName") || target);
  var senderName = String(admin.fullName || "").trim() || MAIL_SENDER_NAME;
  notifyMail([target], mailDirectMessage({ senderName: senderName, message: message }), {
    en: ["💬 " + senderName, message.slice(0, 160)],
    es: ["💬 " + senderName, message.slice(0, 160)]
  }, { replyTo: admin.email });
  return okData(admin, tr("msgDirectSent", { name: playerName }));
}

function adminToggleBlastVisibility(data) {
  var admin = requireAdmin(data);
  if (!admin) return fail(tr("errAdminOnly"));
  var t = T("Blasts");
  var id = String(data.blastId || "");
  var r = tfind(t, function(i) { return String(tv(t, i, "BlastID")) === id; });
  if (r === -1) return fail(tr("errNotFound"));
  var hideFlag = !!data.hidden;
  tw(t, r, "Hidden", hideFlag);
  return okData(admin, tr(hideFlag ? "msgBlastHidden" : "msgBlastShown"));
}

// Effective-dated configurations preserve a future edit until its Apply from date,
// even when that date has not yet entered the rolling four-week match window.
function seriesConfigFromRow(t, r) {
  return {
    Title: String(tv(t, r, "Title")), LocationName: String(tv(t, r, "LocationName")),
    DaysOfWeek: String(tv(t, r, "DaysOfWeek")).replace(/^'/, ""), Time: formatTimeStr(tv(t, r, "Time")),
    StartDate: formatDateIso(tv(t, r, "StartDate")), EndDate: formatDateIso(tv(t, r, "EndDate")),
    MaxCap: Number(tv(t, r, "MaxCap") || getLeagueSettings().defaultMax),
    MinThreshold: Number(tv(t, r, "MinThreshold") || getLeagueSettings().defaultMin)
  };
}
function seriesTimeline(t, r) {
  var base = seriesConfigFromRow(t, r), changes = [];
  try { changes = JSON.parse(String(tv(t, r, "ChangesJSON") || "[]")); } catch (e) { changes = []; }
  if (!Array.isArray(changes)) changes = [];
  changes = changes.filter(function(c) { return c && c.config && /^\d{4}-\d{2}-\d{2}$/.test(String(c.effectiveDate)); });
  if (!changes.length) changes = [{ effectiveDate: "0001-01-01", config: base }];
  changes.sort(function(a, b) { return a.effectiveDate.localeCompare(b.effectiveDate); });
  return changes;
}
function seriesConfigAt(t, r, date) {
  var versions = seriesTimeline(t, r), result = versions[0].config;
  for (var i = 0; i < versions.length; i++) if (versions[i].effectiveDate <= date) result = versions[i].config;
  return result;
}
function seriesWantsDate(config, date) {
  if (!config || date < config.StartDate || date > config.EndDate) return false;
  var p = date.split("-"), day = new Date(Date.UTC(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 12)).getUTCDay();
  return String(config.DaysOfWeek || "").split(",").indexOf(String(day)) !== -1;
}
function seriesOccurrences(t, r, from, through) {
  var out = [], date = from, safety = 800;
  while (date <= through && safety-- > 0) {
    var config = seriesConfigAt(t, r, date);
    if (seriesWantsDate(config, date)) out.push({ date: date, config: config });
    date = isoPlusDays(date, 1);
  }
  return out;
}

function computeRecurringDates(startStr, endStr, daysCsv) {
  var allowed = String(daysCsv || "").split(",").map(function(d) { return String(parseInt(d, 10)); });
  var date = startStr, out = [], safety = 800;
  while (date <= endStr && safety-- > 0) {
    var p = date.split("-"), day = new Date(Date.UTC(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 12)).getUTCDay();
    if (allowed.indexOf(String(day)) !== -1) out.push(date);
    date = isoPlusDays(date, 1);
  }
  return out;
}

function readMinMax(data, fallback) {
  var s = getLeagueSettings();
  var fbMin = fallback ? fallback.min : s.defaultMin, fbMax = fallback ? fallback.max : s.defaultMax;
  var mn = (data.minThreshold === undefined || data.minThreshold === "" || data.minThreshold === null) ? fbMin : Number(data.minThreshold);
  var mx = (data.maxCap === undefined || data.maxCap === "" || data.maxCap === null) ? fbMax : Number(data.maxCap);
  if (!isFinite(mn) || !isFinite(mx) || Math.floor(mn) !== mn || Math.floor(mx) !== mx || mn < 2 || mx > 60 || mn > mx) return null;
  return { min: mn, max: mx };
}

function activeLocationExists(name) {
  var t = T("Locations");
  return tfind(t, function(r) { return String(tv(t, r, "Name")) === name; }) !== -1;
}

function adminSaveRecurringSeries(data) {
  var admin = requireAdmin(data);
  if (!admin) return fail(tr("errAdminOnly"));
  var rT = T("Recurring"), gT = T("Games"), sT = T("Signups");
  var today = formatDateIso(new Date()), isEdit = !!data.seriesId;
  var seriesId = String(data.seriesId || ("SER_" + new Date().getTime()));
  var sr = tfind(rT, function(r) { return String(tv(rT, r, "SeriesID")) === seriesId; });
  if (isEdit && sr === -1) return fail(tr("errNotFound"));
  var title = cleanTitle(data.title || "Lunch soccer"), location = String(data.locationName || "").trim();
  var days = String(data.daysOfWeek || "").split(",").map(function(d) { return d.trim(); }).filter(function(d, i, all) { return /^[0-6]$/.test(d) && all.indexOf(d) === i; }).join(",");
  var time = normalizeTimeInput(data.time), start = normalizeDateInput(data.startDate), end = normalizeDateInput(data.endDate), mm = readMinMax(data);
  var effective = normalizeDateInput(data.effectiveDate) || today;
  if (!start || !end || end < start) return fail(tr("errSeriesDates"));
  if (isEdit && effective < today) return fail(tr("errSeriesEffectivePast"));
  if (!days) return fail(tr("errSeriesDays"));
  if (!time) return fail(tr("errTime"));
  if (!mm) return fail(tr("errMinMax"));
  if (!location || !activeLocationExists(location)) return fail(tr("errLocation"));
  var config = { Title: title, LocationName: location, DaysOfWeek: days, Time: time, StartDate: start, EndDate: end, MaxCap: mm.max, MinThreshold: mm.min };
  var changes = sr === -1 ? [] : seriesTimeline(rT, sr).filter(function(c) { return c.effectiveDate < effective; });
  if (sr === -1) changes = [{ effectiveDate: "0001-01-01", config: config }];
  else changes.push({ effectiveDate: effective, config: config });
  var previousThrough = sr === -1 ? "" : formatDateIso(tv(rT, sr, "GeneratedThrough"));
  var row = { Title: title, LocationName: location, DaysOfWeek: "'" + days, Time: "'" + time, StartDate: "'" + start, EndDate: "'" + end, MaxCap: mm.max, MinThreshold: mm.min, Status: "Active", UpdatedAt: "'" + today, ChangesJSON: JSON.stringify(changes) };
  if (sr === -1) { row.SeriesID = seriesId; sr = tadd(rT, row); }
  else twMany(rT, sr, row);

  var notified = {}, cancelledGames = [], existingDates = {}, created = 0, updated = 0, cancelled = 0, nowMs = new Date().getTime();
  for (var i = 0; i < gT.rows.length; i++) {
    if (String(tv(gT, i, "SeriesID")) !== seriesId) continue;
    var date = formatDateIso(tv(gT, i, "Date")), gameId = String(tv(gT, i, "GameID"));
    existingDates[date] = true;
    // A bulk edit never rewrites started matches or revives a cancelled exception.
    if (date < effective || hasGameStartedOrPassed(date, formatTimeStr(tv(gT, i, "Time"))) || String(tv(gT, i, "Status")) === "Cancelled") continue;
    var cfg = seriesConfigAt(rT, sr, date), recipients = rsvpEmailsForGame(gameId);
    if (!seriesWantsDate(cfg, date)) {
      var old = gameRecord(i);
      tw(gT, i, "Status", "Cancelled");
      appendGameActivityLog(i, admin.fullName, "Cancelled by recurring series change");
      cancelledGames.push({ title: old.title, date: old.date, time: old.time, location: old.location });
      recipients.forEach(function(em) { notified[em] = true; });
      cancelled++;
      continue;
    }
    var oldTime = formatTimeStr(tv(gT, i, "Time")), oldLocation = String(tv(gT, i, "LocationName"));
    var moved = oldTime !== cfg.Time || oldLocation !== cfg.LocationName;
    var committed = gameCommittedCounts(gameId);
    var upd = { Title: cfg.Title, LocationName: cfg.LocationName, Time: "'" + cfg.Time, MaxCap: Math.max(cfg.MaxCap, committed.reserved), MinThreshold: cfg.MinThreshold };
    if (moved) {
      var diff = [];
      if (oldTime !== cfg.Time) diff.push("🕒 " + oldTime + " → " + cfg.Time);
      if (oldLocation !== cfg.LocationName) diff.push("📍 " + oldLocation + " → " + cfg.LocationName);
      upd.ScheduleChangeAt = nowMs; upd.ScheduleChangeSummary = diff.join(" | ");
      recipients.forEach(function(em) { notified[em] = true; });
    }
    twMany(gT, i, upd);
    var recon = reconcileGameWaitlistAndStatus(gameId);
    if (recon.promotedEmails.length) sendWaitlistPromotionEmails(recon.gameObj, recon.promotedEmails);
    if (recon.statusFlipped) notifyStatusFlip(recon.gameObj, recon.newGameStatus, recon.recalculatedConfirmed);
    updated++;
  }
  var horizon = isoPlusDays(today, SERIES_WINDOW_DAYS), newGames = [], stamp = new Date().getTime();
  seriesOccurrences(rT, sr, today, horizon).forEach(function(item) {
    var date = item.date, cfg = item.config;
    if (existingDates[date] || hasGameStartedOrPassed(date, cfg.Time)) return;
    // Keep old deliberate deletions before the apply date, but fill any newly
    // opened dates with the configuration that is effective on each actual day.
    if (isEdit && date < effective && previousThrough && date <= previousThrough) return;
    existingDates[date] = true;
    newGames.push({ GameID: "G_" + stamp + "_" + newGames.length, Title: cfg.Title, LocationName: cfg.LocationName, Date: "'" + date, Time: "'" + cfg.Time, MaxCap: cfg.MaxCap, MinThreshold: cfg.MinThreshold, Status: "Pending", SeriesID: seriesId, MatchFormat: "Standard",
      AdminActivityJSON: JSON.stringify([{ actor: admin.fullName, summary: "Created by recurring series", timestamp: formatDateTimeStr(new Date()) }]) });
  });
  taddMany(gT, newGames); created = newGames.length;
  // Restart the rolling checkpoint at this horizon after an edit. Existing later
  // dates are deduplicated, while newly selected weekdays can still be generated.
  tw(rT, sr, "GeneratedThrough", "'" + horizon);
  var affected = Object.keys(notified);
  if (affected.length) notifyMail(affected, mailSeriesUpdated({ title: title, effectiveDate: effective, endDate: end, time: time, location: location, cancelledGames: cancelledGames }), {
    en: ["⚠️ Schedule change: " + title, "The series changed from " + mailDate(effective, "en") + ". Check cancelled dates and re-confirm any changed match."],
    es: ["⚠️ Cambio de horario: " + title, "La serie cambió desde el " + mailDate(effective, "es") + ". Revisa las fechas canceladas y vuelve a confirmar los partidos que cambiaron."]
  }, { critical: true, supersedeKey: "series-update:" + seriesId });
  return okData(admin, isEdit ? tr("msgSeriesUpdated", { date: effective, u: updated, c: created, r: cancelled }) : tr("msgSeriesCreated", { c: created }));
}

// Rolling window: each new day, every active series gets its matches up to four weeks ahead.
// Dates already covered once are never re-created, so a match an admin deleted stays deleted.
function extendRecurringSeries() {
  var rT = T("Recurring");
  if (!rT.rows.length || !tHas(rT, "GeneratedThrough")) return 0;
  var gT = T("Games"), today = formatDateIso(new Date()), horizon = isoPlusDays(today, SERIES_WINDOW_DAYS), existing = {};
  for (var g = 0; g < gT.rows.length; g++) {
    var sid = String(tv(gT, g, "SeriesID"));
    if (sid) existing[sid + "|" + formatDateIso(tv(gT, g, "Date"))] = true;
  }
  var add = [], stamp = new Date().getTime();
  for (var r = 0; r < rT.rows.length; r++) {
    if (String(tv(rT, r, "Status") || "Active") !== "Active") continue;
    var seriesId = String(tv(rT, r, "SeriesID")), through = formatDateIso(tv(rT, r, "GeneratedThrough"));
    var from = through && isoPlusDays(through, 1) > today ? isoPlusDays(through, 1) : today;
    if (!seriesId || from > horizon) continue;
    seriesOccurrences(rT, r, from, horizon).forEach(function(item) {
      var date = item.date, cfg = item.config, key = seriesId + "|" + date;
      if (existing[key] || hasGameStartedOrPassed(date, cfg.Time)) return;
      existing[key] = true;
      add.push({ GameID: "G_" + stamp + "_" + add.length, Title: cfg.Title, LocationName: cfg.LocationName, Date: "'" + date, Time: "'" + cfg.Time, MaxCap: cfg.MaxCap, MinThreshold: cfg.MinThreshold, Status: "Pending", SeriesID: seriesId, MatchFormat: "Standard",
        AdminActivityJSON: JSON.stringify([{ actor: "Recurring series", summary: "Opened 4 weeks ahead", timestamp: formatDateTimeStr(new Date()) }]) });
    });
    tw(rT, r, "GeneratedThrough", "'" + horizon);
  }
  taddMany(gT, add);
  return add.length;
}

// Runs at most once a day, on the first request of the day (and from the 9 AM job), so the
// window keeps rolling even if a scheduled trigger is ever missing.
function extendSeriesIfNewDay() {
  var props = PropertiesService.getScriptProperties();
  var today = formatDateIso(new Date());
  if (props.getProperty("SERIES_EXTENDED_ON") === today) return;
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) return; // busy right now; the next request will do it
  try {
    if (props.getProperty("SERIES_EXTENDED_ON") === today) return;
    __TBL = {};
    extendRecurringSeries();
    props.setProperty("SERIES_EXTENDED_ON", today);
  } catch (e) {
    Logger.log("Series window not extended: " + e);
  } finally {
    lock.releaseLock();
  }
}

function adminStopRecurringSeries(data) {
  var admin = requireAdmin(data);
  if (!admin) return fail(tr("errAdminOnly"));
  var seriesId = String(data.seriesId || ""), today = formatDateIso(new Date());
  var effective = normalizeDateInput(data.effectiveDate) || today;
  if (effective < today) return fail(tr("errSeriesEffectivePast"));
  var rT = T("Recurring"), gT = T("Games");
  var sr = tfind(rT, function(r) { return String(tv(rT, r, "SeriesID")) === seriesId; });
  if (sr === -1) return fail(tr("errNotFound"));
  var title = String(tv(rT, sr, "Title")), cutoff = isoPlusDays(effective, -1);
  var changes = seriesTimeline(rT, sr).map(function(change) {
    var cfg = {}, old = change.config;
    Object.keys(old).forEach(function(k) { cfg[k] = old[k]; });
    if (cfg.EndDate > cutoff) cfg.EndDate = cutoff;
    return { effectiveDate: change.effectiveDate, config: cfg };
  });
  var end = formatDateIso(tv(rT, sr, "EndDate"));
  if (end > cutoff) end = cutoff;
  // Keep the definition so dates before a future stop can still enter the window.
  twMany(rT, sr, { EndDate: "'" + end, Status: effective <= today ? "Stopped" : "Active", ChangesJSON: JSON.stringify(changes), UpdatedAt: "'" + today });
  var cancelledGames = [], notified = {}, cancelled = 0;
  for (var i = 0; i < gT.rows.length; i++) {
    if (String(tv(gT, i, "SeriesID")) !== seriesId) continue;
    var date = formatDateIso(tv(gT, i, "Date")), time = formatTimeStr(tv(gT, i, "Time"));
    if (date < effective || hasGameStartedOrPassed(date, time) || String(tv(gT, i, "Status")) === "Cancelled") continue;
    var g = gameRecord(i);
    tw(gT, i, "Status", "Cancelled");
    appendGameActivityLog(i, admin.fullName, "Cancelled by series stop from " + effective);
    cancelled++;
    var recipients = rsvpEmailsForGame(g.id);
    if (recipients.length) {
      cancelledGames.push({ title: g.title, date: g.date, time: g.time, location: g.location });
      recipients.forEach(function(em) { notified[em] = true; });
    }
  }
  var affected = Object.keys(notified);
  if (affected.length) {
    cancelledGames.sort(function(a, b) { return a.date.localeCompare(b.date); });
    notifyMail(affected, mailSeriesStopped({ title: title, effectiveDate: effective, games: cancelledGames }), {
      en: ["🚫 Series stopped: " + title, "Upcoming matches from " + mailDate(effective, "en") + " are cancelled. Earlier matches stay scheduled."],
      es: ["🚫 Serie detenida: " + title, "Se cancelaron los partidos desde el " + mailDate(effective, "es") + ". Los anteriores siguen programados."]
    }, { critical: true, supersedeKey: "series-update:" + seriesId });
  }
  return okData(admin, tr("msgSeriesStopped", { title: title, date: effective, n: cancelled, p: affected.length }));
}

function cleanTitle(v) {
  return String(v || "").replace(/[<>\u0000-\u001F]/g, "").replace(/\s+/g, " ").trim().slice(0, 60);
}

function adminSaveGame(data) {
  var admin = requireAdmin(data);
  if (!admin) return fail(tr("errAdminOnly"));
  var gT = T("Games");
  var isEdit = !!data.gameId;
  var gameId = String(data.gameId || ("G_" + new Date().getTime()));
  var newDate = normalizeDateInput(data.date);
  var newTime = normalizeTimeInput(data.time);
  var newLoc = String(data.locationName || "").trim();
  var newTitle = cleanTitle(data.title);
  var existingRow = isEdit ? findGameRow(gameId) : -1;
  var existing = existingRow !== -1 ? gameRecord(existingRow) : null;
  var mm = readMinMax(data, existing ? { min: existing.minThreshold, max: existing.maxCap } : null);
  var settingsNow = getLeagueSettings();
  var newCapEmail = String(data.captainEmail || "").toLowerCase().trim();
  var newFormat = String(data.matchFormat || "Standard").trim().slice(0, 40) || "Standard";

  if (!newTitle) return fail(tr("errTitle"));
  if (!newDate) return fail(tr("errDate"));
  if (!newTime) return fail(tr("errTime"));
  if (!mm) return fail(tr("errMinMax"));
  if (!newLoc || !activeLocationExists(newLoc)) return fail(tr("errLocation"));
  if (!settingsNow.featCaptains) newCapEmail = existing ? existing.captainEmail : "";   // switch off: leave captains as they are
  if (newCapEmail && !getValidCaptainProfile(newCapEmail)) return fail(tr("errCaptain"));

  var oldGame = null, gr = -1, scheduleChanged = false, captainChangedTo = "", recon = null;
  if (isEdit) {
    gr = findGameRow(gameId);
    if (gr === -1) return fail(tr("errGameNotFound"));
    oldGame = gameRecord(gr);
    var headcount = 0;
    var sT = T("Signups");
    for (var s = 0; s < sT.rows.length; s++) {
      if (String(tv(sT, s, "GameID")) === gameId && String(tv(sT, s, "Status")) === "Yes") headcount += 1 + parseGuestNamesList(tv(sT, s, "PlusOne")).length;
    }
    if (mm.max < headcount) return fail(tr("errCapBelowRoster", { n: headcount }));
    twMany(gT, gr, { Title: newTitle, LocationName: newLoc, Date: "'" + newDate, Time: "'" + newTime, MaxCap: mm.max, MinThreshold: mm.min, CaptainEmail: newCapEmail, MatchFormat: newFormat });
    if (newCapEmail && newCapEmail !== oldGame.captainEmail) {
      captainChangedTo = newCapEmail;
      appendGameActivityLog(gr, admin.fullName, "Assigned Match Captain: " + newCapEmail);
    } else if (!newCapEmail && oldGame.captainEmail) {
      appendGameActivityLog(gr, admin.fullName, "Cleared Match Captain (Managed by System Admins)");
    }
    scheduleChanged = (oldGame.date !== newDate || oldGame.time !== newTime || oldGame.location !== newLoc);
    if (scheduleChanged) {
      var diffParts = [];
      if (oldGame.date !== newDate) diffParts.push("🗓️ " + oldGame.date + " → " + newDate);
      if (oldGame.time !== newTime) diffParts.push("🕒 " + oldGame.time + " → " + newTime);
      if (oldGame.location !== newLoc) diffParts.push("📍 " + oldGame.location + " → " + newLoc);
      var diffSummary = diffParts.join(" | ");
      twMany(gT, gr, { ScheduleChangeAt: new Date().getTime(), ScheduleChangeSummary: diffSummary });
      appendGameActivityLog(gr, admin.fullName, "Schedule Change: " + diffSummary);
    }
    recon = reconcileGameWaitlistAndStatus(gameId);
  } else {
    tadd(gT, { GameID: gameId, Title: newTitle, LocationName: newLoc, Date: "'" + newDate, Time: "'" + newTime, MaxCap: mm.max, MinThreshold: mm.min, Status: "Pending", CaptainEmail: newCapEmail, MatchFormat: newFormat,
      AdminActivityJSON: JSON.stringify([{ actor: admin.fullName, summary: "Created Match" + (newCapEmail ? " (Captain: " + newCapEmail + ")" : " (Managed by System Admins)"), timestamp: formatDateTimeStr(new Date()) }]) });
    if (newCapEmail) captainChangedTo = newCapEmail;
  }

  if (captainChangedTo && settingsNow.featCaptains) {
    notifyMail([captainChangedTo], mailCaptainAssigned({ title: newTitle, date: newDate, time: newTime, location: newLoc }), {
      en: ["👑 You're match captain!", newTitle + " · " + mailWhen(newDate, newTime, "en")],
      es: ["👑 ¡Eres el capitán!", newTitle + " · " + mailWhen(newDate, newTime, "es")]
    });
  }
  if (recon && recon.promotedEmails && recon.promotedEmails.length > 0) sendWaitlistPromotionEmails(recon.gameObj, recon.promotedEmails);
  if (recon && recon.statusFlipped) notifyStatusFlip(recon.gameObj, recon.newGameStatus, recon.recalculatedYes);

  if (isEdit) {
    var notifiedCount = 0;
    if (scheduleChanged && oldGame.status !== "Cancelled" && newDate >= formatDateIso(new Date())) {
      notifiedCount = notifyGameChangeToRsvps(oldGame, { id: gameId, title: newTitle, locationName: newLoc, date: newDate, time: newTime });
    }
    return okData(admin, scheduleChanged ? tr("msgGameUpdatedNotified", { n: notifiedCount }) : tr("msgGameUpdated"));
  }
  return okData(admin, tr("msgGameCreated"), { createdGameId: gameId });
}

function notifyGameChangeToRsvps(oldGame, newGame) {
  var appData = getAppData();
  var game = appData.games.filter(function(g) { return String(g.GameID) === String(newGame.id); })[0];
  if (!game || !game.roster) return 0;
  var rsvpRecipients = game.roster.filter(function(r) { return r.status === "Yes" || r.status === "Maybe" || r.status === "Waitlist"; }).map(function(r) { return r.email; });
  if (rsvpRecipients.length === 0) return 0;
  notifyMail(rsvpRecipients, mailScheduleChange({ oldGame: oldGame, newGame: newGame, address: game.locationDetails ? game.locationDetails.address : "" }), {
    en: ["⚠️ Schedule change: " + newGame.title, "Now " + mailWhen(newGame.date, newGame.time, "en") + " at " + newGame.locationName + ". Tap to re-confirm."],
    es: ["⚠️ Cambio de horario: " + newGame.title, "Ahora " + mailWhen(newGame.date, newGame.time, "es") + " en " + newGame.locationName + ". Toca para reconfirmar."]
  });
  return rsvpRecipients.length;
}

function rsvpEmailsForGame(gameId) {
  var sT = T("Signups");
  var out = [];
  for (var s = 0; s < sT.rows.length; s++) {
    if (String(tv(sT, s, "GameID")) !== String(gameId)) continue;
    var st = String(tv(sT, s, "Status"));
    var em = String(tv(sT, s, "Email")).toLowerCase().trim();
    if ((st === "Yes" || st === "Maybe" || st === "Waitlist") && out.indexOf(em) === -1) out.push(em);
  }
  return out;
}

function adminCancelGame(data) {
  var admin = requireAdmin(data);
  if (!admin) return fail(tr("errAdminOnly"));
  var gr = findGameRow(data.gameId);
  if (gr === -1) return fail(tr("errGameNotFound"));
  var g = gameRecord(gr);
  if (g.status === "Cancelled") return fail(tr("errAlreadyCancelled"));
  tw(T("Games"), gr, "Status", "Cancelled");
  appendGameActivityLog(gr, admin.fullName, "Cancelled Match");
  var recipients = rsvpEmailsForGame(g.id);
  var notifiedCount = 0;
  if (g.date >= formatDateIso(new Date()) && recipients.length > 0) {
    notifiedCount = recipients.length;
    notifyMail(recipients, mailCancelled(g), {
      en: ["🚫 Match cancelled: " + g.title, g.title + " · " + mailWhen(g.date, g.time, "en") + " is cancelled."],
      es: ["🚫 Partido cancelado: " + g.title, g.title + " · " + mailWhen(g.date, g.time, "es") + " fue cancelado."]
    });
  }
  return okData(admin, tr("msgCancelled", { n: notifiedCount }));
}

function adminRestoreGame(data) {
  var admin = requireAdmin(data);
  if (!admin) return fail(tr("errAdminOnly"));
  var gr = findGameRow(data.gameId);
  if (gr === -1) return fail(tr("errGameNotFound"));
  var g = gameRecord(gr);
  if (g.status !== "Cancelled") return fail(tr("errNotCancelled"));
  tw(T("Games"), gr, "Status", "Pending");
  appendGameActivityLog(gr, admin.fullName, "Restored Match (un-cancelled)");
  var recon = reconcileGameWaitlistAndStatus(g.id);
  if (recon && recon.promotedEmails && recon.promotedEmails.length > 0) sendWaitlistPromotionEmails(recon.gameObj, recon.promotedEmails);
  var recipients = rsvpEmailsForGame(g.id);
  if (g.date >= formatDateIso(new Date()) && recipients.length > 0) {
    notifyMail(recipients, mailRestored(g), {
      en: ["✅ Back on: " + g.title, g.title + " · " + mailWhen(g.date, g.time, "en") + " is back on. Your answer still stands."],
      es: ["✅ Vuelve el partido: " + g.title, g.title + " · " + mailWhen(g.date, g.time, "es") + " vuelve al calendario. Tu respuesta sigue activa."]
    });
  }
  return okData(admin, tr("msgRestored", { n: recipients.length }));
}

function adminDeleteGame(data) {
  var admin = requireAdmin(data);
  if (!admin) return fail(tr("errAdminOnly"));
  if (!data.hardDelete) return adminCancelGame(data);
  var gr = findGameRow(data.gameId);
  if (gr === -1) return fail(tr("errGameNotFound"));
  var g = gameRecord(gr);
  var sT = T("Signups");
  var signupCount = 0;
  for (var s = 0; s < sT.rows.length; s++) if (String(tv(sT, s, "GameID")) === g.id) signupCount++;
  if (g.status !== "Cancelled" && signupCount > 0) return fail(tr("errDeleteNeedsCancel"));
  tdel(T("Games"), gr);
  for (var k = sT.rows.length - 1; k >= 0; k--) if (String(tv(sT, k, "GameID")) === g.id) tdel(sT, k);
  return okData(admin, tr("msgDeleted"));
}

function adminSaveLocation(data) {
  var admin = requireAdmin(data);
  if (!admin) return fail(tr("errAdminOnly"));
  var t = T("Locations");
  var name = cleanTitle(data.name).slice(0, 50);
  var address = String(data.address || "").replace(/[<>]/g, "").trim().slice(0, 200);
  var lat = String(data.lat || "").trim(), lng = String(data.lng || "").trim();
  var notes = String(data.notes || "").replace(/[<>]/g, "").trim().slice(0, 200);
  if (!name || !address) return fail(tr("errLocationFields"));
  if ((lat && isNaN(Number(lat))) || (lng && isNaN(Number(lng)))) return fail(tr("errLatLng"));
  var dup = tfind(t, function(r) { return String(tv(t, r, "Name")).toLowerCase().trim() === name.toLowerCase() && String(tv(t, r, "LocationID")) !== String(data.locationId || ""); });
  if (dup !== -1) return fail(tr("errLocationDup"));

  if (data.locationId) {
    var r = tfind(t, function(i) { return String(tv(t, i, "LocationID")) === String(data.locationId); });
    if (r === -1) return fail(tr("errNotFound"));
    var oldName = String(tv(t, r, "Name"));
    twMany(t, r, { Name: name, Address: address, Latitude: lat, Longitude: lng, Notes: notes });
    if (oldName && oldName !== name) {
      var gT = T("Games");
      for (var g = 0; g < gT.rows.length; g++) if (String(tv(gT, g, "LocationName")) === oldName) tw(gT, g, "LocationName", name);
      var rT = T("Recurring");
      for (var q = 0; q < rT.rows.length; q++) if (String(tv(rT, q, "LocationName")) === oldName) tw(rT, q, "LocationName", name);
    }
    return okData(admin, tr("msgLocationUpdated"));
  }
  tadd(t, { LocationID: "LOC_" + new Date().getTime(), Name: name, Address: address, Latitude: lat, Longitude: lng, Notes: notes, LocationStatus: "Active" });
  return okData(admin, tr("msgLocationCreated"));
}

function adminToggleLocationStatus(data) {
  var admin = requireAdmin(data);
  if (!admin) return fail(tr("errAdminOnly"));
  var t = T("Locations");
  var r = tfind(t, function(i) { return String(tv(t, i, "LocationID")) === String(data.locationId); });
  if (r === -1) return fail(tr("errNotFound"));
  var newStatus = String(data.newStatus) === "Active" ? "Active" : "Inactive";
  tw(t, r, "LocationStatus", newStatus);
  return okData(admin, tr(newStatus === "Active" ? "msgLocationOn" : "msgLocationOff"));
}

function adminDeleteLocation(data) {
  var admin = requireAdmin(data);
  if (!admin) return fail(tr("errAdminOnly"));
  var t = T("Locations");
  var r = tfind(t, function(i) { return String(tv(t, i, "LocationID")) === String(data.locationId); });
  if (r === -1) return fail(tr("errNotFound"));
  var locName = String(tv(t, r, "Name")).toLowerCase().trim();
  var gT = T("Games");
  var used = tfind(gT, function(i) { return String(tv(gT, i, "LocationName")).toLowerCase().trim() === locName; }) !== -1;
  if (used) {
    tw(t, r, "LocationStatus", "Inactive");
    return okData(admin, tr("msgLocationArchived"));
  }
  tdel(t, r);
  return okData(admin, tr("msgLocationDeleted"));
}

function sendInviteEmail(email, inviterName) {
  var link = buildInviteUrl(email);
  var intro = "Join our weekday soccer group. We play Monday–Friday at 12:15 PM Pacific; check each match for the field and any changes.";
  var spanish = "Únete a nuestro grupo de fútbol entre semana. Jugamos de lunes a viernes a las 12:15 PM, hora del Pacífico; revisa cada partido para ver la cancha y cualquier cambio.";
  var plain = "Pitch Pulse League\n\n" + String(inviterName || MAIL_SENDER_NAME) + " invited you.\n" + intro + "\n\n" + spanish + "\n\nCreate your profile / Crear tu perfil:\n" + link +
    "\n\n1. Open your personal link.\n2. Enter your full name and choose a six-digit PIN.\n3. Open today's match and select I'm in, Maybe, or Out.\n\nIf the button does not work, copy the full link into Safari or Chrome. This personal link is for " + email + "; please do not forward it.";
  var html = "<div style='font-family:Arial,sans-serif;max-width:560px;padding:24px;color:#0f172a;background:#fff;'>" +
    "<h1>Pitch Pulse</h1><p>" + escapeHtmlServer(inviterName || MAIL_SENDER_NAME) + " invited you / te invita.</p><p>" + escapeHtmlServer(intro) + "</p><p>" + escapeHtmlServer(spanish) + "</p>" +
    "<p><a href='" + escapeHtmlServer(link) + "' style='display:inline-block;background:#047857;color:#fff;padding:14px 22px;border-radius:10px;text-decoration:none;font-weight:bold;'>Create my profile / Crear mi perfil</a></p>" +
    "<ol><li>Open your personal link / Abre tu enlace personal.</li><li>Enter your name and choose a six-digit PIN / Escribe tu nombre y elige un PIN de seis dígitos.</li><li>Open a match and answer / Abre un partido y responde.</li></ol>" +
    "<p>If the button does not work, copy this full link into Safari or Chrome:<br>Si el botón no funciona, copia este enlace completo en Safari o Chrome:</p>" +
    "<p style='overflow-wrap:anywhere;word-break:break-all;'><a href='" + escapeHtmlServer(link) + "'>" + escapeHtmlServer(link) + "</a></p>" +
    "<p>This personal link is for " + escapeHtmlServer(email) + ". Please do not forward it. / Este enlace es personal; no lo reenvíes.</p></div>";
  return dispatchLeagueNotification([email], "⚽ Pitch Pulse invitation / Invitación", html, "", "", SITE_URL,
    { emailOnly: true, invite: true, priority: 40, mail: function() { return { subject: "⚽ Pitch Pulse invitation / Invitación", html: html, text: plain }; } });
}

function adminSendInvite(data) {
  var admin = requireAdmin(data); if (!admin) return fail(tr("errAdminOnly"));
  var email = String(data.email || "").toLowerCase().trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return fail(tr("errEmail"));
  if (findProfileRow(email) !== -1) return fail(tr("errAlreadyMember", { email: email }));
  var t = T("Invites"), r = tfind(t, function(i) { return String(tv(t, i, "Email")).toLowerCase().trim() === email; });
  if (r === -1) r = tadd(t, { Email: email, Status: "Queued", InvitedDate: "'" + formatDateIso(new Date()) });
  var result = sendInviteEmail(email, admin.fullName);
  twMany(t, r, { InvitedDate: "'" + formatDateIso(new Date()), Status: result.queued ? "Pending" : "Queued" });
  return okData(admin, tr("msgInviteSent", { email: email }));
}

// Paste a whole list: what fits under today's email limit goes out now, the rest waits in the
// queue and goes out automatically (checked every hour) as the limit allows.
var INVITE_BATCH_PER_REQUEST = 25;
function adminSendInvites(data) {
  var admin = requireAdmin(data);
  if (!admin) return fail(tr("errAdminOnly"));
  var list = Array.isArray(data.emails) ? data.emails : String(data.emails || "").split(/[\s,;]+/);
  var seen = {}, bad = [], members = 0, already = 0, added = 0;
  var t = T("Invites");
  list.forEach(function(raw) {
    var em = String(raw || "").replace(/^mailto:/i, "").replace(/[<>()"']/g, "").toLowerCase().trim();
    if (!em || seen[em]) return;
    seen[em] = true;
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(em)) { bad.push(em); return; }
    if (findProfileRow(em) !== -1) { members++; return; }
    var r = tfind(t, function(i) { return String(tv(t, i, "Email")).toLowerCase().trim() === em; });
    if (r !== -1) { already++; return; }
    tadd(t, { Email: em, Status: "Queued", InvitedDate: "'" + formatDateIso(new Date()) });
    added++;
  });
  if (!added && bad.length) return fail(tr("errInviteListBad", { list: bad.slice(0, 5).join(", ") }));
  var sendingNow = scheduleInviteBatch(admin.fullName, INVITE_BATCH_PER_REQUEST);
  var queued = countQueuedInvites();
  var parts = [tr("msgInvitesSent", { n: sendingNow })];
  if (queued) parts.push(tr("msgInvitesQueued", { n: queued }));
  if (members || already) parts.push(tr("msgInvitesSkipped", { n: members + already }));
  if (bad.length) parts.push(tr("msgInvitesBad", { list: bad.slice(0, 5).join(", ") }));
  return okData(admin, parts.join(" "));
}

function countQueuedInvites() {
  var t = T("Invites"), n = 0;
  for (var r = 0; r < t.rows.length; r++) if (String(tv(t, r, "Status")) === "Queued" && findProfileRow(tv(t, r, "Email")) === -1) n++;
  return n;
}

// Marks the next batch as sent (under the lock) and sends the emails once the lock is released.
function scheduleInviteBatch(inviterName, maxCount) {
  var t = T("Invites"), queued = 0;
  for (var r = 0; r < t.rows.length && queued < maxCount; r++) {
    var em = String(tv(t, r, "Email") || "").toLowerCase().trim();
    if (!em || String(tv(t, r, "Status")) !== "Queued") continue;
    if (findProfileRow(em) !== -1) { tw(t, r, "Status", "Registered"); continue; }
    var result = sendInviteEmail(em, inviterName);
    if (!result.queued) break;
    twMany(t, r, { Status: "Pending", InvitedDate: "'" + formatDateIso(new Date()) }); queued++;
  }
  return queued;
}

// Hourly: send whatever the daily limit allows from the queue.
function processInviteQueue() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return 0;
  var n = 0;
  try {
    __HAS_LOCK = true;
    __TBL = {};
    if (!countQueuedInvites()) return 0;
    n = scheduleInviteBatch(MAIL_SENDER_NAME, 50);
  } finally {
    __HAS_LOCK = false;
    lock.releaseLock();
  }
  runAfterLockJobs();
  return n;
}

function adminDeleteInvite(data) {
  var admin = requireAdmin(data);
  if (!admin) return fail(tr("errAdminOnly"));
  var targetEmail = String(data.email || "").toLowerCase().trim();
  if (isProtectedAdmin(targetEmail)) return fail(tr("errProtectedInvite"));
  var t = T("Invites");
  var r = tfind(t, function(i) { return String(tv(t, i, "Email")).toLowerCase().trim() === targetEmail; });
  if (r !== -1) tdel(t, r);
  return okData(admin, tr("msgInviteRevoked"));
}

function adminUpdateUserRole(data) {
  var admin = requireAdmin(data);
  if (!admin) return fail(tr("errAdminOnly"));
  var targetEmail = String(data.targetEmail || "").toLowerCase().trim();
  var newRole = String(data.newRole || "Player").trim() === "Admin" ? "Admin" : "Player";
  if (isProtectedAdmin(targetEmail) && newRole !== "Admin") return fail(tr("errProtectedDemote"));
  var r = findProfileRow(targetEmail);
  if (r === -1) return fail(tr("errNoProfile"));
  tw(T("Profiles"), r, "Role", newRole);
  return okData(admin, tr(newRole === "Admin" ? "msgMadeAdmin" : "msgMadePlayer", { email: targetEmail }));
}

function adminToggleSuspendUser(data) {
  var admin = requireAdmin(data);
  if (!admin) return fail(tr("errAdminOnly"));
  var targetEmail = String(data.targetEmail || "").toLowerCase().trim();
  if (isProtectedAdmin(targetEmail)) return fail(tr("errProtectedSuspend"));
  var pr = findProfileRow(targetEmail);
  if (pr === -1) return fail(tr("errNoProfile"));
  var newStatus = String(data.newStatus || "Suspended").trim() === "Active" ? "Active" : "Suspended";
  tw(T("Profiles"), pr, "AccountStatus", newStatus);
  var promoted = [];
  if (newStatus === "Suspended") {
    revokeAllSessions(targetEmail);
    var todayStr = formatDateIso(new Date());
    var gT = T("Games");
    var futureGames = {};
    for (var g = 0; g < gT.rows.length; g++) {
      if (!hasGameStartedOrPassed(formatDateIso(tv(gT, g, "Date")), formatTimeStr(tv(gT, g, "Time")))) {
        futureGames[String(tv(gT, g, "GameID"))] = g;
        if (String(tv(gT, g, "CaptainEmail") || "").toLowerCase().trim() === targetEmail) tw(gT, g, "CaptainEmail", "");
      }
    }
    var sT = T("Signups");
    var affected = {};
    for (var s = 0; s < sT.rows.length; s++) {
      var gId = String(tv(sT, s, "GameID"));
      var st = String(tv(sT, s, "Status"));
      if (String(tv(sT, s, "Email")).toLowerCase().trim() === targetEmail && futureGames[gId] !== undefined && (st === "Yes" || st === "Waitlist" || st === "Maybe")) {
        twMany(sT, s, { Status: "No", BringBall: false, BringPinnies: false, PlayGK: false, PlusOne: "", AvailabilityNote: "" });
        affected[gId] = true;
      }
    }
    Object.keys(affected).forEach(function(gId) {
      var recon = reconcileGameWaitlistAndStatus(gId);
      syncTeamsForGame(futureGames[gId]);
      if (recon && ((recon.promotedEmails && recon.promotedEmails.length > 0) || recon.statusFlipped)) promoted.push(recon);
    });
  }
  promoted.forEach(function(r) {
    if (r.promotedEmails && r.promotedEmails.length) sendWaitlistPromotionEmails(r.gameObj, r.promotedEmails);
    if (r.statusFlipped && r.gameObj) notifyStatusFlip(r.gameObj, r.newGameStatus, r.recalculatedYes);
  });
  return okData(admin, tr(newStatus === "Suspended" ? "msgSuspended" : "msgReactivated", { email: targetEmail }));
}

function adminDeleteUser(data) {
  var admin = requireAdmin(data);
  if (!admin) return fail(tr("errAdminOnly"));
  var targetEmail = String(data.targetEmail || "").toLowerCase().trim();
  if (isProtectedAdmin(targetEmail)) return fail(tr("errProtectedRemove"));
  var sT = T("Signups");
  var hasHistory = tfind(sT, function(r) { return String(tv(sT, r, "Email")).toLowerCase().trim() === targetEmail; }) !== -1;
  if (hasHistory) return adminToggleSuspendUser({ requesterEmail: data.requesterEmail, requesterPin: data.requesterPin, targetEmail: targetEmail, newStatus: "Suspended" });
  var pT = T("Profiles");
  for (var k = pT.rows.length - 1; k >= 0; k--) if (String(tv(pT, k, "Email")).toLowerCase().trim() === targetEmail) tdel(pT, k);
  var iT = T("Invites");
  for (var j = iT.rows.length - 1; j >= 0; j--) if (String(tv(iT, j, "Email")).toLowerCase().trim() === targetEmail) tdel(iT, j);
  revokeAllSessions(targetEmail);
  return okData(admin, tr("msgUserRemoved", { email: targetEmail }));
}

// ============================================================================
// AUTOMATED EMAIL + PUSH TRIGGERS (Respects Player NotifyPref & AwayUntil)
// ============================================================================


// The day before: only matches still short of their minimum, and only players who haven't
// answered (away and suspended players are skipped). Push, plus a personal email with
// one-tap answer buttons.


// Every hour: lock in the switches of matches that just kicked off, send invites waiting in
// the queue, and (only if that message is switched on) the two-hour reminder push.


// Keeps Script Properties well under Google's 500 KB limit by dropping one-time markers
// from matches older than 30 days, expired sessions and stale lockouts.
function housekeeping() {
  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties();
  var now = new Date().getTime();
  var cutoff = formatDateIso(new Date(now - 30 * 86400000));
  var gameDates = {};
  var gT = T("Games");
  for (var r = 0; r < gT.rows.length; r++) gameDates[String(tv(gT, r, "GameID"))] = formatDateIso(tv(gT, r, "Date"));
  var removed = 0;
  Object.keys(all).forEach(function(k) {
    var m = k.match(/^(SENT_2HR|LATEDROP|LATEEARLY|NUDGE_COOLDOWN)_(G_\d+(?:_\d+(?=_|$))?)/);
    if (m) {
      var d = gameDates[m[2]];
      if (!d || d < cutoff) { props.deleteProperty(k); removed++; }
      return;
    }
    if (k.indexOf("SESSION_") === 0) { writeSessions(k.slice(8), readSessions(k.slice(8))); return; }
    if (k.indexOf("THROTTLE_") === 0 || k.indexOf("RESETREQ_") === 0 || k.indexOf("RESET_") === 0) {
      try {
        var o = JSON.parse(all[k]);
        if ((o.lockUntil !== undefined && o.lockUntil < now) || (o.day && o.day < formatDateIso(new Date())) || (o.expiresAt && o.expiresAt < now)) { props.deleteProperty(k); removed++; }
      } catch (e) {}
    }
  });
  return removed;
}

// ============================================================================
// KICKOFF-HOUR WEATHER ENGINE & UTILITY HELPERS
// ============================================================================
function escapeHtmlServer(str) {
  return String(str === undefined || str === null ? "" : str)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

function fetchBackendWeatherString(lat, lng, dateStr, timeStr) {
  return weatherText(fetchBackendWeather(lat, lng, dateStr, timeStr), "en");
}

function isTruthy(val) {
  return val === true || String(val || "").toLowerCase() === "true";
}

function getDataRows(sheet) {
  if (!sheet) return [];
  var rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return [];
  var headers = rows[0];
  var data = [];
  for (var i = 1; i < rows.length; i++) {
    var obj = {};
    for (var h = 0; h < headers.length; h++) obj[headers[h]] = rows[i][h];
    data.push(obj);
  }
  return data;
}

function formatDateIso(val) {
  if (!val) return "";
  if (Object.prototype.toString.call(val) === "[object Date]") {
    if (isNaN(val.getTime())) return "";
    return Utilities.formatDate(val, leagueTimeZone(), "yyyy-MM-dd");
  }
  return String(val).replace(/^'/, "").slice(0, 10);
}

function formatTimeStr(val) {
  if (!val) return "12:00 PM";
  if (Object.prototype.toString.call(val) === "[object Date]") return Utilities.formatDate(val, leagueTimeZone(), "h:mm a");
  var n = normalizeTimeInput(val);
  return n || String(val);
}

function formatDateTimeStr(val) {
  if (!val) return "";
  if (Object.prototype.toString.call(val) === "[object Date]") return Utilities.formatDate(val, leagueTimeZone(), "MMM d, h:mm a");
  return String(val);
}

function parseTimeMinutes(timeStr) {
  var str = String(timeStr || "12:00 PM").trim().toUpperCase();
  var isPM = str.indexOf("PM") !== -1;
  var isAM = str.indexOf("AM") !== -1;
  var parts = str.replace(/[^0-9:]/g, "").split(":");
  var hrs = parseInt(parts[0] || "12", 10);
  var mins = parseInt(parts[1] || "0", 10);
  if (isPM && hrs < 12) hrs += 12;
  if (isAM && hrs === 12) hrs = 0;
  return hrs * 60 + mins;
}

// "2026-10-30" (or a Date) → "2026-10-30"; anything that isn't a real calendar date → ""
function normalizeDateInput(v) {
  if (!v) return "";
  if (Object.prototype.toString.call(v) === "[object Date]") return formatDateIso(v);
  var m = String(v).trim().replace(/^'/, "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return "";
  var d = new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12));
  if (d.getUTCFullYear() !== Number(m[1]) || d.getUTCMonth() !== Number(m[2]) - 1 || d.getUTCDate() !== Number(m[3])) return "";
  return m[1] + "-" + m[2] + "-" + m[3];
}

// "6:30 pm", "18:30", "6:30PM" → "6:30 PM"; invalid → ""
function normalizeTimeInput(v) {
  if (!v) return "";
  if (Object.prototype.toString.call(v) === "[object Date]") return Utilities.formatDate(v, leagueTimeZone(), "h:mm a");
  var s = String(v).trim().replace(/^'/, "").toUpperCase();
  var m = s.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)?$/);
  if (!m) return "";
  var h = Number(m[1]), mi = Number(m[2]);
  if (mi > 59) return "";
  if (m[3]) {
    if (h < 1 || h > 12) return "";
    return h + ":" + (mi < 10 ? "0" + mi : mi) + " " + m[3];
  }
  if (h > 23) return "";
  var ap = h >= 12 ? "PM" : "AM";
  var h12 = h % 12 === 0 ? 12 : h % 12;
  return h12 + ":" + (mi < 10 ? "0" + mi : mi) + " " + ap;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================================
// SERVER MESSAGES (English / Spanish — follows the language chosen in the app)
// ============================================================================
function tr(key, p) {
  var dict = SRV_MSG[__LANG] || SRV_MSG.en;
  var s = dict[key] || SRV_MSG.en[key] || key;
  if (p) for (var k in p) s = s.split("{" + k + "}").join(String(p[k]));
  return s;
}

var SRV_MSG = {
  en: {
    errLocked: "Too many failed PIN attempts. Try again in {mins} minute(s).",
    errCancelled: "This match was cancelled — it's locked.",
    errScored: "The final score is in, so RSVPs are closed. Ask the captain or an admin to fix attendance.",
    errPast: "This match has already happened. Only the captain or an admin can adjust attendance.",
    errBusy: "The league is busy saving another change. Please try again in a moment.",
    errAdminOnly: "Admin access required.",
    errVaultPin: "Wrong PIN. Enter your actual 6-digit PIN to unlock the push keys.",
    errVaultPinNoChange: "Wrong PIN. Push settings were not changed.",
    msgPushSaved: "🔔 Push keys saved — push notifications are active.",
    msgPushCleared: "🔕 Push keys incomplete — players will get email instead.",
    errBadStatus: "Invalid RSVP choice.",
    errBadToken: "This link is invalid or expired.",
    errFeatureOff: "That feature is switched off for this match.",
    errGuestsLeagueOff: "Guests aren't allowed in this league right now.",
    errMaxGuests: "Guests per player must be a whole number from 0 to {max}.",
    rsvpGone: "That match isn't open for answers anymore.",
    errInviteListBad: "No valid email addresses found ({list}).",
    msgInvitesSent: "{n} invitation(s) queued for email.",
    msgInvitesQueued: "{n} more are queued and go out automatically as the daily email limit allows.",
    msgInvitesSkipped: "{n} skipped (already members or already invited).",
    msgInvitesBad: "Not valid: {list}.",
    msgBlastSentEmail: "Announcement posted. Email and phone notifications queued according to each player’s preferences ({n} players).",
    errBadPin: "Incorrect 6-digit PIN.",
    errSuspended: "Your account is suspended. Please contact a league admin.",
    errSysAdminRsvp: "The league system account can't play — log in with your personal account to RSVP.",
    msgMagicWaitlist: "⏳ The match is full — you're on the waitlist and will be moved in automatically if a spot opens.",
    msgMagicConfirmed: "✅ Done! Your RSVP is {status}.",
    statusYes: "YES", statusMaybe: "MAYBE", statusNo: "NO", statusWaitlist: "WAITLIST",
    errSession: "Your session expired. Please log in again.",
    errRegisterFields: "Enter your full name, a valid email, and a 6-digit PIN.",
    errInviteRequired: "{email} hasn't been invited yet. Ask a league admin for an invite.",
    errInviteLink: "Please open the personal link in your invitation email to finish signing up. (Can't find it? Ask an admin to tap Resend.)",
    errAccountExists: "You already have an account — log in with your PIN.",
    errPinFormat: "Your PIN must be exactly 6 digits.",
    errLogin: "Email or PIN is incorrect.",
    errUnauthorized: "Please log in again.",
    msgProfileSaved: "✅ Profile saved.",
    errBackupOnly: "Emergency recovery to {backup} is only for {primary}.",
    errPrimaryNotRegistered: "The primary admin profile hasn't been registered yet.",
    errNoProfile: "No player profile found for that email.",
    errResetCooldown: "A code was just sent. Please wait {secs} seconds before asking again.",
    errResetDaily: "Too many reset codes today. Try again tomorrow or contact an admin.",
    errResetMail: "Couldn't send the code email: {err}",
    msgResetSentBackup: "Emergency code sent to {backup} for {email}.",
    msgResetSent: "A 6-digit code (valid 15 minutes) was emailed to {email}.",
    errNewPinFormat: "Your new PIN must be exactly 6 digits.",
    errNoResetCode: "No active reset code. Please request a new one.",
    errResetExpired: "That code expired (15-minute limit). Please request a new one.",
    errResetTooMany: "Too many wrong codes — request a new one.",
    errResetBadCode: "That code isn't right.",
    msgPinUpdated: "✅ PIN updated. Log in with your new PIN (other devices were signed out).",
    errAuth: "Please log in again.",
    errGameNotFound: "Match not found — it may have been removed.",
    msgWaitlisted: "⏳ The match is full ({cap}). You're on the waitlist and will move in automatically if a spot opens.",
    msgPromoted: "{n} player(s) moved up from the waitlist.",
    msgGuestsReleased: "Your {n} guest spot(s) were released too.",
    errMustBeYes: "RSVP Yes first to add gear, guests or a timing note.",
    msgNoteSaved: "Timing note saved. An organizer notice was queued.",
    msgNoteCleared: "🕒 Timing note cleared.",
    errGuestsDisabled: "Guests are turned off for your account.",
    errGuestDuplicateOwn: "You entered '{name}' twice.",
    errGuestMax: "You can bring up to {max} guests.",
    errGuestTaken: "'{name}' is already coming as {host}'s guest.",
    errGuestIsMember: "'{name}' is a registered league member. They should answer from their own profile; ask an admin if this is a different person with the same name.",
    errMemberAlreadyGuest: "A guest with your name already has a spot with {host}. Ask them to remove that guest entry before you join. If it is a different person with the same name, ask an admin to clarify the guest name.",
    errNoShowBeforeKickoff: "No-shows can only be recorded after kickoff.",
    errNoShowNotConfirmed: "Only a player marked In can be recorded as a no-show.",
    errSeriesEffectivePast: "Choose today or a later date for a series change.",
    errGuestCapacity: "Not enough room for {n} guest(s): only {left} spot(s) left before the {cap}-player cap.",
    msgGuestsSaved: "👥 {n} guest(s) saved: {names}",
    msgGuestsCleared: "👥 Guests removed.",
    errBadGear: "Unknown gear option.",
    errGearOff: "The league isn't tracking that item right now.",
    errManagerOnly: "Only this match's captain or an admin can do that.",
    errAfterKickoff: "Kickoff has passed.",
    errNudgeCooldown: "Players were nudged recently. Try again in {mins} minute(s).",
    errNudgeNobody: "Everyone has already answered for this match!",
    msgNudged: "Reminder queued for {n} player(s). Sending starts on the next scheduled run.",
    errNotCaptain: "You're not the captain of this match.",
    msgSteppedDown: "You stepped down as captain. The admins now manage this match.",
    errSmallSidedMin: "Small-sided needs at least {n} players (5 v 5).",
    msgSmallSided: "🟢 Approved {format} — the match is ON with {n} players.",
    errAttendanceFuture: "Attendance opens on match day.",
    errAttendanceArgs: "Pick a player and Attended or No-show.",
    msgAttended: "✅ {name} marked as attended.",
    msgNoShow: "❌ {name} marked as no-show.",
    errMotmPick: "Pick a player to vote for.",
    errMotmSelf: "You can't vote for yourself.",
    errMotmLocked: "Voting opens after the scheduled match ends or an organizer saves the final score.",
    errMotmVoter: "Only players who played can vote.",
    errMotmNominee: "That player wasn't in this match.",
    msgMotmSaved: "⭐ Vote saved!",
    msgReconfirmedIn: "👍 Thanks — you're still in with the new time.",
    msgReconfirmedMaybe: "👍 Got it — you're still a maybe.",
    errTeamsMin: "Need at least 2 confirmed players to make teams.",
    msgTeamsDefaults: "👕 Default teams set: {a} {na} – {nb} {b}.",
    msgTeamsSmart: "👕 Teams balanced: {a} {na} – {nb} {b}.",
    errNoTeams: "No lineup yet — balance the teams first.",
    errNotInLineup: "That player isn't in the lineup.",
    msgMoved: "🔄 {name} moved to {team}.",
    errSwapPick: "Pick one player from each team.",
    msgSwapped: "🔄 Swapped {a} and {b}.",
    errScoreBeforeKickoff: "You can record the score once the match has kicked off.",
    msgScoreCleared: "Score cleared — RSVPs are open again.",
    errScoreValue: "Scores must be whole numbers from 0 to 99.",
    msgScoreSaved: "🏆 Final score saved: {a} {sa} – {sb} {b}.",
    errProtectedGuests: "Protected admins always keep guest privileges.",
    msgGuestsOn: "✅ Guests enabled for {email}.",
    msgGuestsOff: "🚫 Guests disabled for {email}.",
    msgDefaultTeamSet: "{name}'s default team is now {team}.",
    msgDefaultTeamCleared: "{name} no longer has a default team.",
    errTeamNameEmpty: "Team names can't be empty.",
    errTeamNamesSame: "The two teams need different names.",
    errShirt: "Pick a shirt color from the list.",
    errMinMax: "Minimum and maximum must be whole numbers, minimum 2, maximum 60, and min ≤ max.",
    msgSettingsSaved: "✅ League settings saved.",
    msgRulesSaved: "📖 League rules updated.",
    errBlastEmpty: "Write a message first.",
    errDirectTarget: "You can't send a message to that account.",
    msgDirectSent: "Message queued for {name}. Check Admin Settings for sending status.",
    msgBlastSent: "Announcement posted. Notifications queued for {n} player(s).",
    errNotFound: "Not found — it may have been removed.",
    msgBlastHidden: "Message hidden from players (admins still see it).",
    msgBlastShown: "Message visible to players again.",
    errSeriesDates: "Choose a start date and an end date after it.",
    errSeriesDays: "Pick at least one day of the week.",
    errTime: "Choose a valid kickoff time.",
    errLocation: "Choose a field from the list (add it under Fields first).",
    msgSeriesUpdated: "Series updated from {date}: {u} updated, {c} added, {r} cancelled.",
    msgSeriesCreated: "Series saved — the next {c} match(es) are open. New ones open automatically, 4 weeks ahead.",
    msgSeriesStopped: "Series '{title}' stops from {date}: {n} upcoming match(es) cancelled; notices queued for {p} player(s).",
    errTitle: "Give the match a title.",
    errDate: "Choose a valid date.",
    errCaptain: "That captain isn't an active player.",
    errCapBelowRoster: "Max players can't be lower than the {n} already confirmed.",
    msgGameUpdatedNotified: "Match updated. Reconfirmation notices queued for {n} player(s).",
    msgGameUpdated: "Match updated.",
    msgGameCreated: "New match published!",
    errAlreadyCancelled: "This match is already cancelled.",
    msgCancelled: "Match cancelled. Notices queued for {n} player(s); check Admin Settings for sending status.",
    errNotCancelled: "This match isn't cancelled.",
    msgRestored: "Match restored. Notices queued for {n} player(s).",
    errDeleteNeedsCancel: "Cancel the match first — matches with RSVPs can't be deleted directly.",
    msgDeleted: "Match permanently deleted.",
    errLocationFields: "Enter a field name and an address.",
    errLatLng: "Latitude and longitude must be numbers.",
    errLocationDup: "A field with that name already exists.",
    msgLocationUpdated: "Field updated.",
    msgLocationCreated: "Field added.",
    msgLocationOn: "Field reactivated.",
    msgLocationOff: "Field deactivated.",
    msgLocationArchived: "This field has match history, so it was deactivated instead of deleted.",
    msgLocationDeleted: "Unused field removed.",
    errEmail: "Enter a valid email address.",
    errAlreadyMember: "{email} is already a league member.",
    msgInviteNoQuota: "Invite saved for {email}, but today's email limit is used up — tap Resend tomorrow.",
    msgInviteMailFailed: "Invite saved for {email}, but the email failed: {err}",
    msgInviteSent: "Invite queued for {email}. Sending starts on the next scheduled run.",
    errProtectedInvite: "Protected admin invites can't be revoked.",
    msgInviteRevoked: "Invite revoked.",
    errProtectedDemote: "Protected admins can't be demoted.",
    msgMadeAdmin: "{email} is now an admin.",
    msgMadePlayer: "{email} is now a player.",
    errProtectedSuspend: "Protected admins can't be suspended.",
    msgSuspended: "{email} suspended — removed from upcoming matches and signed out.",
    msgReactivated: "{email} reactivated.",
    errProtectedRemove: "Protected admins can't be removed.",
    msgUserRemoved: "Removed {email} (no match history)."
  },
  es: {
    errLocked: "Demasiados intentos con PIN incorrecto. Intenta de nuevo en {mins} minuto(s).",
    errCancelled: "Este partido fue cancelado — está bloqueado.",
    errScored: "Ya se registró el marcador final, así que las confirmaciones están cerradas. Pide al capitán o a un admin que corrija la asistencia.",
    errPast: "Este partido ya se jugó. Solo el capitán o un admin pueden ajustar la asistencia.",
    errBusy: "La liga está guardando otro cambio. Intenta de nuevo en un momento.",
    errAdminOnly: "Se requiere acceso de administrador.",
    errVaultPin: "PIN incorrecto. Escribe tu PIN real de 6 dígitos para ver las claves push.",
    errVaultPinNoChange: "PIN incorrecto. No se cambió la configuración push.",
    msgPushSaved: "🔔 Claves push guardadas — las notificaciones push están activas.",
    msgPushCleared: "🔕 Faltan claves push — los jugadores recibirán correo.",
    errBadStatus: "Opción de asistencia no válida.",
    errBadToken: "Este enlace no es válido o ya venció.",
    errFeatureOff: "Esa función está desactivada para este partido.",
    errGuestsLeagueOff: "Por ahora la liga no permite invitados.",
    errMaxGuests: "Los invitados por jugador deben ser un número entero de 0 a {max}.",
    rsvpGone: "Ese partido ya no acepta respuestas.",
    errInviteListBad: "No encontramos correos válidos ({list}).",
    msgInvitesSent: "{n} invitación(es) en cola para correo.",
    msgInvitesQueued: "{n} más quedan en cola y saldrán solas según el límite diario de correos.",
    msgInvitesSkipped: "{n} omitida(s) (ya son miembros o ya tenían invitación).",
    msgInvitesBad: "No válidos: {list}.",
    msgBlastSentEmail: "Aviso publicado. Correo y notificaciones en cola según las preferencias de cada jugador ({n}).",
    errBadPin: "PIN de 6 dígitos incorrecto.",
    errSuspended: "Tu cuenta está suspendida. Habla con un administrador de la liga.",
    errSysAdminRsvp: "La cuenta del sistema no juega — entra con tu cuenta personal para confirmar.",
    msgMagicWaitlist: "⏳ El partido está lleno — quedaste en lista de espera y entras automáticamente si se libera un lugar.",
    msgMagicConfirmed: "✅ ¡Listo! Tu respuesta es {status}.",
    statusYes: "SÍ", statusMaybe: "TAL VEZ", statusNo: "NO", statusWaitlist: "EN ESPERA",
    errSession: "Tu sesión venció. Vuelve a iniciar sesión.",
    errRegisterFields: "Escribe tu nombre completo, un correo válido y un PIN de 6 dígitos.",
    errInviteRequired: "{email} todavía no tiene invitación. Pídela a un administrador.",
    errInviteLink: "Abre el enlace personal de tu correo de invitación para terminar el registro. (¿No lo encuentras? Pide a un admin que toque Reenviar.)",
    errAccountExists: "Ya tienes una cuenta — inicia sesión con tu PIN.",
    errPinFormat: "Tu PIN debe tener exactamente 6 dígitos.",
    errLogin: "El correo o el PIN no son correctos.",
    errUnauthorized: "Vuelve a iniciar sesión.",
    msgProfileSaved: "✅ Perfil guardado.",
    errBackupOnly: "La recuperación de emergencia a {backup} es solo para {primary}.",
    errPrimaryNotRegistered: "La cuenta principal de administrador aún no está registrada.",
    errNoProfile: "No hay ningún perfil con ese correo.",
    errResetCooldown: "Acabamos de enviar un código. Espera {secs} segundos antes de pedir otro.",
    errResetDaily: "Demasiados códigos hoy. Intenta mañana o habla con un admin.",
    errResetMail: "No se pudo enviar el correo con el código: {err}",
    msgResetSentBackup: "Código de emergencia enviado a {backup} para {email}.",
    msgResetSent: "Enviamos un código de 6 dígitos (válido 15 minutos) a {email}.",
    errNewPinFormat: "Tu nuevo PIN debe tener exactamente 6 dígitos.",
    errNoResetCode: "No hay un código activo. Pide uno nuevo.",
    errResetExpired: "El código venció (15 minutos). Pide uno nuevo.",
    errResetTooMany: "Demasiados códigos incorrectos — pide uno nuevo.",
    errResetBadCode: "Ese código no es correcto.",
    msgPinUpdated: "✅ PIN actualizado. Entra con tu nuevo PIN (se cerró la sesión en otros dispositivos).",
    errAuth: "Vuelve a iniciar sesión.",
    errGameNotFound: "No encontramos el partido — tal vez lo eliminaron.",
    msgWaitlisted: "⏳ El partido está lleno ({cap}). Quedaste en lista de espera y entras automáticamente si se libera un lugar.",
    msgPromoted: "{n} jugador(es) subieron desde la lista de espera.",
    msgGuestsReleased: "También se liberaron los lugares de tus {n} invitado(s).",
    errMustBeYes: "Primero confirma que juegas (Sí) para agregar equipo, invitados o un aviso de horario.",
    msgNoteSaved: "Aviso de horario guardado. Se puso en cola un aviso para el organizador.",
    msgNoteCleared: "🕒 Aviso de horario borrado.",
    errGuestsDisabled: "Los invitados están desactivados para tu cuenta.",
    errGuestDuplicateOwn: "Escribiste '{name}' dos veces.",
    errGuestMax: "Puedes traer hasta {max} invitados.",
    errGuestTaken: "'{name}' ya viene como invitado de {host}.",
    errGuestIsMember: "'{name}' ya es miembro de la liga y debe responder desde su perfil. Si es otra persona con el mismo nombre, pide ayuda a un admin.",
    errMemberAlreadyGuest: "{host} ya reservó un lugar para un invitado con tu nombre. Pídele que quite esa entrada antes de confirmar. Si es otra persona con el mismo nombre, pide al admin que aclare el nombre del invitado.",
    errNoShowBeforeKickoff: "Solo se pueden registrar ausencias después del inicio del partido.",
    errNoShowNotConfirmed: "Solo se puede registrar una ausencia de alguien que estaba confirmado.",
    errSeriesEffectivePast: "Elige hoy o una fecha posterior para cambiar la serie.",
    errGuestCapacity: "No hay lugar para {n} invitado(s): solo quedan {left} lugar(es) antes del cupo de {cap}.",
    msgGuestsSaved: "👥 {n} invitado(s) guardado(s): {names}",
    msgGuestsCleared: "👥 Invitados eliminados.",
    errBadGear: "Opción de equipo desconocida.",
    errGearOff: "La liga no está registrando ese artículo ahora.",
    errManagerOnly: "Solo el capitán de este partido o un admin pueden hacer eso.",
    errAfterKickoff: "El partido ya empezó.",
    errNudgeCooldown: "Ya se envió un recordatorio hace poco. Intenta en {mins} minuto(s).",
    errNudgeNobody: "¡Todos ya respondieron para este partido!",
    msgNudged: "Recordatorio en cola para {n} jugador(es). Se enviará en la próxima ejecución.",
    errNotCaptain: "No eres el capitán de este partido.",
    msgSteppedDown: "Dejaste la capitanía. Ahora los admins manejan este partido.",
    errSmallSidedMin: "El formato reducido necesita al menos {n} jugadores (5 contra 5).",
    msgSmallSided: "🟢 {format} aprobado — hay partido con {n} jugadores.",
    errAttendanceFuture: "La asistencia se abre el día del partido.",
    errAttendanceArgs: "Elige un jugador y Asistió o No vino.",
    msgAttended: "✅ {name} marcado como presente.",
    msgNoShow: "❌ {name} marcado como ausente.",
    errMotmPick: "Elige un jugador para votar.",
    errMotmSelf: "No puedes votar por ti mismo.",
    errMotmLocked: "La votación se abre al terminar el horario del partido o al guardarse el marcador final.",
    errMotmVoter: "Solo pueden votar los que jugaron.",
    errMotmNominee: "Ese jugador no estuvo en este partido.",
    msgMotmSaved: "⭐ ¡Voto guardado!",
    msgReconfirmedIn: "👍 Gracias, sigues confirmado con el nuevo horario.",
    msgReconfirmedMaybe: "👍 Entendido, sigues en tal vez.",
    errTeamsMin: "Se necesitan al menos 2 jugadores confirmados para armar equipos.",
    msgTeamsDefaults: "👕 Equipos fijos aplicados: {a} {na} – {nb} {b}.",
    msgTeamsSmart: "👕 Equipos equilibrados: {a} {na} – {nb} {b}.",
    errNoTeams: "Todavía no hay equipos — primero equilíbralos.",
    errNotInLineup: "Ese jugador no está en los equipos.",
    msgMoved: "🔄 {name} pasó a {team}.",
    errSwapPick: "Elige un jugador de cada equipo.",
    msgSwapped: "🔄 Se cambiaron {a} y {b}.",
    errScoreBeforeKickoff: "Puedes registrar el marcador cuando el partido ya haya empezado.",
    msgScoreCleared: "Marcador borrado — las confirmaciones están abiertas otra vez.",
    errScoreValue: "Los goles deben ser números enteros del 0 al 99.",
    msgScoreSaved: "🏆 Marcador final guardado: {a} {sa} – {sb} {b}.",
    errProtectedGuests: "Los admins protegidos siempre pueden traer invitados.",
    msgGuestsOn: "✅ Invitados activados para {email}.",
    msgGuestsOff: "🚫 Invitados desactivados para {email}.",
    msgDefaultTeamSet: "Ahora el equipo fijo de {name} es {team}.",
    msgDefaultTeamCleared: "{name} ya no tiene equipo fijo.",
    errTeamNameEmpty: "Los nombres de equipo no pueden quedar vacíos.",
    errTeamNamesSame: "Los dos equipos necesitan nombres distintos.",
    errShirt: "Elige un color de camiseta de la lista.",
    errMinMax: "Mínimo y máximo deben ser números enteros, mínimo 2, máximo 60, y mín ≤ máx.",
    msgSettingsSaved: "✅ Ajustes de la liga guardados.",
    msgRulesSaved: "📖 Reglas de la liga actualizadas.",
    errBlastEmpty: "Primero escribe un mensaje.",
    errDirectTarget: "No puedes enviar un mensaje a esa cuenta.",
    msgDirectSent: "Mensaje en cola para {name}. Revisa el estado en Admin > Ajustes.",
    msgBlastSent: "Aviso publicado. Notificaciones en cola para {n} jugador(es).",
    errNotFound: "No se encontró — tal vez lo eliminaron.",
    msgBlastHidden: "Mensaje oculto para los jugadores (los admins lo siguen viendo).",
    msgBlastShown: "El mensaje vuelve a estar visible.",
    errSeriesDates: "Elige una fecha de inicio y una fecha final posterior.",
    errSeriesDays: "Elige al menos un día de la semana.",
    errTime: "Elige una hora de inicio válida.",
    errLocation: "Elige una cancha de la lista (agrégala primero en Canchas).",
    msgSeriesUpdated: "Serie actualizada desde {date}: {u} actualizados, {c} nuevos, {r} cancelados.",
    msgSeriesCreated: "Serie guardada — los próximos {c} partido(s) ya están abiertos. Los demás se abren solos, con 4 semanas de anticipación.",
    msgSeriesStopped: "Series '{title}' stops from {date}: {n} upcoming match(es) cancelled; notices queued for {p} player(s).",
    errTitle: "Ponle un nombre al partido.",
    errDate: "Elige una fecha válida.",
    errCaptain: "Ese capitán no es un jugador activo.",
    errCapBelowRoster: "El máximo no puede ser menor que los {n} ya confirmados.",
    msgGameUpdatedNotified: "Partido actualizado. Avisos para reconfirmar en cola para {n} jugador(es).",
    msgGameUpdated: "Partido actualizado.",
    msgGameCreated: "¡Nuevo partido publicado!",
    errAlreadyCancelled: "Este partido ya está cancelado.",
    msgCancelled: "Partido cancelado. Avisos en cola para {n} jugador(es); revisa el estado en Admin > Ajustes.",
    errNotCancelled: "Este partido no está cancelado.",
    msgRestored: "Partido restaurado. Avisos en cola para {n} jugador(es).",
    errDeleteNeedsCancel: "Primero cancela el partido — los partidos con respuestas no se pueden borrar directamente.",
    msgDeleted: "Partido eliminado definitivamente.",
    errLocationFields: "Escribe el nombre de la cancha y una dirección.",
    errLatLng: "La latitud y la longitud deben ser números.",
    errLocationDup: "Ya existe una cancha con ese nombre.",
    msgLocationUpdated: "Cancha actualizada.",
    msgLocationCreated: "Cancha agregada.",
    msgLocationOn: "Cancha reactivada.",
    msgLocationOff: "Cancha desactivada.",
    msgLocationArchived: "Esta cancha tiene historial, así que se desactivó en lugar de borrarse.",
    msgLocationDeleted: "Cancha sin uso eliminada.",
    errEmail: "Escribe un correo válido.",
    errAlreadyMember: "{email} ya es miembro de la liga.",
    msgInviteNoQuota: "Invitación guardada para {email}, pero ya se usó el límite de correos de hoy — toca Reenviar mañana.",
    msgInviteMailFailed: "Invitación guardada para {email}, pero falló el correo: {err}",
    msgInviteSent: "Invitación en cola para {email}. Se enviará en la próxima ejecución.",
    errProtectedInvite: "Las invitaciones de admins protegidos no se pueden revocar.",
    msgInviteRevoked: "Invitación revocada.",
    errProtectedDemote: "Los admins protegidos no se pueden degradar.",
    msgMadeAdmin: "{email} ahora es administrador.",
    msgMadePlayer: "{email} ahora es jugador.",
    errProtectedSuspend: "Los admins protegidos no se pueden suspender.",
    msgSuspended: "{email} suspendido — se quitó de los próximos partidos y se cerró su sesión.",
    msgReactivated: "{email} reactivado.",
    errProtectedRemove: "Los admins protegidos no se pueden eliminar.",
    msgUserRemoved: "Se eliminó a {email} (sin historial de partidos)."
  }
};

// ============================================================================
// EMAILS IN EACH PLAYER'S LANGUAGE
// Every notification email is built by a function that takes "en" or "es" and returns
// { subject, html }. Each player gets it in the language they use in the app (Profiles → Lang).
// ============================================================================
var MAIL_DAYS = { en: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"], es: ["dom", "lun", "mar", "mié", "jue", "vie", "sáb"] };
var MAIL_MONTHS = { en: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], es: ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"] };
var SHIRT_WORDS = {
  en: { white: "white", black: "black or dark", gray: "gray", red: "red", blue: "blue", navy: "navy", green: "green", yellow: "yellow", orange: "orange", purple: "purple" },
  es: { white: "blanca", black: "negra u oscura", gray: "gris", red: "roja", blue: "azul", navy: "azul marino", green: "verde", yellow: "amarilla", orange: "naranja", purple: "morada" }
};
var TIMING_WORDS = {
  "Running ~5m Late": { en: "~5 min late", es: "~5 min tarde" },
  "Running ~10m Late": { en: "~10 min late", es: "~10 min tarde" },
  "Running 15m+ Late": { en: "15+ min late", es: "15+ min tarde" },
  "Leaving ~15m Early": { en: "Leaving ~15 min early", es: "Sale ~15 min antes" },
  "Leaving ~30m Early": { en: "Leaving ~30 min early", es: "Sale ~30 min antes" },
  "Leaving Early @ 12:30 PM": { en: "Leaving at 12:30 PM", es: "Sale a las 12:30 PM" },
  "Leaving Early @ 12:45 PM": { en: "Leaving at 12:45 PM", es: "Sale a las 12:45 PM" }
};

function mailLang(l) { return l === "es" ? "es" : "en"; }
function firstName(name) { return String(name || "").trim().split(/\s+/)[0] || String(name || ""); }
function esc(s) { return escapeHtmlServer(s === undefined || s === null ? "" : String(s)); }

function mailDate(iso, lang) {
  var p = String(iso || "").split("-");
  if (p.length !== 3) return String(iso || "");
  var d = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 12);
  if (mailLang(lang) === "es") return MAIL_DAYS.es[d.getDay()] + " " + d.getDate() + " " + MAIL_MONTHS.es[d.getMonth()];
  return MAIL_DAYS.en[d.getDay()] + ", " + MAIL_MONTHS.en[d.getMonth()] + " " + d.getDate();
}
function mailWhen(iso, time, lang) { return mailDate(iso, lang) + " · " + String(time || ""); }
function shirtWord(key, lang) { var m = SHIRT_WORDS[mailLang(lang)]; return m[key] || String(key || ""); }
function timingWord(note, lang) { var w = TIMING_WORDS[String(note || "")]; return w ? w[mailLang(lang)] : String(note || ""); }
function teamNameLang(side, lang) {
  var s = getLeagueSettings();
  var n = side === "dark" ? s.teamBName : s.teamAName;
  return mailLang(lang) === "es" && DEFAULT_TEAM_NAMES_ES.hasOwnProperty(n) ? DEFAULT_TEAM_NAMES_ES[n] : n;
}
function statusWordMail(status, lang) {
  var es = mailLang(lang) === "es";
  return ({ Yes: es ? "Voy" : "In", Maybe: es ? "Tal vez" : "Maybe", No: es ? "No voy" : "Out", Waitlist: es ? "En espera" : "Waitlist" })[status] || String(status || "");
}
function mapsUrl(addressOrName) { return "https://www.google.com/maps/search/?api=1&query=" + encodeURIComponent(addressOrName || ""); }

// One layout for every email: heading, paragraphs, a highlighted box, more paragraphs, buttons, footnote.
function mailLayout(o) {
  var accent = o.accent || "#00f59b";
  var p = function(html) { return html ? "<p style='margin:0 0 12px;'>" + html + "</p>" : ""; };
  var h = "<div style='font-family:-apple-system,BlinkMacSystemFont,\"Segoe UI\",Roboto,Helvetica,Arial,sans-serif;padding:24px;background:#0d1612;color:#ffffff;border-radius:12px;max-width:600px;line-height:1.5;font-size:15px;'>";
  h += "<h2 style='color:" + accent + ";margin:0 0 14px;font-size:20px;line-height:1.3;'>" + o.heading + "</h2>";
  (o.paras || []).forEach(function(x) { h += p(x); });
  if (o.box) h += "<div style='background:#16281f;padding:16px;border-left:4px solid " + (o.boxAccent || accent) + ";border-radius:8px;margin:16px 0;" + (o.boxPre ? "white-space:pre-wrap;" : "") + "'>" + o.box + "</div>";
  (o.after || []).forEach(function(x) { h += p(x); });
  if (o.buttons && o.buttons.length) {
    h += "<div style='margin:20px 0 6px;'>";
    o.buttons.forEach(function(b) {
      h += "<a href='" + b.url + "' style='display:inline-block;margin:0 10px 10px 0;padding:12px 22px;background:" + (b.bg || "#00f59b") + ";color:" + (b.fg || "#070b09") + ";text-decoration:none;font-weight:bold;border-radius:8px;'>" + b.label + "</a>";
    });
    h += "</div>";
  }
  if (o.foot) h += "<p style='font-size:12px;color:#94a3b8;margin:14px 0 0;'>" + o.foot + "</p>";
  return h + "</div>";
}
function mailBtn(label, url, bg, fg) { return { label: label, url: url || SITE_URL, bg: bg, fg: fg }; }
function hubBtn(lang, en, es) { return mailBtn(mailLang(lang) === "es" ? (es || "Abrir Pitch Pulse →") : (en || "Open Pitch Pulse →"), SITE_URL); }
function rsvpButtons(lang, playerEmail, gameId) {
  var es = mailLang(lang) === "es";
  return [
    mailBtn(es ? "✅ Voy" : "✅ I'm in", buildMagicRsvpUrl(playerEmail, gameId, "Yes")),
    mailBtn(es ? "🤔 Tal vez" : "🤔 Maybe", buildMagicRsvpUrl(playerEmail, gameId, "Maybe"), "#fbbf24"),
    mailBtn(es ? "❌ No voy" : "❌ Can't make it", buildMagicRsvpUrl(playerEmail, gameId, "No"), "#f43f5e", "#ffffff")
  ];
}
function rosterNames(list, lang, withExtras) {
  return list.map(function(p) {
    var s = esc(p.fullName);
    if (withExtras && (p.playPosition || p.position)) s += " [" + esc(p.playPosition || p.position) + "]";
    if (p.guestNames && p.guestNames.length) s += " (+" + p.guestNames.length + ": " + esc(p.guestNames.join(", ")) + ")";
    if (withExtras && p.availabilityNote) s += " [" + esc(timingWord(p.availabilityNote, lang)) + "]";
    return s;
  }).join(", ");
}

// Sends one notification: email built per language + push in each player's language.
// push = { en: [title, body], es: [title, body] }
function notifyMail(recipients, mailFn, push, opts) {
  opts = Object.assign({}, mailFn && mailFn.notificationOptions || {}, opts || {});
  opts.mail = mailFn || null;
  push = push || {};
  if (push.es) { opts.pushTitleEs = push.es[0]; opts.pushBodyEs = push.es[1]; }
  var en = mailFn ? mailFn("en") : { subject: push.en ? push.en[0] : "", html: "" };
  return dispatchLeagueNotification(recipients, en.subject, en.html, push.en ? push.en[0] : en.subject, push.en ? push.en[1] : en.subject, opts.url || SITE_URL, opts);
}

// ---- Weather (fetched once, written in each reader's language) ------------------
function fetchBackendWeather(lat, lng, dateStr, timeStr) {
  if (!lat || !lng) return null;
  try {
    var kickoffHour = Math.floor(parseTimeMinutes(timeStr || "12:00 PM") / 60);
    var hourIsoPrefix = dateStr + "T" + (kickoffHour < 10 ? "0" + kickoffHour : String(kickoffHour)) + ":00";
    var url = "https://api.open-meteo.com/v1/forecast?latitude=" + lat + "&longitude=" + lng +
      "&hourly=temperature_2m,precipitation_probability,weathercode,windspeed_10m" +
      "&daily=temperature_2m_max,temperature_2m_min,precipitation_probability_max,weathercode" +
      "&temperature_unit=fahrenheit&windspeed_unit=mph&timezone=auto";
    var json = JSON.parse(UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText());
    if (json && json.hourly && json.hourly.time) {
      var h = json.hourly.time.indexOf(hourIsoPrefix);
      if (h !== -1) return { hourly: true, code: json.hourly.weathercode[h], temp: Math.round(json.hourly.temperature_2m[h]), rain: json.hourly.precipitation_probability[h] || 0, wind: Math.round(json.hourly.windspeed_10m[h] || 0) };
    }
    if (json && json.daily && json.daily.time) {
      var d = json.daily.time.indexOf(dateStr);
      if (d !== -1) return { hourly: false, code: json.daily.weathercode[d], hi: Math.round(json.daily.temperature_2m_max[d]), lo: Math.round(json.daily.temperature_2m_min[d]), rain: json.daily.precipitation_probability_max[d] || 0 };
    }
  } catch (e) {}
  return null;
}
function weatherText(w, lang) {
  var es = mailLang(lang) === "es";
  if (!w) return es ? "Pronóstico en la app" : "Forecast in the app";
  var c = w.code;
  var cond = c === 0 ? (es ? "☀️ Despejado" : "☀️ Clear") : (c < 3 ? (es ? "⛅ Parcialmente nublado" : "⛅ Partly cloudy") : (c < 50 ? (es ? "☁️ Nublado" : "☁️ Overcast") : (es ? "🌧️ Probable lluvia" : "🌧️ Rain likely")));
  if (w.hourly) {
    var s = cond + ", " + w.temp + "°F " + (es ? "al inicio" : "at kickoff") + " (" + w.rain + "% " + (es ? "de lluvia" : "rain") + ", " + (es ? "viento " : "wind ") + w.wind + " mph)";
    if (w.temp >= 85) s += es ? " 🔥 Calor: trae agua extra" : " 🔥 Heat: bring extra water";
    return s;
  }
  return cond + ", " + (es ? "máx " : "high ") + w.hi + "°F / " + (es ? "mín " : "low ") + w.lo + "°F (" + w.rain + "% " + (es ? "de lluvia" : "rain") + ")";
}

// ---- Builders ---------------------------------------------------------------------
// g = { title, date, time, location } (gameRecord shape)

function mailWaitlistPromoted(g) {
  return notificationMail(function(lang) {
    var es = mailLang(lang) === "es", when = mailWhen(g.date, g.time, lang);
    return {
      subject: es ? "🎉 ¡Entraste! Se liberó un lugar en " + g.title : "🎉 You're in! A spot opened in " + g.title,
      html: mailLayout({
        heading: es ? "🎉 Saliste de la lista de espera: ¡vas!" : "🎉 You're off the waitlist — you're in!",
        paras: [es ? "Se liberó un lugar en <b>" + esc(g.title) + "</b> (" + esc(when) + ") y subiste automáticamente. Ya estás confirmado."
                   : "A spot opened in <b>" + esc(g.title) + "</b> (" + esc(when) + ") and you moved up automatically. You're confirmed."],
        box: "📍 " + esc(g.location),
        after: [es ? "¿Ya no puedes ir? Cambia a <b>No voy</b> en la app para que el siguiente tome el lugar."
                   : "Can't make it anymore? Switch to <b>Out</b> in the app so the next person gets the spot."],
        buttons: [hubBtn(lang, "View the match →", "Ver el partido →")]
      })
    };
  }, g, { critical: true, expectedSignupStatus: "Yes", timeSensitive: true });
}

function mailGearAlert(o) { // { playerName, status, g, ball, pinnies, gk }
  return notificationMail(function(lang) {
    var es = mailLang(lang) === "es", g = o.g, when = mailWhen(g.date, g.time, lang);
    var items = [];
    if (o.ball) items.push(es ? "⚽ Balón" : "⚽ Match ball");
    if (o.pinnies) items.push(es ? "🦺 Chalecos" : "🦺 Pinnies");
    if (o.gk) items.push(es ? "🧤 Portero" : "🧤 Goalkeeper");
    return {
      subject: es ? "⚠️ Alerta: " + o.playerName + " se bajó de " + g.title : "⚠️ Gear alert: " + o.playerName + " dropped out of " + g.title,
      html: mailLayout({
        accent: "#f87171",
        heading: es ? "⚠️ Ya nadie cubre esto" : "⚠️ Nobody is covering this anymore",
        paras: [es ? "<b>" + esc(o.playerName) + "</b> cambió a <b>" + esc(statusWordMail(o.status, lang)) + "</b> en <b>" + esc(g.title) + "</b> (" + esc(when) + "). El partido se queda sin:"
                   : "<b>" + esc(o.playerName) + "</b> switched to <b>" + esc(statusWordMail(o.status, lang)) + "</b> for <b>" + esc(g.title) + "</b> (" + esc(when) + "). That leaves the match without:"],
        box: "<b style='color:#f87171;font-size:16px;'>" + items.join("<br>") + "</b>",
        after: [es ? "Manda un aviso desde <b>⚡ Herramientas</b> para pedir que alguien lo cubra." : "Send a quick announcement from <b>⚡ Organizer tools</b> to ask someone to cover."],
        buttons: [hubBtn(lang)]
      })
    };
  }, o.g, { critical: true, timeSensitive: true });
}

function mailRsvpConfirmed(o) { // { g, yes, min, max }
  return notificationMail(function(lang) {
    var es = mailLang(lang) === "es", g = o.g, when = mailWhen(g.date, g.time, lang);
    return {
      subject: es ? "✅ Confirmado: " + g.title + " (" + when + ")" : "✅ You're in: " + g.title + " (" + when + ")",
      html: mailLayout({
        heading: es ? "✅ Estás en la lista" : "✅ You're on the roster",
        paras: [es ? "Estás marcado como <b>Voy</b> en <b>" + esc(g.title) + "</b>." : "You're marked <b>In</b> for <b>" + esc(g.title) + "</b>."],
        box: "🗓️ " + esc(when) + "<br>📍 " + esc(g.location) + "<br>📊 " + o.yes + "/" + o.max + (es ? " confirmados (mínimo " + o.min + ")" : " players in (minimum " + o.min + ")"),
        buttons: [hubBtn(lang, "See the roster & your shirt →", "Ver la lista y tu camiseta →")]
      })
    };
  }, o.g, { priority: 10, timeSensitive: true, expectedSignupStatus: "Yes", expiresAtMs: new Date().getTime() + 15 * 60000 });
}

function mailTimingAlert(o) { // { playerName, playerEmail, note, g }
  return notificationMail(function(lang) {
    var es = mailLang(lang) === "es", g = o.g, when = mailWhen(g.date, g.time, lang), note = timingWord(o.note, lang);
    return {
      subject: "🕒 " + o.playerName + ": " + note + " — " + g.title,
      html: mailLayout({
        accent: "#fbbf24",
        heading: es ? "🕒 Aviso de llegada / salida" : "🕒 Arrival / departure update",
        paras: [es ? "<b>" + esc(o.playerName) + "</b> (" + esc(o.playerEmail) + ") avisó sobre <b>" + esc(g.title) + "</b> (" + esc(when) + "):"
                   : "<b>" + esc(o.playerName) + "</b> (" + esc(o.playerEmail) + ") posted an update for <b>" + esc(g.title) + "</b> (" + esc(when) + "):"],
        box: "<b style='color:#fbbf24;font-size:16px;'>" + esc(note) + "</b>",
        after: [es ? "Se ve junto a su nombre en la lista del partido." : "It shows next to their name on the match roster."],
        buttons: [hubBtn(lang)]
      })
    };
  }, o.g, { priority: 50 });
}

function mailNudge(o) { // { playerEmail, playerName, managerName, game (appData game), weather }
  return notificationMail(function(lang) {
    var es = mailLang(lang) === "es", game = o.game, when = mailWhen(game.Date, game.Time, lang);
    var committed = game.confirmedCount === undefined ? Number(game.yesCount || 0) : Number(game.confirmedCount);
    var need = Math.max(0, game.MinThreshold - committed), left = Math.max(0, game.MaxCap - game.yesCount);
    var countLine = notificationCountSummary(game, lang) + " — " +
      (need > 0 ? "<b style='color:#fbbf24;'>" + (es ? "faltan " + need + " para que haya partido" : "only " + need + " more needed for the game to be on") + "</b>"
                : "<b style='color:#00f59b;'>" + (es ? "hay partido, quedan " + left + " lugares" : "game on, " + left + " spot(s) left") + "</b>");
    return {
      subject: es ? "⚽ ¿Juegas en " + game.Title + " (" + when + ")?" : "⚽ Are you in for " + game.Title + " (" + when + ")?",
      html: mailLayout({
        heading: es ? "Hola " + esc(firstName(o.playerName)) + ", ¿juegas?" : "Hi " + esc(firstName(o.playerName)) + " — are you playing?",
        paras: [es ? "<b>" + esc(o.managerName) + "</b> está contando jugadores para <b>" + esc(game.Title) + "</b> (" + esc(when) + ")."
                   : "<b>" + esc(o.managerName) + "</b> is checking numbers for <b>" + esc(game.Title) + "</b> (" + esc(when) + ")."],
        box: "📍 " + esc(game.LocationName) + "<br>📊 " + countLine + (o.showWeather === false ? "" : "<br>" + esc(weatherText(o.weather, lang))),
        after: [es ? "<b>Toca tu respuesta:</b>" : "<b>Tap your answer:</b>"],
        buttons: rsvpButtons(lang, o.playerEmail, game.GameID),
        foot: "<a href='" + SITE_URL + "' style='color:#00f59b;'>" + (es ? "O abre Pitch Pulse →" : "Or open Pitch Pulse →") + "</a>"
      })
    };
  }, o.game, { priority: 30, timeSensitive: true, skipAway: true, requireNoAnswer: true });
}

function mailTomorrow(o) { // { playerEmail, playerName, game, weather, spotsLeft }
  return notificationMail(function(lang) {
    var es = mailLang(lang) === "es", game = o.game, when = mailWhen(game.Date, game.Time, lang);
    return {
      subject: es ? "⚽ ¿Juegas mañana? " + game.Title : "⚽ Playing tomorrow? " + game.Title,
      html: mailLayout({
        heading: es ? "Hola " + esc(firstName(o.playerName)) + ", ¿juegas mañana?" : "Hi " + esc(firstName(o.playerName)) + " — are you in for tomorrow?",
        paras: [es ? "Hay un partido programado <b>mañana (" + esc(when) + ")</b> y todavía no respondiste." : "There's a match <b>tomorrow (" + esc(when) + ")</b> and we haven't seen your answer yet."],
        box: "🏆 " + esc(game.Title) + "<br>📍 " + esc(game.LocationName) + "<br>📊 " + notificationCountSummary(game, lang) + (es ? " · quedan " + o.spotsLeft + " lugares" : " · " + o.spotsLeft + " spot(s) left") + (o.showWeather === false ? "" : "<br>" + esc(weatherText(o.weather, lang))),
        after: [es ? "<b>Toca tu respuesta:</b>" : "<b>Tap your answer:</b>"],
        buttons: rsvpButtons(lang, o.playerEmail, game.GameID),
        foot: "<a href='" + SITE_URL + "' style='color:#00f59b;'>" + (es ? "O abre Pitch Pulse" : "Or open Pitch Pulse") + "</a>" +
          (es ? " · ¿Vas a estar fuera? Usa «Fuera hasta» en tu perfil." : " · Away for a while? Set “Away until” in your profile.")
      })
    };
  }, o.game, { priority: 20, timeSensitive: true, skipAway: true, requireNoAnswer: true });
}

function mailMaybeMorning(o) { // { playerEmail, playerName, game (appData game), spotsLeft }
  return notificationMail(function(lang) {
    var es = mailLang(lang) === "es", game = o.game;
    var need = Math.max(0, Number(game.MinThreshold) - Number(game.confirmedCount === undefined ? game.yesCount : game.confirmedCount));
    var countLine = notificationCountSummary(game, lang) + (es ? " · quedan " + o.spotsLeft + " lugar(es)" : " · " + o.spotsLeft + " spot(s) left") +
      (need > 0 ? (es ? " · faltan " + need + " para que haya partido" : " · " + need + " more needed for the game to be on") : "");
    return {
      subject: es ? "🤔 ¿Sigues en tal vez para hoy? " + game.Title : "🤔 Still a maybe for today? " + game.Title,
      html: mailLayout({
        accent: "#fbbf24",
        heading: es ? "Hola " + esc(firstName(o.playerName)) + ", ¿juegas hoy?" : "Hi " + esc(firstName(o.playerName)) + ", are you playing today?",
        paras: [es ? "Estás en <b>tal vez</b> para <b>" + esc(game.Title) + "</b> hoy a las " + esc(game.Time) + (o.spotsLeft > 0 ? ". Todavía hay lugar, y es para quien responda primero." : ". Está lleno: puedes entrar a la lista de espera.")
                   : "You're down as <b>Maybe</b> for <b>" + esc(game.Title) + "</b> today at " + esc(game.Time) + (o.spotsLeft > 0 ? ". There's still room, and spots go to whoever answers first." : ". It is full: you can join the waitlist.")],
        box: "📍 " + esc(game.LocationName) + "<br>📊 " + countLine,
        after: [es ? "<b>Toca tu respuesta:</b>" : "<b>Tap your answer:</b>"],
        buttons: [
          mailBtn(o.spotsLeft > 0 ? (es ? "✅ Voy" : "✅ I'm in") : (es ? "⏳ A la espera" : "⏳ Join waitlist"), buildMagicRsvpUrl(o.playerEmail, game.GameID, "Yes")),
          mailBtn(es ? "❌ No voy" : "❌ Can't make it", buildMagicRsvpUrl(o.playerEmail, game.GameID, "No"), "#f43f5e", "#ffffff")
        ],
        foot: "<a href='" + SITE_URL + "' style='color:#00f59b;'>" + (es ? "O abre Pitch Pulse →" : "Or open Pitch Pulse →") + "</a>"
      })
    };
  }, o.game, { priority: 30, timeSensitive: true, skipAway: true, expectedSignupStatus: "Maybe" });
}

function mailCaptainStepDown(o) { // { playerName, title }
  return function(lang) {
    var es = mailLang(lang) === "es";
    return {
      subject: es ? "👑 " + o.playerName + " dejó la capitanía de " + o.title : "👑 " + o.playerName + " stepped down as captain of " + o.title,
      html: mailLayout({
        accent: "#fbbf24",
        heading: es ? "👑 Cambio de capitán" : "👑 Captain stepped down",
        paras: [es ? "<b>" + esc(o.playerName) + "</b> ya no es capitán de <b>" + esc(o.title) + "</b>. Ahora el partido lo llevan los admins de la liga."
                   : "<b>" + esc(o.playerName) + "</b> is no longer captain of <b>" + esc(o.title) + "</b>. The league admins now run this match."],
        buttons: [hubBtn(lang)]
      })
    };
  };
}

function mailCaptainAssigned(o) { // { title, date, time, location }
  return function(lang) {
    var es = mailLang(lang) === "es", when = mailWhen(o.date, o.time, lang);
    return {
      subject: es ? "👑 Eres capitán de " + o.title + " (" + when + ")" : "👑 You're captain of " + o.title + " (" + when + ")",
      html: mailLayout({
        heading: es ? "👑 Eres el capitán del partido" : "👑 You're the match captain",
        paras: [es ? "Eres capitán de <b>" + esc(o.title) + "</b> el <b>" + esc(when) + "</b> (" + esc(o.location) + ")."
                   : "You're captain of <b>" + esc(o.title) + "</b> on <b>" + esc(when) + "</b> (" + esc(o.location) + ").",
                es ? "En la tarjeta del partido toca <b>⚡ Herramientas del organizador</b> para equilibrar equipos, mandar recordatorios y avisos, pasar asistencia y anotar el marcador."
                   : "On the match card, tap <b>⚡ Organizer tools</b> to balance teams, nudge players, post announcements, take attendance and record the score."],
        buttons: [hubBtn(lang)]
      })
    };
  };
}

function mailNoShow(o) { // { playerName, managerName, g }
  return notificationMail(function(lang) {
    var es = mailLang(lang) === "es", g = o.g, when = mailWhen(g.date, g.time, lang);
    return {
      subject: es ? "📋 Asistencia: te marcaron ausente en " + g.title : "📋 Attendance: you were marked absent for " + g.title,
      html: mailLayout({
        heading: es ? "📋 Actualización de asistencia" : "📋 Attendance update",
        paras: [es ? "Hola " + esc(firstName(o.playerName)) + ":" : "Hi " + esc(firstName(o.playerName)) + ",",
                es ? "<b>" + esc(o.managerName) + "</b> te marcó como <b style='color:#f87171;'>ausente</b> en <b>" + esc(g.title) + "</b> (" + esc(when) + ")."
                   : "<b>" + esc(o.managerName) + "</b> marked you as a <b style='color:#f87171;'>no-show</b> for <b>" + esc(g.title) + "</b> (" + esc(when) + ").",
                es ? "Si es un error, avísale: lo puede corregir en la app." : "If that's a mistake, let them know — they can fix it in the app."],
        buttons: [hubBtn(lang)]
      })
    };
  }, o.g, { priority: 50 });
}

function mailStatusFlip(o) { // { game (appData game), confirmed, yesCount }
  return notificationMail(function(lang) {
    var es = mailLang(lang) === "es", game = o.game, when = mailWhen(game.Date, game.Time, lang);
    var yes = game.roster.filter(function(r) { return r.status === "Yes"; });
    return {
      subject: o.confirmed ? (es ? "🟢 ¡Hay partido! " + game.Title + " (" + when + ")" : "🟢 Game on! " + game.Title + " (" + when + ")")
                           : (es ? "⚠️ " + game.Title + " (" + when + ") bajó del mínimo" : "⚠️ " + game.Title + " (" + when + ") dropped below the minimum"),
      html: mailLayout({
        accent: o.confirmed ? "#00f59b" : "#fbbf24",
        heading: o.confirmed ? (es ? "🟢 ¡Hay partido! Ya está confirmado" : "🟢 Game on — the match is confirmed")
                             : (es ? "⚠️ Otra vez pendiente: faltan jugadores" : "⚠️ Back to pending — not enough players"),
        paras: ["<b>" + esc(game.Title) + "</b> · " + esc(when) + " · " + esc(game.LocationName)],
        box: "<b>" + (es ? "Van (" : "In (") + o.yesCount + "/" + game.MaxCap + (es ? ", mínimo " : ", minimum ") + game.MinThreshold + "):</b><br>" + (yes.length ? rosterNames(yes, lang) : (es ? "Nadie" : "Nobody")),
        after: [o.confirmed ? "" : (es ? "Consejo: usa <b>Recordatorio</b> o <b>Aviso</b> en Herramientas para sumar jugadores." : "Tip: use <b>Nudge</b> or <b>Announcement</b> in Organizer tools to bring in a few more players.")],
        buttons: [hubBtn(lang)]
      })
    };
  }, o.game, { priority: o.confirmed ? 40 : 100, timeSensitive: true, expectedGameStatus: o.confirmed ? "Confirmed" : "Pending", supersedeKey: "game-status:" + o.game.GameID });
}

function mailBlast(o) { // { senderName, message, g }
  return notificationMail(function(lang) {
    var es = mailLang(lang) === "es", g = o.g, when = mailWhen(g.date, g.time, lang);
    return {
      subject: es ? "📢 " + g.title + " (" + when + "): aviso de " + o.senderName : "📢 " + g.title + " (" + when + "): message from " + o.senderName,
      html: mailLayout({
        heading: es ? "📢 Aviso de " + esc(o.senderName) : "📢 Announcement from " + esc(o.senderName),
        paras: [(es ? "Sobre <b>" : "About <b>") + esc(g.title) + "</b> · " + esc(when) + " · " + esc(g.location) + ":"],
        box: esc(o.message), boxPre: true,
        buttons: [hubBtn(lang, "View the match →", "Ver el partido →")]
      })
    };
  }, o.g, { critical: true });
}

function mailDirectMessage(o) { // { senderName, message }
  return function(lang) {
    var es = mailLang(lang) === "es", who = esc(o.senderName);
    return {
      subject: es ? "💬 Mensaje de " + o.senderName + " (Pitch Pulse)" : "💬 Message from " + o.senderName + " (Pitch Pulse)",
      html: mailLayout({
        heading: es ? "💬 Mensaje de " + who : "💬 Message from " + who,
        box: esc(o.message), boxPre: true,
        after: [es ? "Si respondes a este correo, le llega a " + esc(firstName(o.senderName)) + "." : "Reply to this email and it goes straight to " + esc(firstName(o.senderName)) + "."],
        buttons: [hubBtn(lang)]
      })
    };
  };
}

function mailSeriesUpdated(o) { // { title, effectiveDate, endDate, time, location }
  return function(lang) {
    var es = mailLang(lang) === "es";
    var cancelled = (o.cancelledGames || []).map(function(g) { return "<li>" + esc(mailWhen(g.date, g.time, lang)) + " · " + esc(g.title) + " · " + esc(g.location) + "</li>"; }).join("");
    return {
      subject: es ? "⚠️ Cambio de horario en " + o.title + " (desde el " + mailDate(o.effectiveDate, lang) + ")" : "⚠️ Schedule change for " + o.title + " (from " + mailDate(o.effectiveDate, lang) + ")",
      html: mailLayout({
        accent: "#fbbf24",
        heading: es ? "⚠️ Cambió el horario semanal" : "⚠️ The weekly schedule changed",
        paras: [es ? "Un admin actualizó <b>" + esc(o.title) + "</b> a partir del <b>" + esc(mailDate(o.effectiveDate, lang)) + "</b>."
                   : "An admin updated <b>" + esc(o.title) + "</b> starting <b>" + esc(mailDate(o.effectiveDate, lang)) + "</b>."],
        box: "📍 " + esc(o.location) + "<br>🕒 " + esc(o.time) + "<br>🗓️ " + esc(mailDate(o.effectiveDate, lang)) + " → " + esc(mailDate(o.endDate, lang)),
        after: [(cancelled ? "<b>" + (es ? "Partidos cancelados:" : "Cancelled matches:") + "</b><ul>" + cancelled + "</ul>" : ""), es ? "Habías respondido para próximos partidos de esta serie: revisa que tus respuestas sigan valiendo."
                   : "You had answered for upcoming matches in this series — please check your answers still work."],
        buttons: [hubBtn(lang, "Review my matches →", "Revisar mis partidos →")]
      })
    };
  };
}

function mailSeriesStopped(o) { // { title, effectiveDate, games: [{title, date, time, location}] }
  return function(lang) {
    var es = mailLang(lang) === "es";
    var items = o.games.map(function(x) { return "<li>🗓️ <b>" + esc(x.title) + "</b> — " + esc(mailWhen(x.date, x.time, lang)) + " (" + esc(x.location) + ")</li>"; }).join("");
    return {
      subject: es ? "🚫 Cancelado: " + o.title + " (desde el " + mailDate(o.effectiveDate, lang) + ")" : "🚫 Cancelled: " + o.title + " (from " + mailDate(o.effectiveDate, lang) + ")",
      html: mailLayout({
        accent: "#f87171",
        heading: es ? "🚫 Se cancelaron próximos partidos" : "🚫 Upcoming matches cancelled",
        paras: [es ? "Un admin detuvo los partidos semanales de <b>" + esc(o.title) + "</b> desde el <b>" + esc(mailDate(o.effectiveDate, lang)) + "</b>."
                   : "An admin stopped the weekly <b>" + esc(o.title) + "</b> matches from <b>" + esc(mailDate(o.effectiveDate, lang)) + "</b>."],
        box: "<b style='color:#f87171;'>" + (es ? "Partidos en los que habías respondido:" : "Matches you had answered for:") + "</b><ul style='line-height:1.8;margin:8px 0 0;padding-left:20px;'>" + items + "</ul>",
        buttons: [hubBtn(lang, "See the schedule →", "Ver el calendario →")]
      })
    };
  };
}

function mailScheduleChange(o) { // { oldGame: {date,time,location}, newGame: {title,date,time,locationName}, address }
  return notificationMail(function(lang) {
    var es = mailLang(lang) === "es", a = o.oldGame, b = o.newGame;
    var li = function(icon, label, oldV, newV) {
      return oldV !== newV
        ? "<li>" + icon + " <b>" + label + ":</b> <span style='text-decoration:line-through;color:#f87171;'>" + esc(oldV) + "</span> → <b style='color:#00f59b;'>" + esc(newV) + "</b></li>"
        : "<li>" + icon + " <b>" + label + ":</b> " + esc(newV) + " (" + (es ? "sin cambio" : "unchanged") + ")</li>";
    };
    var list = li("🗓️", es ? "Fecha" : "Date", mailDate(a.date, lang), mailDate(b.date, lang)) + li("🕒", es ? "Hora" : "Kickoff", a.time, b.time) + li("📍", es ? "Cancha" : "Field", a.location, b.locationName);
    return {
      subject: es ? "⚠️ Cambio de horario: confirma de nuevo " + b.title : "⚠️ Schedule change: please re-confirm " + b.title,
      html: mailLayout({
        accent: "#fbbf24",
        heading: es ? "⚠️ Este partido cambió: confirma otra vez" : "⚠️ This match changed — please re-confirm",
        paras: [es ? "Un admin cambió <b>" + esc(b.title) + "</b>, en el que habías respondido." : "An admin changed <b>" + esc(b.title) + "</b>, which you had answered for."],
        box: "<ul style='line-height:1.8;margin:0;padding-left:20px;'>" + list + "</ul><p style='margin:10px 0 0;font-size:13px;'><a href='" + mapsUrl(o.address || b.locationName) + "' style='color:#00f59b;'>🗺️ " + (es ? "Ver en Google Maps" : "Open in Google Maps") + "</a></p>",
        after: [es ? "Tu lugar sigue reservado y marcado como <b>pendiente de confirmar</b>. Dinos si todavía puedes ir (o cambia a <b>No voy</b>)."
                   : "Your spot is held and marked <b>awaiting re-confirmation</b>. Let us know if you can still make it (or switch to <b>Out</b>)."],
        buttons: [hubBtn(lang, "Re-confirm my answer →", "Confirmar mi respuesta →")]
      })
    };
  }, o.newGame, { critical: true, supersedeKey: "game-update:" + o.newGame.id });
}

function mailCancelled(g) {
  return notificationMail(function(lang) {
    var es = mailLang(lang) === "es", when = mailWhen(g.date, g.time, lang);
    return {
      subject: es ? "🚫 Cancelado: " + g.title + " (" + when + ")" : "🚫 Cancelled: " + g.title + " (" + when + ")",
      html: mailLayout({
        accent: "#f87171",
        heading: es ? "🚫 Partido cancelado" : "🚫 Match cancelled",
        paras: [es ? "Un admin canceló un partido en el que habías respondido:" : "An admin cancelled a match you had answered for:"],
        box: "<span style='text-decoration:line-through;color:#f87171;font-weight:bold;'>" + esc(g.title) + "</span><br>🗓️ " + esc(when) + "<br>📍 " + esc(g.location),
        after: [es ? "Revisa la app para ver otros partidos." : "Check the app for other upcoming matches."],
        buttons: [hubBtn(lang)]
      })
    };
  }, g, { critical: true, supersedeKey: "game-update:" + g.id });
}

function mailRestored(g) {
  return notificationMail(function(lang) {
    var es = mailLang(lang) === "es", when = mailWhen(g.date, g.time, lang);
    return {
      subject: es ? "✅ Vuelve el partido: " + g.title + " (" + when + ")" : "✅ Back on: " + g.title + " (" + when + ")",
      html: mailLayout({
        heading: es ? "✅ El partido vuelve" : "✅ The match is back on",
        paras: [es ? "<b>" + esc(g.title) + "</b> del <b>" + esc(when) + "</b> (" + esc(g.location) + ") vuelve al calendario. Tu respuesta sigue activa; cámbiala en la app si tus planes cambiaron."
                   : "<b>" + esc(g.title) + "</b> on <b>" + esc(when) + "</b> (" + esc(g.location) + ") is back on the schedule. Your answer still stands — update it in the app if your plans changed."],
        buttons: [hubBtn(lang)]
      })
    };
  }, g, { critical: true, supersedeKey: "game-update:" + g.id });
}

function mailMatchday(o) { // { game (appData game), settings, weather }
  return notificationMail(function(lang) {
    var es = mailLang(lang) === "es", game = o.game, s = o.settings;
    var committed = game.confirmedCount === undefined ? Number(game.yesCount || 0) : Number(game.confirmedCount);
    var confirmed = committed >= Number(game.MinThreshold);
    var yes = game.roster.filter(function(r) { return r.status === "Yes" && !r.needsReconfirm; });
    var reconfirm = game.roster.filter(function(r) { return r.status === "Yes" && r.needsReconfirm; });
    var wait = game.roster.filter(function(r) { return r.status === "Waitlist"; });
    var maybe = game.roster.filter(function(r) { return r.status === "Maybe"; });
    var no = game.roster.filter(function(r) { return r.status === "No"; });
    var names = function(list) { return esc(list.map(function(p) { return p.fullName; }).join(", ")); };
    var need = es ? "⚠️ ¡Falta!" : "⚠️ Needed!";
    var gear = [];
    if (s.trackBall) { var b = yes.filter(function(p) { return p.bringBall; }); gear.push((es ? "Balón: " : "Ball: ") + (b.length ? names(b) : need)); }
    if (s.trackPinnies) { var pi = yes.filter(function(p) { return p.bringPinnies; }); gear.push((es ? "Chalecos: " : "Pinnies: ") + (pi.length ? names(pi) : need)); }
    if (s.trackGK) { var gk = yes.filter(function(p) { return p.playGK; }); gear.push((es ? "Portero: " : "GK: ") + (gk.length ? names(gk) : (es ? "libre" : "open"))); }
    var captainLine = game.captainName ? "👑 " + (es ? "Capitán: " : "Captain: ") + esc(game.captainName) : "🛡️ " + (es ? "A cargo de los admins de la liga" : "Run by the league admins");
    var format = String(game.MatchFormat || "Standard");
    if (format === "Standard") format = es ? "Estándar" : "Standard";
    var teams = "";
    if (game.teams && game.teams.white.length + game.teams.dark.length > 0) {
      var fmtTeam = function(list) { return list.map(function(x) { return esc(x.fullName) + (x.isGuest ? (es ? " (invitado)" : " (guest)") : " (" + esc(x.position) + ")"); }).join(", "); };
      var side = function(key, shirtKey, list) { return "<b>" + esc(teamNameLang(key, lang)) + " (" + (es ? "camiseta " : "") + esc(shirtWord(shirtKey, lang)) + (es ? "" : " shirts") + "):</b> " + fmtTeam(list); };
      teams = "<div style='background:#111c16;padding:14px;border-radius:10px;margin:14px 0;border:1px solid #00f59b;'><b style='color:#00f59b;'>👕 " + (es ? "Equipos" : "Lineups") + ":</b><br>" +
        side("white", s.teamAShirt, game.teams.white) + "<br>" + side("dark", s.teamBShirt, game.teams.dark) + "</div>";
    }
    var section = function(color, label, list, withExtras, emptyWord) {
      return "<h3 style='color:" + color + ";margin:16px 0 6px;font-size:16px;'>" + label + " (" + (list === yes ? committed : (list === reconfirm ? Number(game.pendingReconfirmCount || reconfirm.length) : list.length)) + ")</h3><p style='margin:0;'>" + (list.length ? (withExtras ? rosterNames(list, lang, true) : names(list)) : emptyWord) + "</p>";
    };
    var html = mailLayout({
      accent: confirmed ? "#00f59b" : "#fbbf24",
      heading: confirmed ? (es ? "⚽ ¡Hoy hay partido!" : "⚽ Today's match is on!") : (es ? "⚠️ Al partido de hoy le faltan jugadores" : "⚠️ Today's match still needs players"),
      paras: [(es ? "Resumen del día para <b>" : "Your matchday briefing for <b>") + esc(game.Title) + "</b> (" + captainLine + "):"],
      box: "<b>📊 " + (es ? "Estado" : "Status") + ":</b> " + (confirmed ? (es ? "CONFIRMADO" : "CONFIRMED") : (es ? "PENDIENTE" : "PENDING")) + " (" + notificationCountSummary(game, lang) + (es ? " · mínimo " : " · min ") + game.MinThreshold + " · " + esc(format) + ")<br>" +
        "<b>🕒 " + (es ? "Inicio" : "Kickoff") + ":</b> " + esc(game.Time) + "<br>" +
        "<b>📍 " + (es ? "Cancha" : "Field") + ":</b> " + esc(game.LocationName) + " (<a href='" + mapsUrl(game.locationDetails.address || game.LocationName) + "' style='color:#00f59b;'>" + (es ? "cómo llegar" : "directions") + "</a>)<br>" +
        "<b>🌤️ " + (es ? "Pronóstico" : "Forecast") + ":</b> " + esc(weatherText(o.weather, lang)) +
        (gear.length ? "<br><b>⚽ " + (es ? "Equipo" : "Gear") + ":</b> " + gear.join(" | ") : ""),
      after: [teams +
        section("#00f59b", es ? "✅ Confirmados" : "✅ Confirmed", yes, true, es ? "Nadie aún" : "Nobody yet") +
        (reconfirm.length ? section("#fbbf24", es ? "Por reconfirmar (lugares reservados)" : "Awaiting re-confirmation (places held)", reconfirm, true, "") : "") +
        (wait.length ? section("#38bdf8", es ? "⏳ En espera" : "⏳ Waitlist", wait, false, "") : "") +
        section("#fbbf24", es ? "🤔 Tal vez" : "🤔 Maybe", maybe, false, es ? "Nadie" : "Nobody") +
        section("#f87171", es ? "❌ No van" : "❌ Out", no, false, es ? "Nadie" : "Nobody")],
      buttons: [hubBtn(lang)]
    });
    return {
      subject: (confirmed ? (es ? "🟢 Hoy hay partido: " : "🟢 Game on today: ") : (es ? "🟡 Hoy faltan jugadores: " : "🟡 Players needed today: ")) + game.Title + " · " + game.Time,
      html: html
    };
  }, o.game, { priority: 30, skipAway: true });
}

function mailTwoHour(o) { // { game }
  return notificationMail(function(lang) {
    var es = mailLang(lang) === "es", game = o.game;
    var yes = game.roster.filter(function(r) { return r.status === "Yes" && !r.needsReconfirm; });
    return {
      subject: (es ? "🚨 Empezamos en ~2 horas: " : "🚨 Kickoff in ~2 hours: ") + game.Title + " · " + game.Time,
      html: mailLayout({
        heading: es ? "¡Empezamos en unas 2 horas!" : "Kickoff in about 2 hours!",
        paras: [(es ? "<b>" + esc(game.Title) + "</b> se juega en <b>" : "<b>" + esc(game.Title) + "</b> is on at <b>") + esc(game.LocationName) + "</b> (" + esc(game.Time) + ").",
                "<a href='" + mapsUrl(game.locationDetails.address || game.LocationName) + "' style='color:#00f59b;font-weight:bold;'>📍 " + (es ? "Cómo llegar" : "Directions") + "</a>"],
        box: "<b>" + (es ? "Confirmados (" : "Confirmed players (") + (game.confirmedCount === undefined ? game.yesCount : game.confirmedCount) + "):</b><br>" + rosterNames(yes, lang, true),
        buttons: [hubBtn(lang, "See your team & shirt →", "Ver tu equipo y camiseta →")]
      })
    };
  }, o.game, { priority: 20, timeSensitive: true, skipAway: true, expectedSignupStatus: "Yes", expectedGameStatus: "Confirmed" });
}

// Read-only deployment check. Run from the Apps Script editor after upgrading.
// This function does not send messages, change profiles, or rewrite league data.
function checkDeploymentV540() {
  var expectedTimeZone = "America/Los_Angeles";
  var report = {
    version: APP_VERSION,
    schemaVersion: PropertiesService.getScriptProperties().getProperty("SCHEMA_VERSION") || "",
    expectedTimeZone: expectedTimeZone,
    scriptTimeZone: Session.getScriptTimeZone(),
    spreadsheetTimeZone: SS.getSpreadsheetTimeZone(),
    missingSheets: [], missingColumns: [], warnings: [],
    tables: {}, automationTriggers: []
  };
  Object.keys(SCHEMA).forEach(function(name) {
    var sheet = SS.getSheetByName(name);
    if (!sheet) { report.missingSheets.push(name); return; }
    var width = sheet.getLastColumn();
    var headers = width ? sheet.getRange(1, 1, 1, width).getValues()[0].map(function(h) { return String(h || "").trim(); }) : [];
    SCHEMA[name].forEach(function(h) {
      if (headers.indexOf(h) === -1) report.missingColumns.push(name + "." + h);
    });
    report.tables[name] = { rows: Math.max(0, sheet.getLastRow() - 1) };
  });
  var owned = ["leagueAutomationDispatcher", "send9amGameDayUpdate", "sendDayBeforeNonRsvpReminder", "sendTwoHourReminders"];
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    var handler = trigger.getHandlerFunction();
    if (owned.indexOf(handler) !== -1) report.automationTriggers.push(handler);
  });
  if (report.scriptTimeZone !== expectedTimeZone) report.warnings.push("Set the Apps Script project time zone to America/Los_Angeles.");
  if (report.spreadsheetTimeZone !== expectedTimeZone) report.warnings.push("Set the spreadsheet time zone to America/Los_Angeles (Pacific Time).");
  if (report.automationTriggers.length !== 1 || report.automationTriggers[0] !== "leagueAutomationDispatcher") {
    report.warnings.push("Run installProductionTriggers() under the league owner account. There should be one leagueAutomationDispatcher and no legacy league reminder triggers for that account.");
  }
  if (report.version !== "5.4.0" || report.schemaVersion !== "5.4.0") report.warnings.push("Check the installed code and run upgradeToV540().");
  report.schemaReady = report.missingSheets.length === 0 && report.missingColumns.length === 0;
  report.ready = report.schemaReady && report.warnings.length === 0;
  Logger.log(JSON.stringify(report, null, 2));
  return report;
}
