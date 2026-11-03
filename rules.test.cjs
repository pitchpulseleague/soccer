#!/usr/bin/env node
'use strict';

// Standalone regression tests. Google services are replaced by in-memory doubles;
// no network requests, real spreadsheet writes, emails, or notifications occur.
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');

const root = path.resolve(__dirname, '..');
const sourcePath = process.argv[2] ? path.resolve(process.argv[2]) :
  [path.join(root, 'release/Code.gs'), path.join(root, 'Code.gs'), path.join(root, 'work/Code.rules.gs')].find(fs.existsSync);
if (!sourcePath) throw new Error('Pass the path to Code.gs as the first argument.');
const source = fs.readFileSync(sourcePath, 'utf8');
const NativeDate = Date;

function zoneParts(date, zone) {
  const out = {};
  new Intl.DateTimeFormat('en-US', { timeZone: zone, year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' }).formatToParts(date)
    .forEach(p => { if (p.type !== 'literal') out[p.type] = p.value; });
  return out;
}
function formatDate(date, zone, pattern) {
  const p = zoneParts(date, zone), hour = +p.hour, twelve = hour % 12 || 12;
  const iso = `${p.year}-${p.month}-${p.day}`, time = `${twelve}:${p.minute} ${hour >= 12 ? 'PM' : 'AM'}`;
  if (pattern === 'yyyy-MM-dd') return iso;
  if (pattern === 'HH:mm') return `${p.hour}:${p.minute}`;
  if (pattern === 'h:mm a') return time;
  if (pattern === 'MMM d, h:mm a') return `${new Intl.DateTimeFormat('en-US', { month: 'short', timeZone: zone }).format(date)} ${+p.day}, ${time}`;
  throw new Error(`Unimplemented test format: ${pattern}`);
}
function parseDate(text, zone) {
  const m = text.match(/^(\d{4})-(\d{2})-(\d{2}) (\d{1,2}):(\d{2}) (AM|PM)$/);
  if (!m) throw new Error(`Invalid parse fixture: ${text}`);
  const target = NativeDate.UTC(+m[1], +m[2] - 1, +m[3], +m[4] % 12 + (m[6] === 'PM' ? 12 : 0), +m[5]);
  let instant = target;
  for (let i = 0; i < 4; i++) {
    const p = zoneParts(new NativeDate(instant), zone);
    const represented = NativeDate.UTC(+p.year, +p.month - 1, +p.day, +p.hour, +p.minute, +p.second);
    instant += target - represented;
  }
  return new NativeDate(instant);
}
function fixture(when = '2026-11-03T18:00:00Z') {
  let clock = NativeDate.parse(when);
  class FakeDate extends NativeDate {
    constructor(...args) { super(...(args.length ? args : [clock])); }
    static now() { return clock; }
  }
  const props = {}, notices = [], statusNotices = [], activity = [];
  const store = { getProperty: k => props[k] ?? null, setProperty(k, v) { props[k] = String(v); return this; },
    deleteProperty(k) { delete props[k]; return this; }, getProperties: () => ({ ...props }),
    setProperties(values) { Object.assign(props, values); return this; } };
  const ctx = vm.createContext({ Date: FakeDate, console,
    SpreadsheetApp: { getActiveSpreadsheet: () => ({}) },
    PropertiesService: { getScriptProperties: () => store },
    Session: { getScriptTimeZone: () => 'UTC' }, // Deliberately different from the league.
    Utilities: { formatDate, parseDate, getUuid: () => 'local-test-uuid' },
    MailApp: { getRemainingDailyQuota: () => 100 },
    Logger: { log() {} },
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: s => ({ text: s, setMimeType() { return this; } }) }
  });
  vm.runInContext(source, ctx, { filename: sourcePath });
  const tables = Object.fromEntries(Object.keys(ctx.SCHEMA).map(n => [n, { name: n, rows: [] }]));
  const value = v => typeof v === 'string' && v.startsWith("'") ? v.slice(1) : v;
  ctx.T = name => tables[name];
  ctx.tv = (t, row, field) => t.rows[row]?.[field] ?? '';
  ctx.tHas = (t, field) => ctx.SCHEMA[t.name].includes(field);
  ctx.tw = (t, row, field, val) => { t.rows[row][field] = value(val); };
  ctx.twMany = (t, row, obj) => Object.entries(obj).forEach(([k, v]) => ctx.tw(t, row, k, v));
  ctx.tadd = (t, obj) => { const row = {}; Object.entries(obj).forEach(([k, v]) => { row[k] = value(v); }); t.rows.push(row); return t.rows.length - 1; };
  ctx.taddMany = (t, objs) => objs.forEach(o => ctx.tadd(t, o));
  ctx.tdel = (t, row) => t.rows.splice(row, 1);
  ctx.tobj = (t, row) => ({ ...t.rows[row] });
  ctx.tobjs = t => t.rows.map(o => ({ ...o }));
  const settings = { ...ctx.DEFAULT_LEAGUE_SETTINGS };
  ctx.getLeagueSettings = () => settings;
  ctx.getEmailQuotaLeft = () => 100;
  const admin = { email: 'admin@example.test', fullName: 'Admin', isAdmin: true };
  let user = { email: 'alice@example.test', fullName: 'Alice One', canBringGuests: true, isProtected: false };
  ctx.requireAdmin = () => admin;
  ctx.verifyGameManagerRequest = () => admin;
  ctx.verifyUserCredentials = () => user;
  ctx.resolvePlayerRef = v => v;
  ctx.okData = (_who, message, extra) => ({ success: true, message, ...(extra || {}) });
  ctx.fail = (error, extra) => ({ success: false, error, ...(extra || {}) });
  ctx.notifyMail = (...args) => notices.push(args);
  ctx.notifyStatusFlip = (...args) => statusNotices.push(args);
  ctx.appendGameActivityLog = (...args) => activity.push(args);
  ctx.revokeAllSessions = () => {};
  ctx.sendWaitlistPromotionEmails = (...args) => notices.push(['promotion', ...args]);
  const add = (name, row) => ctx.tadd(tables[name], row);
  for (const [email, fullName] of [['alice@example.test', 'Alice One'], ['bob@example.test', 'Bob Two'], ['charlie@example.test', 'Charlie Three']]) {
    add('Profiles', { Email: email, FullName: fullName, Position: 'MID', Role: 'Player', AccountStatus: 'Active', LateDrops: 0 });
  }
  add('Locations', { LocationID: 'L1', Name: 'Usual field', Address: 'San Marcos, CA', LocationStatus: 'Active' });
  const game = (id = 'G_test', opts = {}) => add('Games', { GameID: id, Title: 'Lunch soccer', Date: '2026-11-03', Time: '12:15 PM',
    LocationName: 'Usual field', MinThreshold: 2, MaxCap: 18, Status: 'Pending', ...opts });
  const signup = (email = 'alice@example.test', opts = {}) => add('Signups', { SignupID: `S_${tables.Signups.rows.length}`, GameID: 'G_test', Email: email,
    Status: 'Yes', Timestamp: new FakeDate(clock - 60000), PlusOne: '', ...opts });
  return { ctx, tables, settings, props, notices, statusNotices, activity, add, game, signup,
    setTime: iso => { clock = NativeDate.parse(iso); }, setUser: next => { user = { ...user, ...next }; },
    data: () => ctx.getAppData(), now: () => clock };
}

let passed = 0;
function test(name, fn) { try { fn(); passed++; console.log(`PASS ${name}`); } catch (e) { console.error(`FAIL ${name}`); throw e; } }

test('Pacific kickoff instants follow both DST transitions even when project zone is UTC', () => {
  const { ctx } = fixture();
  for (const [day, instant] of [['2026-10-30', '2026-10-30T19:15:00Z'], ['2026-11-03', '2026-11-03T20:15:00Z'],
    ['2027-03-12', '2027-03-12T20:15:00Z'], ['2027-03-15', '2027-03-15T19:15:00Z']]) {
    assert.equal(ctx.gameInstantMs(day, '12:15 PM'), NativeDate.parse(instant));
  }
  assert.equal(ctx.isoPlusDays('2026-11-01', 1), '2026-11-02');
  assert.equal(ctx.normalizeDateInput('2026-02-30'), '');
});

test('reconfirmation retains reserved seats but excludes them from Game on, even after kickoff', () => {
  const f = fixture(); f.game('G_test', { Date: '2026-11-04', MinThreshold: 3, ScheduleChangeAt: f.now() - 1000 });
  f.signup(undefined, { PlusOne: 'Guest A, Guest B' }); f.signup('bob@example.test');
  let g = f.data().games[0];
  assert.equal(g.yesCount, 4); assert.equal(g.confirmedCount, 0); assert.equal(g.pendingReconfirmCount, 4); assert.equal(g.pendingReconfirmPlayers, 2);
  assert.equal(g.Status, 'Pending');
  assert.equal(f.ctx.processSignupCore('G_test', 'alice@example.test', 'Yes').success, true);
  g = f.data().games[0]; assert.equal(g.confirmedCount, 3); assert.equal(g.Status, 'Confirmed');
  f.setTime('2026-11-04T20:30:00Z');
  g = f.data().games[0]; assert.equal(g.roster.find(p => p.email === 'bob@example.test').needsReconfirm, true);
  f.ctx.processSignupCore('G_test', 'bob@example.test', 'Yes');
  assert.equal(f.data().games[0].confirmedCount, 4);
});

test('legacy guest/member collision blocks a duplicate signup without merging identities, in both languages', () => {
  const f = fixture(); f.game(); f.signup(undefined, { PlusOne: 'Bob Two' });
  const before = f.tables.Signups.rows.length;
  let result = f.ctx.processSignupCore('G_test', 'bob@example.test', 'Yes');
  assert.equal(result.success, false); assert.match(result.error, /Alice One/); assert.match(result.error, /different person/);
  f.ctx.__LANG = 'es'; result = f.ctx.processSignupCore('G_test', 'bob@example.test', 'Yes');
  assert.equal(result.success, false); assert.match(result.error, /Alice One/); assert.match(result.error, /invitado/);
  assert.equal(f.tables.Signups.rows.length, before); assert.equal(f.data().games[0].yesCount, 2);
});

test('registered members cannot be newly added as guests even before they answer', () => {
  const f = fixture(); f.game(); f.signup();
  const result = f.ctx.updateSignupExtras({ gameId: 'G_test', field: 'guestNames', value: ['Bob Two'] });
  assert.equal(result.success, false); assert.match(result.error, /registered league member/);
});

test('existing guests can be removed after account privilege is revoked and league limit becomes zero', () => {
  const f = fixture(); f.game(); f.signup(undefined, { PlusOne: 'Outside A, Outside B' });
  f.settings.maxGuests = 0; f.setUser({ canBringGuests: false });
  assert.equal(f.ctx.updateSignupExtras({ gameId: 'G_test', field: 'guestNames', value: ['Outside A'] }).success, true);
  assert.equal(f.data().games[0].yesCount, 2);
  assert.equal(f.ctx.updateSignupExtras({ gameId: 'G_test', field: 'guestNames', value: [] }).success, true);
  assert.equal(f.data().games[0].yesCount, 1);
  assert.equal(f.ctx.updateSignupExtras({ gameId: 'G_test', field: 'guestNames', value: ['New Guest'] }).success, false);
});

test('releasing guests remains possible when an organizer walk-in has put a match over capacity', () => {
  const f = fixture(); f.game('G_test', { MaxCap: 2 });
  f.signup(undefined, { PlusOne: 'Outside A, Outside B' }); f.signup('bob@example.test');
  assert.equal(f.ctx.updateSignupExtras({ gameId: 'G_test', field: 'guestNames', value: ['Outside A'] }).success, true);
  assert.equal(f.data().games[0].yesCount, 3);
});

test('legacy guest collision is not double-counted by automatic waitlist promotion', () => {
  const f = fixture(); f.game('G_test', { MaxCap: 3 }); f.signup(undefined, { PlusOne: 'Bob Two' });
  f.signup('bob@example.test', { Status: 'Waitlist', Timestamp: new NativeDate(f.now() - 2000) });
  f.signup('charlie@example.test', { Status: 'Waitlist', Timestamp: new NativeDate(f.now() - 1000) });
  const recon = f.ctx.reconcileGameWaitlistAndStatus('G_test');
  assert.equal(recon.recalculatedYes, 3);
  assert.deepEqual(Array.from(recon.promotedEmails), ['charlie@example.test']);
  assert.equal(f.tables.Signups.rows[1].Status, 'Waitlist');
});

test('no-show event is idempotent, counts one outcome, and an attendance correction repairs it', () => {
  const f = fixture('2026-11-03T21:00:00Z'); f.settings.featAttendance = true; f.settings.featLateDrops = true;
  f.game(); f.signup();
  assert.equal(f.data().profiles['alice@example.test'].reliability, null);
  assert.equal(f.ctx.adminUpdateAttendance({ gameId: 'G_test', targetEmail: 'alice@example.test', newStatus: 'No' }).success, true);
  f.tables.Signups.rows[0].LateDropAt = new NativeDate('2026-11-03T18:00:00Z');
  let p = f.data().profiles['alice@example.test'];
  assert.equal(p.reliability, 0); assert.equal(p.reliabilitySampleCount, 1); assert.equal(p.noShows, 1);
  const notices = f.notices.length;
  f.ctx.adminUpdateAttendance({ gameId: 'G_test', targetEmail: 'alice@example.test', newStatus: 'No' });
  assert.equal(f.notices.length, notices); assert.equal(f.data().profiles['alice@example.test'].reliabilitySampleCount, 1);
  f.ctx.adminUpdateAttendance({ gameId: 'G_test', targetEmail: 'alice@example.test', newStatus: 'Yes' });
  p = f.data().profiles['alice@example.test'];
  assert.equal(p.reliability, 100); assert.equal(p.reliabilitySampleCount, 1); assert.equal(p.noShows, 0); assert.equal(p.lateDrops, 0);
  assert.equal(p.verifiedGamesPlayed, 1);
});

test('new players and legacy aggregate-only history do not get fabricated reliability scores', () => {
  const f = fixture(); f.tables.Profiles.rows[0].LateDrops = 7;
  const p = f.data().profiles['alice@example.test'];
  assert.equal(p.reliability, null); assert.equal(p.reliabilitySampleCount, 0); assert.equal(p.legacyLateDrops, 7);
});

test('no-show cannot be recorded before kickoff or against an unconfirmed player', () => {
  const f = fixture(); f.settings.featAttendance = true; f.game(); f.signup();
  assert.equal(f.ctx.adminUpdateAttendance({ gameId: 'G_test', targetEmail: 'alice@example.test', newStatus: 'No' }).success, false);
  f.setTime('2026-11-03T21:00:00Z');
  assert.equal(f.ctx.adminUpdateAttendance({ gameId: 'G_test', targetEmail: 'bob@example.test', newStatus: 'No' }).success, false);
});

test('player-of-the-match voting waits until scheduled end or final score and excludes recorded no-shows', () => {
  const f = fixture('2026-11-03T20:30:00Z');
  Object.assign(f.settings, { featTeams: true, featScores: true, featMotm: true, featAttendance: true });
  f.game(); f.signup(); f.signup('bob@example.test');
  assert.equal(f.ctx.voteMotm({ gameId: 'G_test', nomineeEmail: 'bob@example.test' }).success, false);
  f.tables.Games.rows[0].ScoreWhite = 2; f.tables.Games.rows[0].ScoreDark = 1;
  assert.equal(f.ctx.voteMotm({ gameId: 'G_test', nomineeEmail: 'bob@example.test' }).success, true);
  f.tables.Games.rows[0].ScoreWhite = ''; f.tables.Games.rows[0].ScoreDark = '';
  f.setTime('2026-11-03T21:31:00Z');
  assert.equal(f.ctx.voteMotm({ gameId: 'G_test', nomineeEmail: 'bob@example.test' }).success, true);
  f.ctx.adminUpdateAttendance({ gameId: 'G_test', targetEmail: 'bob@example.test', newStatus: 'No' });
  assert.equal(f.ctx.voteMotm({ gameId: 'G_test', nomineeEmail: 'bob@example.test' }).success, false);
});

test('late-drop boundaries preserve the existing three-hour through first-hour-of-play policy', () => {
  for (const [clock, expected] of [['2026-11-03T17:14:59Z', false], ['2026-11-03T17:15:00Z', true],
    ['2026-11-03T21:14:59Z', true], ['2026-11-03T21:15:00Z', false]]) {
    const f = fixture(clock); f.settings.featLateDrops = true; f.game(); f.signup();
    f.ctx.processSignupCore('G_test', 'alice@example.test', 'No');
    assert.equal(!!f.tables.Signups.rows[0].LateDropAt, expected, clock);
    assert.equal(f.tables.Profiles.rows[0].LateDrops, 0);
  }
});

test('repeated late drop is one event; organizer-confirmed attendance is not erased by a later RSVP', () => {
  const f = fixture('2026-11-03T18:00:00Z'); f.settings.featLateDrops = true; f.settings.featAttendance = true;
  f.game(); f.signup();
  f.ctx.processSignupCore('G_test', 'alice@example.test', 'No');
  const first = f.tables.Signups.rows[0].LateDropAt.getTime();
  f.ctx.processSignupCore('G_test', 'alice@example.test', 'Yes'); f.ctx.processSignupCore('G_test', 'alice@example.test', 'No');
  assert.equal(f.tables.Signups.rows[0].LateDropAt.getTime(), first);
  f.setTime('2026-11-03T21:00:00Z');
  f.ctx.adminUpdateAttendance({ gameId: 'G_test', targetEmail: 'alice@example.test', newStatus: 'Yes' });
  f.ctx.processSignupCore('G_test', 'alice@example.test', 'No');
  const p = f.data().profiles['alice@example.test'];
  assert.equal(p.verifiedGamesPlayed, 1); assert.equal(p.reliability, 100); assert.equal(p.reliabilitySampleCount, 1);
});

test('suspension frees future spots while preserving completed same-day roster, captain, and teams', () => {
  const f = fixture('2026-11-03T21:00:00Z'); f.settings.featTeams = true; f.settings.featCaptains = true;
  const lineup = JSON.stringify({ white: [{ key: 'alice@example.test', email: 'alice@example.test', fullName: 'Alice One', position: 'MID' }], dark: [] });
  f.game('G_done', { CaptainEmail: 'alice@example.test', TeamsJSON: lineup });
  f.game('G_future', { Date: '2026-11-04', CaptainEmail: 'alice@example.test', TeamsJSON: lineup });
  f.signup(undefined, { GameID: 'G_done' }); f.signup(undefined, { GameID: 'G_future' });
  f.ctx.adminToggleSuspendUser({ targetEmail: 'alice@example.test', newStatus: 'Suspended' });
  assert.equal(f.tables.Signups.rows[0].Status, 'Yes'); assert.equal(f.tables.Signups.rows[1].Status, 'No');
  assert.equal(f.tables.Games.rows[0].CaptainEmail, 'alice@example.test'); assert.equal(f.tables.Games.rows[0].TeamsJSON, lineup);
  assert.equal(f.tables.Games.rows[1].CaptainEmail, '');
});

function createSeries(f) {
  const input = { title: 'Lunch soccer', locationName: 'Usual field', daysOfWeek: '1,2,3,4,5', time: '12:15 PM',
    startDate: '2026-11-03', endDate: '2027-02-26', minThreshold: 2, maxCap: 18, effectiveDate: '2026-11-03' };
  assert.equal(f.ctx.adminSaveRecurringSeries(input).success, true);
  return { ...input, seriesId: f.tables.Recurring.rows[0].SeriesID };
}
const onDate = (f, date) => f.tables.Games.rows.filter(g => g.Date === date);

test('future series edits beyond 28 days retain old weekdays/time until the actual effective date', () => {
  const f = fixture(), original = createSeries(f);
  f.ctx.adminSaveRecurringSeries({ ...original, time: '1:00 PM', daysOfWeek: '1,3,5', effectiveDate: '2026-12-15' });
  assert.equal(onDate(f, '2026-11-05')[0].Time, '12:15 PM');
  f.setTime('2026-11-20T18:00:00Z'); f.ctx.extendRecurringSeries();
  assert.equal(onDate(f, '2026-12-03')[0].Time, '12:15 PM');
  assert.equal(onDate(f, '2026-12-14')[0].Time, '12:15 PM');
  assert.equal(onDate(f, '2026-12-15').length, 0);
  assert.equal(onDate(f, '2026-12-16')[0].Time, '1:00 PM');
});

test('future edits inside the current horizon update only applicable dates and preserve cancelled exceptions', () => {
  const f = fixture(), original = createSeries(f);
  onDate(f, '2026-11-17')[0].Status = 'Cancelled';
  f.ctx.adminSaveRecurringSeries({ ...original, time: '1:00 PM', effectiveDate: '2026-11-16' });
  assert.equal(onDate(f, '2026-11-13')[0].Time, '12:15 PM');
  assert.equal(onDate(f, '2026-11-16')[0].Time, '1:00 PM');
  assert.equal(onDate(f, '2026-11-17')[0].Status, 'Cancelled');
  assert.equal(onDate(f, '2026-11-17')[0].Time, '12:15 PM');
  assert.equal(new Set(f.tables.Games.rows.map(g => g.Date)).size, f.tables.Games.rows.length);
});

test('editing a legacy long-horizon series still opens newly selected future weekdays', () => {
  const f = fixture(), original = createSeries(f);
  f.tables.Recurring.rows[0].GeneratedThrough = '2027-02-01';
  f.ctx.adminSaveRecurringSeries({ ...original, daysOfWeek: '1,2,3,4,5,6', effectiveDate: '2026-11-04' });
  assert.equal(f.tables.Recurring.rows[0].GeneratedThrough, '2026-12-01');
  f.setTime('2026-11-20T18:00:00Z'); f.ctx.extendRecurringSeries();
  assert.equal(onDate(f, '2026-12-05').length, 1);
});

test('stopping beyond the generated horizon continues creating all applicable pre-stop games', () => {
  const f = fixture(), original = createSeries(f);
  f.ctx.adminStopRecurringSeries({ seriesId: original.seriesId, effectiveDate: '2027-01-01' });
  assert.equal(f.tables.Recurring.rows.length, 1);
  f.setTime('2026-12-10T18:00:00Z'); f.ctx.extendRecurringSeries();
  assert.equal(onDate(f, '2026-12-31').length, 1); assert.equal(onDate(f, '2027-01-01').length, 0);
  assert.equal(f.tables.Recurring.rows[0].EndDate, '2026-12-31');
});

test('series stop cancels without deleting RSVP history; started games are protected and no duplicate dates regenerate', () => {
  const f = fixture(), original = createSeries(f), today = onDate(f, '2026-11-03')[0], next = onDate(f, '2026-11-04')[0];
  f.signup(undefined, { GameID: today.GameID }); f.signup('bob@example.test', { GameID: next.GameID });
  f.setTime('2026-11-03T21:00:00Z');
  f.ctx.adminStopRecurringSeries({ seriesId: original.seriesId, effectiveDate: '2026-11-03' });
  assert.notEqual(today.Status, 'Cancelled'); assert.equal(next.Status, 'Cancelled'); assert.equal(f.tables.Signups.rows.length, 2);
  next.Status = 'Pending'; // An organizer can restore one cancelled occurrence.
  f.ctx.extendRecurringSeries(); assert.equal(onDate(f, '2026-11-04').length, 1);
});

test('removing a weekday cancels its future occurrence and retains its RSVP for review', () => {
  const f = fixture(), original = createSeries(f), removed = onDate(f, '2026-11-05')[0];
  f.signup(undefined, { GameID: removed.GameID });
  f.ctx.adminSaveRecurringSeries({ ...original, daysOfWeek: '1,3,5', effectiveDate: '2026-11-04' });
  assert.equal(removed.Status, 'Cancelled'); assert.equal(f.tables.Signups.rows[0].GameID, removed.GameID);
  assert.equal(onDate(f, '2026-11-05').length, 1);
});

test('Rule2B stays open after scheduled end until score or league-day rollover', () => {
  const f = fixture('2026-11-04T01:00:00Z'); f.game(); f.signup(); // 5 PM in California.
  assert.equal(f.ctx.gameLockState(f.ctx.gameRecord(0)).locked, false);
  f.setTime('2026-11-04T08:01:00Z');
  assert.equal(f.ctx.gameLockState(f.ctx.gameRecord(0)).reason, 'past');
});

console.log(`\n${passed} rules regression tests passed against ${sourcePath}`);
