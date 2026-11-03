#!/usr/bin/env node
'use strict';
// Offline behavior tests: every Google service is an in-memory double. No live
// spreadsheet, API, notification, mailbox, or network access is used.
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const crypto = require('node:crypto');
const assert = require('node:assert/strict');
const root = path.resolve(__dirname, '..');
const sourcePath = process.argv[2] ? path.resolve(process.argv[2]) :
  [path.join(root, 'release/Code.gs'), path.join(root, 'Code.gs'), path.join(root, 'work/Code.backend.gs')].find(fs.existsSync);
if (!sourcePath) throw new Error('Pass the path to Code.gs as the first argument.');
const source = fs.readFileSync(sourcePath, 'utf8');
const NativeDate = Date;
function zoneParts(date, zone) {
  const p = {};
  new Intl.DateTimeFormat('en-US', { timeZone: zone, year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' }).formatToParts(date)
    .forEach(x => { if (x.type !== 'literal') p[x.type] = x.value; });
  return p;
}
function formatDate(date, zone, pattern) {
  const p = zoneParts(date, zone), hour = +p.hour;
  const iso = `${p.year}-${p.month}-${p.day}`, time = `${hour % 12 || 12}:${p.minute} ${hour >= 12 ? 'PM' : 'AM'}`;
  if (pattern === 'yyyy-MM-dd') return iso;
  if (pattern === 'H') return String(hour);
  if (pattern === 'm') return String(+p.minute);
  if (pattern === 'HH:mm') return `${p.hour}:${p.minute}`;
  if (pattern === 'h:mm a') return time;
  if (pattern === 'MMM d, h:mm a') return `${new Intl.DateTimeFormat('en-US', { month: 'short', timeZone: zone }).format(date)} ${+p.day}, ${time}`;
  if (pattern === 'yyyy-MM-dd HH:mm:ss') return `${iso} ${p.hour}:${p.minute}:${p.second}`;
  throw new Error(`Unimplemented test date format: ${pattern}`);
}
function parseDate(text, zone) {
  const m = text.match(/^(\d{4})-(\d{2})-(\d{2}) (\d{1,2}):(\d{2}) (AM|PM)$/);
  if (!m) throw new Error(`Invalid parse fixture: ${text}`);
  const target = NativeDate.UTC(+m[1], +m[2] - 1, +m[3], +m[4] % 12 + (m[6] === 'PM' ? 12 : 0), +m[5]);
  let instant = target;
  for (let i = 0; i < 4; i++) {
    const p = zoneParts(new NativeDate(instant), zone);
    instant += target - NativeDate.UTC(+p.year, +p.month - 1, +p.day, +p.hour, +p.minute, +p.second);
  }
  return new NativeDate(instant);
}
const memo = v => typeof v === 'string' && v.startsWith("'") ? v.slice(1) : v;
function sheet(name, initial = []) {
  const data = initial.map(r => r.slice());
  return {
    name, data, getName() { return name; }, getLastRow() { return data.length; },
    getLastColumn() { return Math.max(0, ...data.map(r => r.length)); },
    appendRow(row) { data.push(Array.from(row, memo)); }, setFrozenRows() { return this; },
    deleteRow(n) { data.splice(n - 1, 1); },
    getDataRange() { return this.getRange(1, 1, this.getLastRow(), this.getLastColumn()); },
    getRange(row, col, rows = 1, cols = 1) {
      const range = {
        getValues: () => Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => data[row + r - 1]?.[col + c - 1] ?? '')),
        setValue(v) { return this.setValues([[v]]); },
        setValues(values) {
          values.forEach((line, r) => line.forEach((v, c) => {
            while (data.length < row + r) data.push([]);
            data[row + r - 1][col + c - 1] = memo(v);
          })); return this;
        }, setFontWeight() { return this; }, setBackground() { return this; }, setFontColor() { return this; }
      }; return range;
    }
  };
}
function fixture(when = '2026-11-03T16:35:00Z') {
  let clock = NativeDate.parse(when), quota = 100, uuid = 0, lockHeld = false;
  class FakeDate extends NativeDate { constructor(...args) { super(...(args.length ? args : [clock])); } static now() { return clock; } }
  const props = { MAGIC_SECRET: 'unchanged-private-test-secret', ONESIGNAL_APP_ID: 'test-app', ONESIGNAL_REST_KEY: 'test-key' };
  const sheets = {}, mail = [], push = [], logs = [], triggers = [];
  let pushResponse = { code: 200, body: { id: 'provider-test-id' } }, mailFailure = null;
  const store = { getProperty: k => props[k] ?? null, setProperty(k, v) { props[k] = String(v); return this; },
    deleteProperty(k) { delete props[k]; return this; }, getProperties: () => ({ ...props }),
    setProperties(values) { Object.assign(props, values); return this; } };
  const ss = { getSheetByName: n => sheets[n] || null, insertSheet(n) { return sheets[n] = sheet(n); },
    getSheets: () => Object.values(sheets), deleteSheet(s) { delete sheets[s.name]; }, getSpreadsheetTimeZone: () => 'UTC' };
  const lock = { waitLock() { assert.equal(lockHeld, false, 'Nested script lock would deadlock'); lockHeld = true; },
    tryLock() { if (lockHeld) return false; lockHeld = true; return true; }, releaseLock() { lockHeld = false; } };
  const ctx = vm.createContext({ Date: FakeDate, console,
    SpreadsheetApp: { getActiveSpreadsheet: () => ss },
    PropertiesService: { getScriptProperties: () => store }, LockService: { getScriptLock: () => lock },
    Session: { getScriptTimeZone: () => 'UTC' }, // Deliberately different from Pacific league time.
    Utilities: { formatDate, parseDate, getUuid: () => `00000000-0000-4000-8000-${String(++uuid).padStart(12, '0')}`,
      DigestAlgorithm: { SHA_256: 'sha256' }, computeDigest: (_alg, raw) => [...crypto.createHash('sha256').update(raw).digest()],
      base64EncodeWebSafe: bytes => Buffer.from(bytes).toString('base64url') },
    MailApp: { getRemainingDailyQuota: () => quota, sendEmail(options) { mail.push({ ...options }); if (mailFailure) throw mailFailure; quota--; } },
    UrlFetchApp: { fetch(url, options) {
      assert.match(url, /^https:\/\/api.onesignal.com\//, 'Only the mocked OneSignal request is expected');
      push.push({ url, options, data: JSON.parse(options.payload) });
      if (pushResponse.throw) throw new Error('Mock timeout');
      return { getResponseCode: () => pushResponse.code, getContentText: () => JSON.stringify(pushResponse.body), getAllHeaders: () => pushResponse.headers || {} };
    } },
    ScriptApp: { getProjectTriggers: () => triggers.slice(), deleteTrigger(t) { triggers.splice(triggers.indexOf(t), 1); }, newTrigger(handler) {
      const spec = { handler, getHandlerFunction: () => handler }, chain = { timeBased() { return this; }, everyMinutes(n) { spec.minutes = n; return this; },
        inTimezone(z) { spec.zone = z; return this; }, create() { triggers.push(spec); return spec; } }; return chain;
    } },
    Logger: { log: x => logs.push(x) },
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: text => ({ text, setMimeType() { return this; } }) },
    HtmlService: { createHtmlOutput: html => ({ html, setTitle() { return this; }, addMetaTag() { return this; } }) }
  });
  vm.runInContext(source, ctx, { filename: sourcePath });
  ctx.ensureSchema(true);
  // The notification tests isolate queue behavior from independent match-rule changes.
  ctx.gameInstantMs = (day, time) => parseDate(`${day} ${time}`, 'America/Los_Angeles').getTime();
  ctx.gameRecord = r => {
    const t = ctx.T('Games'), get = h => ctx.tv(t, r, h);
    return { id: get('GameID'), date: get('Date'), time: get('Time'), location: get('LocationName'), title: get('Title'), status: get('Status'),
      maxCap: +get('MaxCap') || 18, minThreshold: +get('MinThreshold') || 2, features: { lateDrops: true } };
  };
  const add = (name, obj) => ctx.tadd(ctx.T(name), obj);
  const profile = (email = 'alice@example.test', pref = 'Both', extra = {}) => {
    add('Profiles', { Email: email, FullName: email.split('@')[0], NotifyPref: pref, AccountStatus: 'Active', Lang: 'en', ...extra });
    ctx.__PREF_MAP = null; ctx.__LANG_MAP = null;
    return email;
  };
  const game = (extra = {}) => add('Games', { GameID: 'G_test', Title: 'Lunch soccer', Date: '2026-11-03', Time: '12:15 PM',
    LocationName: 'Usual field', MinThreshold: 2, MaxCap: 18, Status: 'Pending', ...extra });
  const signup = (email = 'alice@example.test', extra = {}) => add('Signups', { GameID: 'G_test', Email: email, Status: 'Yes', PlusOne: '', ...extra });
  const notifications = () => Array.from(ctx.tobjs(ctx.T('Notifications')));
  const queue = (recipients, opts = {}) => ctx.dispatchLeagueNotification(recipients, 'Test update', '<p>Test</p>', 'Test update', 'Please check the match.', ctx.SITE_URL,
    { mail: () => ({ subject: 'Test update', html: '<p>Test message.</p>', text: 'Test message.' }), ...opts });
  const response = value => JSON.parse(value.text);
  return { ctx, props, sheets, triggers, mail, push, logs, add, profile, game, signup, notifications, queue, response,
    now: () => clock, setTime: iso => { clock = NativeDate.parse(iso); }, advance: ms => { clock += ms; }, setQuota: n => { quota = n; }, quota: () => quota,
    setPush: obj => { pushResponse = obj; }, setMailFailure: err => { mailFailure = err; } };
}
let passed = 0;
function test(name, fn) { try { fn(); passed++; console.log(`PASS ${name}`); } catch (e) { console.error(`FAIL ${name}`); throw e; } }

test('legacy unsigned GET cannot mint an RSVP signature; only the new token namespace is valid', () => {
  const f = fixture(), { ctx } = f; f.profile(); f.game(); f.signup(undefined, { PlusOne: 'Guest A' });
  const old = crypto.createHash('sha256').update('alice@example.test|G_test|' + f.props.MAGIC_SECRET).digest('base64url').slice(0, 22);
  const token = ctx.generateMagicToken('alice@example.test', 'G_test'); assert.notEqual(token, old);
  const unsigned = ctx.doGet({ parameter: { action: 'magicRsvp', email: 'alice@example.test', gameId: 'G_test', status: 'No' } }).html;
  assert.doesNotMatch(unsigned, /sig=|rsvp=/); assert.match(unsigned, /old or invalid/);
  const oldResponse = f.response(ctx.emailAnswer({ email: 'alice@example.test', gameId: 'G_test', status: 'No', sig: old }));
  assert.equal(oldResponse.badLink, true);
  ctx.gameLockState = () => ({ locked: false });
  ctx.gameCommittedCounts = () => ({ reserved: 2, confirmed: 2, pending: 0 });
  let reconfirm = true; ctx.signupNeedsReconfirm = () => reconfirm;
  let writes = 0;
  ctx.processSignupCore = () => { writes++; return { success: true, finalStatus: 'No' }; };
  const args = { email: 'alice@example.test', gameId: 'G_test', status: 'No', sig: token };
  const preview = f.response(ctx.previewEmailAnswer(args));
  assert.equal(preview.success, true); assert.equal(preview.guestCount, 1); assert.equal(preview.currentStatus, 'Yes'); assert.equal(preview.needsReconfirm, true); assert.equal(preview.waitPos, 0); assert.equal(writes, 0);
  reconfirm = false; assert.equal(f.response(ctx.previewEmailAnswer(args)).needsReconfirm, false);
  const landing = ctx.doGet({ parameter: { action: 'magicRsvp', email: args.email, gameId: args.gameId, status: args.status, token } }).html;
  assert.match(landing, /sig=/); assert.equal(writes, 0);
  assert.equal(f.response(ctx.emailAnswer(args)).success, true); assert.equal(writes, 1);
  assert.equal(f.props.MAGIC_SECRET, 'unchanged-private-test-secret');
});

test('signed preview warns for the existing late-drop interval, including the first hour of play', () => {
  const f = fixture(), { ctx } = f; f.profile(); f.game(); f.signup();
  ctx.gameLockState = () => ({ locked: false });
  const args = { email: 'alice@example.test', gameId: 'G_test', status: 'No', sig: ctx.generateMagicToken('alice@example.test', 'G_test') };
  for (const [when, expected] of [
    ['2026-11-03T17:14:59.999Z', false], ['2026-11-03T17:15:00.000Z', true],
    ['2026-11-03T20:15:00.000Z', true], ['2026-11-03T21:14:59.999Z', true], ['2026-11-03T21:15:00.000Z', false]
  ]) {
    f.setTime(when); assert.equal(f.response(ctx.previewEmailAnswer(args)).lateDropApplies, expected, when);
  }
  f.setTime('2026-11-03T20:30:00Z'); args.status = 'Yes';
  assert.equal(f.response(ctx.previewEmailAnswer(args)).lateDropApplies, false);
});

test('setup is additive, idempotent, quiet, preserves credentials/settings, and leaves unrelated triggers', () => {
  const f = fixture(), { ctx } = f;
  f.profile(); f.game(); const snapshots = JSON.stringify(f.sheets.Profiles.data);
  f.props.LEAGUE_SETTINGS = JSON.stringify({ msgConfirm: true, featTeams: true, featScores: true, maxGuests: 1 });
  f.props.SESSION_alice = 'preserved-session'; f.props.ONESIGNAL_REST_KEY = 'preserved-key';
  f.sheets.Extra = sheet('Extra', [['Keep'], ['value']]);
  const other = { getHandlerFunction: () => 'unrelatedFunction' }; f.triggers.push(other, { getHandlerFunction: () => 'send9amGameDayUpdate' });
  const before = f.props.LEAGUE_SETTINGS;
  ctx.upgradeToV540(); ctx.upgradeToV540();
  assert.equal(JSON.stringify(f.sheets.Profiles.data), snapshots); assert.equal(f.props.LEAGUE_SETTINGS, before);
  assert.equal(f.props.SESSION_alice, 'preserved-session'); assert.equal(f.props.ONESIGNAL_REST_KEY, 'preserved-key');
  assert.ok(f.sheets.Extra); assert.equal(f.mail.length + f.push.length, 0);
  assert.equal(f.triggers.length, 2); assert.ok(f.triggers.includes(other));
  assert.equal(f.triggers.find(t => t.handler).minutes, 5); assert.equal(f.triggers.find(t => t.handler).zone, 'America/Los_Angeles');
  ctx.pauseLeagueAutomations(); assert.equal(f.triggers.length, 1); assert.equal(f.triggers[0], other);
});

test('recommended launch switches are explicit and preserve unrelated choices', () => {
  const f = fixture(), { ctx } = f;
  f.props.LEAGUE_SETTINGS = JSON.stringify({ featTeams: true, featScores: true, featMotm: true, featAttendance: true, featLateDrops: true, msgConfirm: true, msgTwoHour: true, maxGuests: 1, msgDayBefore: false });
  ctx.__SETTINGS = null; ctx.freezeStartedMatches = () => {};
  assert.equal(ctx.getLeagueSettings().msgConfirm, true);
  ctx.applyRecommendedLaunchSettings();
  const s = ctx.getLeagueSettings();
  for (const k of ['featScores', 'featMotm', 'featAttendance', 'featLateDrops', 'msgConfirm', 'msgTwoHour']) assert.equal(s[k], false);
  assert.equal(s.featRules, true); assert.equal(s.featTeams, true); assert.equal(s.maxGuests, 1); assert.equal(s.msgDayBefore, false);
  assert.equal(f.mail.length + f.push.length, 0);
});

test('Email preference survives push-first jobs; no subscribed push device falls back to email', () => {
  const f = fixture(); f.profile('email@example.test', 'Email'); f.profile('push@example.test', 'Push');
  f.queue(['email@example.test', 'push@example.test'], { pushOnly: true });
  f.setPush({ code: 200, body: { id: '' } });
  f.ctx.processNotificationOutbox();
  assert.equal(f.mail.length, 2); assert.equal(f.push.length, 1);
  assert.equal(f.push[0].data.include_aliases.external_id[0], 'push@example.test');
  assert.ok(f.notifications().every(r => r.EmailState === 'accepted'));
  assert.equal(f.ctx.notificationSummary().accepted, 2);
});

test('an accepted push is never repeated while its email waits for quota', () => {
  const f = fixture(); f.profile(); f.setQuota(10); f.queue(['alice@example.test']);
  f.ctx.processNotificationOutbox();
  assert.equal(f.push.length, 1); assert.equal(f.mail.length, 0);
  assert.equal(f.notifications()[0].PushState, 'accepted'); assert.equal(f.notifications()[0].EmailState, 'pending');
  f.setQuota(100); f.advance(31 * 60000); f.ctx.processNotificationOutbox();
  assert.equal(f.push.length, 1); assert.equal(f.mail.length, 1); assert.equal(f.notifications()[0].EmailState, 'accepted');
});

test('critical messages are processed before routine jobs, respecting recovery quota reserve', () => {
  const f = fixture(); f.profile('routine@example.test', 'Email'); f.queue(['routine@example.test'], { priority: 10 });
  const critical = Array.from({ length: 25 }, (_, i) => f.profile(`critical${i}@example.test`, 'Email'));
  f.queue(critical, { critical: true }); f.setQuota(27); f.ctx.processNotificationOutbox();
  assert.equal(f.mail.length, 25); assert.ok(f.mail.every(m => m.to.startsWith('critical'))); assert.equal(f.quota(), 2);
  assert.equal(f.notifications().find(r => r.Recipient === 'routine@example.test').EmailState, 'pending');
});

test('ambiguous email outcomes stop for review, while push retries retain their UUID', () => {
  const f = fixture(); f.profile('mail@example.test', 'Email'); f.profile('push@example.test', 'Both');
  f.queue(['mail@example.test']); f.setMailFailure(new Error('Transport interruption'));
  f.ctx.processNotificationOutbox(); assert.equal(f.notifications()[0].EmailState, 'review');
  f.setMailFailure(null); f.advance(3600000); f.ctx.processNotificationOutbox(); assert.equal(f.mail.length, 1);
  f.queue(['push@example.test']); f.setQuota(10); f.setPush({ throw: true }); f.ctx.processNotificationOutbox();
  const first = f.push[0].data.idempotency_key;
  f.advance(31 * 60000); f.setPush({ code: 200, body: { id: 'accepted-on-retry' } }); f.ctx.processNotificationOutbox();
  assert.equal(f.push.length, 2); assert.equal(f.push[1].data.idempotency_key, first);
  assert.equal(f.notifications().find(r => r.Recipient === 'push@example.test').PushState, 'accepted');
});

test('an interrupted email lease is not resent, and accepted channels survive recovery', () => {
  const f = fixture(); f.profile(); f.queue(['alice@example.test']);
  const t = f.ctx.T('Notifications');
  f.ctx.twMany(t, 0, { EmailState: 'sending', PushState: 'accepted', LeaseToken: 'old', LeaseUntil: new Date(f.now() - 1) });
  f.ctx.processNotificationOutbox(); assert.equal(f.notifications()[0].EmailState, 'review');
  assert.equal(f.notifications()[0].PushState, 'accepted'); assert.equal(f.mail.length + f.push.length, 0);
});

test('Pending to Confirmed retains schedule notices; moved or cancelled games expire obsolete notices', () => {
  const f = fixture(); f.profile('keep@example.test', 'Email'); f.profile('stale@example.test', 'Email'); f.game();
  f.queue(['keep@example.test'], { gameId: 'G_test', supersedeKey: 'game-update:G_test', critical: true });
  f.ctx.tw(f.ctx.T('Games'), 0, 'Status', 'Confirmed'); f.ctx.processNotificationOutbox(); assert.equal(f.mail.length, 1);
  f.queue(['stale@example.test'], { gameId: 'G_test', critical: true });
  f.ctx.tw(f.ctx.T('Games'), 0, 'Status', 'Cancelled'); f.ctx.processNotificationOutbox(); assert.equal(f.mail.length, 1);
  assert.equal(f.notifications()[1].EmailState, 'expired');
  f.queue(['stale@example.test'], { gameId: 'G_test', critical: true }); f.ctx.tw(f.ctx.T('Games'), 0, 'Time', '1:00 PM');
  f.ctx.processNotificationOutbox(); assert.equal(f.mail.length, 1); assert.equal(f.notifications()[2].EmailState, 'expired');
});

test('new state changes supersede older unsent changes, and after-match notices retain a usable TTL', () => {
  const f = fixture('2026-11-03T23:00:00Z'); f.profile('alice@example.test', 'Email'); f.game({ Status: 'Cancelled' });
  f.queue(['alice@example.test'], { gameId: 'G_test', supersedeKey: 'game-update:G_test', critical: true });
  f.ctx.tw(f.ctx.T('Games'), 0, 'Status', 'Pending');
  f.queue(['alice@example.test'], { gameId: 'G_test', supersedeKey: 'game-update:G_test', critical: true });
  assert.equal(f.notifications()[0].EmailState, 'expired');
  assert.ok(new Date(f.notifications()[1].ExpiresAt).getTime() > f.now());
  f.ctx.processNotificationOutbox(); assert.equal(f.mail.length, 1); assert.equal(f.notifications()[1].EmailState, 'accepted');
});

test('a newer series notice supersedes only recipients receiving that replacement', () => {
  const f = fixture(); f.profile('alice@example.test', 'Email'); f.profile('bob@example.test', 'Email');
  f.queue(['alice@example.test', 'bob@example.test'], { supersedeKey: 'series-update:series1', critical: true });
  f.queue(['alice@example.test'], { supersedeKey: 'series-update:series1', critical: true });
  const rows = f.notifications();
  assert.equal(rows[0].EmailState, 'expired');
  assert.equal(rows[1].EmailState, 'pending');
  assert.equal(rows[2].EmailState, 'pending');
  f.ctx.processNotificationOutbox();
  assert.equal(f.mail.length, 2); assert.equal(new Set(f.mail.map(m => m.to)).size, 2);
});

test('reminders expire when a player has answered, and at the explicit cutoff', () => {
  const f = fixture(); f.profile('alice@example.test', 'Email'); f.game();
  f.queue(['alice@example.test'], { gameId: 'G_test', requireNoAnswer: true, timeSensitive: true }); f.signup();
  f.ctx.processNotificationOutbox(); assert.equal(f.notifications()[0].EmailState, 'expired');
  f.queue(['alice@example.test'], { gameId: 'G_test', expiresAtMs: f.now() + 1000 });
  f.advance(2000); f.ctx.processNotificationOutbox(); assert.equal(f.notifications()[1].EmailState, 'expired'); assert.equal(f.mail.length, 0);
});

test('morning dispatch queues once around 8:30 Pacific and does not create new reminders after 9:00', () => {
  const f = fixture(), { ctx } = f; f.profile('alice@example.test', 'Email'); f.game(); f.signup();
  const g = { GameID: 'G_test', Title: 'Lunch soccer', Date: '2026-11-03', Time: '12:15 PM', LocationName: 'Usual field', Status: 'Pending',
    yesCount: 1, confirmedCount: 1, MinThreshold: 2, MaxCap: 18, roster: [{ email: 'alice@example.test', fullName: 'Alice', status: 'Yes' }], locationDetails: { address: 'Usual field' } };
  ctx.getAppData = () => ({ games: [g], profiles: { 'alice@example.test': { status: 'Active' } } });
  ctx.getGameOrganizersRecipients = () => [];
  ctx.queueScheduledReminders(new ctx.Date()); ctx.queueScheduledReminders(new ctx.Date());
  assert.equal(f.notifications().length, 1);
  assert.equal(formatDate(new Date(f.notifications()[0].ExpiresAt), 'America/Los_Angeles', 'h:mm a'), '9:15 AM');
  const payload = JSON.parse(f.notifications()[0].PayloadJSON); assert.match(payload.body, /9:15 AM/);
  f.setTime('2026-11-03T17:10:00Z'); ctx.queueScheduledReminders(new ctx.Date()); assert.equal(f.notifications().length, 1);
  f.setTime('2026-11-03T17:16:00Z'); ctx.processNotificationOutbox(); assert.equal(f.mail.length, 0); assert.equal(f.notifications()[0].EmailState, 'expired');
});

test('queue write failure is visible to admin and success responses do not claim delivery', () => {
  const f = fixture(); f.profile(); f.ctx.NOTIFY_MAX_PENDING = 0;
  const result = f.queue(['alice@example.test']); assert.equal(result.queued, false);
  f.ctx.getRoleFilteredAppData = () => ({ success: true }); f.ctx.publicUser = x => x;
  const response = f.response(f.ctx.okData({ email: 'admin@example.test' }, 'Match saved.'));
  assert.equal(response.notificationWarning, true); assert.match(response.message, /could not be queued/);
  assert.match(f.ctx.notificationSummary().lastError, /queue is full/);
  for (const key of ['msgNudged', 'msgBlastSent', 'msgCancelled', 'msgRestored', 'msgInviteSent', 'msgDirectSent']) {
    f.ctx.__LANG = 'en'; assert.doesNotMatch(f.ctx.tr(key), /delivered|\bnotified\b|\bsent\b/);
  }
});

test('personal invite has a full HTML fallback link and plain-text link, and waits in the outbox', () => {
  const f = fixture(); f.add('Invites', { Email: 'new@example.test', Status: 'Queued' });
  const result = f.ctx.sendInviteEmail('new@example.test', 'Organizer'); assert.equal(result.queued, true);
  assert.equal(f.mail.length + f.push.length, 0);
  const row = f.notifications()[0], payload = JSON.parse(row.PayloadJSON), link = f.ctx.buildInviteUrl('new@example.test');
  assert.ok(payload.plain.includes(link)); assert.ok(payload.html.includes(f.ctx.escapeHtmlServer(link)));
  assert.match(payload.html, /12:15 PM Pacific/); assert.match(payload.plain, /six-digit PIN/); assert.equal(row.PushState, 'skipped');
});

test('email templates use committed attendance and describe a full maybe response as waitlist', () => {
  const f = fixture(), { ctx } = f;
  const game = { GameID: 'G_test', Title: 'Lunch soccer', Date: '2026-11-03', Time: '12:15 PM', LocationName: 'Usual field',
    Status: 'Pending', yesCount: 18, confirmedCount: 0, pendingReconfirmCount: 18, MinThreshold: 14, MaxCap: 18,
    roster: [{ email: 'alice@example.test', fullName: 'Alice', status: 'Yes', needsReconfirm: true }], locationDetails: { address: 'Usual field' } };
  const opts = { game, playerEmail: 'alice@example.test', playerName: 'Alice', managerName: 'Admin', spotsLeft: 0, showWeather: false };
  const nudge = ctx.mailNudge(opts)('en').html;
  assert.match(nudge, /0 confirmed/); assert.match(nudge, /18 places held/); assert.match(nudge, /14 more needed/); assert.doesNotMatch(nudge, /game on, 0/);
  assert.match(ctx.mailTomorrow(opts)('en').html, /0 confirmed/);
  const maybe = ctx.mailMaybeMorning(opts)('en').html;
  assert.match(maybe, /join the waitlist/i); assert.doesNotMatch(maybe, /There's still room/);
  const morning = ctx.mailMatchday({ game, settings: ctx.getLeagueSettings(), weather: null })('en');
  assert.match(morning.subject, /Players needed/); assert.match(morning.html, /Awaiting re-confirmation/); assert.doesNotMatch(morning.html, /Today's match is on/);
});

test('an organizer nudge creates one personalized job per player, with no duplicate fallback email', () => {
  const f = fixture(), { ctx } = f; f.profile('alice@example.test', 'Email'); f.profile('bob@example.test', 'Both'); f.game();
  const game = { GameID: 'G_test', Title: 'Lunch soccer', Date: '2026-11-03', Time: '12:15 PM', LocationName: 'Usual field',
    yesCount: 0, confirmedCount: 0, MinThreshold: 2, MaxCap: 18, roster: [], locationDetails: {} };
  ctx.requireManagedGame = () => ({ manager: { email: 'admin@example.test', fullName: 'Organizer' }, gr: 0, g: ctx.gameRecord(0), gameId: 'G_test' });
  ctx.hasGameStartedOrPassed = () => false;
  ctx.getAppData = () => ({ games: [game], profiles: { 'alice@example.test': { fullName: 'Alice' }, 'bob@example.test': { fullName: 'Bob' } } });
  ctx.getLeagueSettings = () => ({ featWeather: false }); ctx.okData = (_who, message) => ({ success: true, message });
  const result = ctx.adminNudgeUnresponded({ gameId: 'G_test' }); assert.equal(result.success, true);
  assert.equal(f.notifications().length, 2); f.setPush({ code: 200, body: { id: '' } }); ctx.processNotificationOutbox();
  assert.equal(f.mail.length, 2); assert.equal(new Set(f.mail.map(m => m.to)).size, 2);
  assert.ok(f.mail.every(m => m.htmlBody.includes('sig=')));
});

test('regular players do not receive notification diagnostics or other admin-only system data', () => {
  const f = fixture(), { ctx } = f;
  ctx.getAppData = () => ({ success: true, profiles: {}, games: [], system: { notifications: { pending: 5 } }, invites: [], recurring: [], locations: [], settings: {}, howWePlayRules: '' });
  const data = ctx.getRoleFilteredAppData({ email: 'alice@example.test', isAdmin: false });
  assert.equal(data.system, undefined);
});
console.log(`\n${passed} backend tests passed against ${sourcePath}`);
