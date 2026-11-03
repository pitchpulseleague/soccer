'use strict';
// Offline regression tests against the real inline app. No browser, Google or OneSignal calls.
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const root = path.resolve(__dirname, '..');
const sourcePath = process.argv[2] || (fs.existsSync(path.join(root, 'release/index.html')) ? path.join(root, 'release/index.html') : path.join(root, 'index.html'));
const html = fs.readFileSync(sourcePath, 'utf8');
const scripts = [...html.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script\s*>/gi)].map(m => m[1]);
scripts.forEach((s, i) => new vm.Script(s, {filename: 'inline-' + i}));
const main = scripts.find(s => s.includes('const I18N ='));
const bootBoundary = main.lastIndexOf('\n    applyThemeAndLanguageUI();');
assert.ok(bootBoundary > 0);
const code = main.slice(0, bootBoundary);

function harness() {
  const listeners = {}, nodes = {}, timers = [];
  const storage = new Map();
  let doc;
  function node(id = '') {
    const attrs = {};
    const n = {id, value: '', innerHTML: '', textContent: '', tagName: 'DIV', dataset: {}, style: {}, checked: false, disabled: false, inert: false, isConnected: true,
      classList: {add(){},remove(){},toggle(){},contains(){return false;}},
      setAttribute(k,v){attrs[k]=String(v);},removeAttribute(k){delete attrs[k];},getAttribute(k){return Object.hasOwn(attrs,k)?attrs[k]:null;},
      addEventListener(){},removeEventListener(){},querySelector(){return null;},querySelectorAll(){return [];},
      appendChild(){},remove(){n.isConnected=false;},contains(x){return x===n;},closest(){return null;},
      focus(){if(doc)doc.activeElement=n;n.focusCount=(n.focusCount||0)+1;},click(){n.clicked=true;},
      getClientRects(){return [{}];},getBoundingClientRect(){return {top:0,left:0,bottom:20,width:20,height:20};}};
    return n;
  }
  for (const id of ['view-container','modal-root','toast-root','app-header','app-main','mobile-bottom-dock','sheet-panel','sync-indicator','manual-refresh-btn']) nodes[id]=node(id);
  doc = {body:node('body'),documentElement:node('html'),activeElement:null,
    addEventListener(type, fn){(listeners[type] ||= []).push(fn);},removeEventListener(){},
    getElementById(id){return nodes[id] || null;},querySelector(){return null;},querySelectorAll(){return [];},createElement(){return node();}};
  const store = {getItem(k){return storage.has(k)?storage.get(k):null;},setItem(k,v){storage.set(k,String(v));},removeItem(k){storage.delete(k);}};
  const location={origin:'https://example.invalid',pathname:'/soccer/',search:'',href:'https://example.invalid/soccer/'};
  const sandbox={console,URL,URLSearchParams,Blob,Intl,Date,JSON,Math,Promise,AbortController,
    navigator:{language:'en',userAgent:'offline regression'},document:doc,location,localStorage:store,sessionStorage:store,
    history:{pushState(){},replaceState(){},back(){}},
    setTimeout(fn){timers.push(fn);return timers.length;},clearTimeout(){},setInterval(){return 0;},clearInterval(){},
    fetch(){throw new Error('Offline test: network is forbidden');},
    window:{location,matchMedia(){return {matches:false,addEventListener(){}};},addEventListener(){},scrollTo(){},innerWidth:390,innerHeight:844}};
  const ctx=vm.createContext(sandbox);
  const run=s=>vm.runInContext(s,ctx);
  run(code);
  run(`currentUser={email:'alex@example.invalid',fullName:'Alex Sample',pin:'SES_synthetic',isAdmin:true,position:'MID',notifyPref:'Both'};
    appData=normalizeAppData(JSON.parse(JSON.stringify(EMPTY_APP)));appData.leagueTimeZone='America/Los_Angeles';
    appData.locations=[{LocationID:'SAMPLE',Name:'Sample field',Address:'Synthetic address',LocationStatus:'Active'}];
    appData.games=[{GameID:'G_SAMPLE',Title:'Lunch soccer',Date:'2099-11-04',Time:'12:15 PM',LocationName:'Sample field',
      MaxCap:18,MinThreshold:14,Status:'Open',yesCount:14,confirmedCount:11,pendingReconfirmCount:3,
      startAtMs:Date.parse('2099-11-04T20:15:00Z'),endAtMs:Date.parse('2099-11-04T21:30:00Z'),ScoreWhite:'',ScoreDark:'',hasScore:false,
      rsvpLocked:false,hasStartedOrPassed:false,roster:[{email:'alex@example.invalid',fullName:'Alex Sample',status:'Yes',position:'MID',guestNames:['Guest One','Guest Two'],needsReconfirm:true}],
      features:{teams:false,scores:false,motm:false,attendance:false,lateDrops:true,arrival:false,captains:false,announcements:true,ball:false,pinnies:false,gk:false},blasts:[],teams:null,adminActivity:[]}];`);
  return {ctx,sandbox,run,nodes,node,doc,listeners,timers,storage};
}
let count=0;
async function test(name, fn){await fn();count++;console.log('PASS '+name);}

(async function(){
  await test('scores must be integers from 0 through 99 without silent truncation',()=>{
    const h=harness();
    for(const [input,expected] of [['0',0],['99',99],[' 12 ',12],['1.5',null],['100',null],['-1',null],['',null],['1e1',null],['NaN',null]]){
      h.sandbox.input=input;assert.equal(h.run('readScoreValue(input)'),expected,input);
    }
  });
  await test('Pacific wall-clock times convert correctly in winter and summer',()=>{
    const h=harness();
    assert.equal(h.run(`new Date(wallTimeToUtcMs('2026-11-04','12:15 PM','America/Los_Angeles')).toISOString()`),'2026-11-04T20:15:00.000Z');
    assert.equal(h.run(`new Date(wallTimeToUtcMs('2026-07-15','12:15 PM','America/Los_Angeles')).toISOString()`),'2026-07-15T19:15:00.000Z');
    assert.equal(h.run(`addDaysIso('2026-03-08',1)`),'2026-03-09');
    h.run(`appData.games[0].Date='2026-11-04';appData.games[0].startAtMs=Date.parse('2026-11-04T20:15:00Z');appData.games[0].endAtMs=Date.parse('2026-11-04T21:30:00Z');`);
    assert.match(h.run('buildIcsCalendar(appData.games[0])'),/DTSTART:20261104T201500Z\r\nDTEND:20261104T213000Z/);
    assert.match(h.run('buildGoogleCalendarUrl(appData.games[0])'),/ctz=America%2FLos_Angeles/);
    h.sandbox.icsValue='First\nSecond, third; slash\\';
    assert.equal(h.run('icsEscapeText(icsValue)'),'First\\nSecond\\, third\\; slash\\\\');
  });
  await test('reserved spots do not falsely confirm a schedule-changed match',()=>{
    const h=harness();
    assert.equal(h.run('gameMeta(appData.games[0]).yesCount'),14);
    assert.equal(h.run('gameMeta(appData.games[0]).confirmedCount'),11);
    assert.equal(h.run('gameMeta(appData.games[0]).isConfirmed'),false);
    assert.match(h.run('statusPill(appData.games[0],gameMeta(appData.games[0]))'),/Needs 3 more/);
    assert.match(h.run('myStatusSentence(gameMeta(appData.games[0]))'),/reserved.*confirm/i);
  });
  await test('small-sided approval requires ten reconfirmed players rather than reserved spots',()=>{
    const h=harness();h.run(`renderSheetModal=function(markup){globalThis.markup=markup;};appData.settings.featSmallSided=true;appData.games[0].confirmedCount=9;openOrganizerTools('G_SAMPLE');`);
    assert.match(h.sandbox.markup,/Needs at least 10 players/);assert.doesNotMatch(h.sandbox.markup,/onclick="confirmSmallSided/);
    h.run(`appData.games[0].confirmedCount=10;openOrganizerTools('G_SAMPLE');`);
    assert.match(h.sandbox.markup,/onclick="confirmSmallSided/);assert.match(h.sandbox.markup,/with the 10 players/);
  });
  await test('no attendance evidence produces no fabricated 100% rating',()=>{
    const h=harness();
    assert.match(h.run('reliabilityText({reliability:null,reliabilitySampleCount:0})'),/Not enough/);
    assert.match(h.run('reliabilityText({reliability:100,reliabilitySampleCount:0})'),/Not enough/);
    assert.match(h.run('reliabilityText({reliability:90,reliabilitySampleCount:10})'),/90%.*10 recorded/);
  });
  await test('new schedules default to weekdays at 12:15 PM',()=>{
    const h=harness();h.run('renderSheetModal=function(markup){globalThis.markup=markup;}');
    h.run(`openGameForm('')`);assert.match(h.sandbox.markup,/value="Lunch soccer"/);assert.match(h.sandbox.markup,/value="15" selected/);
    h.run(`openSeriesForm('')`);
    for(const day of ['1','2','3','4','5'])assert.match(h.sandbox.markup,new RegExp('name="rec-day" value="'+day+'" checked'));
    assert.doesNotMatch(h.sandbox.markup,/name="rec-day" value="[06]" checked/);
    assert.match(h.sandbox.markup,/value="15" selected/);
  });
  await test('changing appearance preserves unsaved profile fields',()=>{
    const h=harness();
    for(const [id,value] of [['prof-notify-pref','Email'],['prof-away-input','2099-12-12'],['prof-pos-select','DEF']]){h.nodes[id]=h.node(id);h.nodes[id].value=value;}
    h.nodes['prof-theme']=h.node('prof-theme');
    h.run(`openMyProfileModal=function(){throw new Error('Theme must not rebuild Profile');};setProfileTheme('dark');`);
    assert.equal(h.nodes['prof-notify-pref'].value,'Email');assert.equal(h.nodes['prof-away-input'].value,'2099-12-12');assert.equal(h.nodes['prof-pos-select'].value,'DEF');
    assert.match(h.nodes['prof-theme'].innerHTML,/aria-checked="true"/);
  });
  await test('only changed settings are dirty and saving one group leaves no phantom draft',()=>{
    const h=harness();h.run(`activeTab='admin';adminSection='settings';`);
    h.sandbox.target={id:'st-min',value:'12'};h.run('captureAdminDraft(target)');
    assert.equal(h.run('JSON.stringify(Object.keys(adminSettingsDraft))'),'["st-min"]');
    h.run(`clearAdminDraft('settings')`);assert.equal(h.run('hasUnsavedSettings()'),false);
    h.sandbox.target={id:'admin-rules-textarea',value:'Custom rules'};h.run('captureAdminDraft(target)');h.run(`clearAdminDraft('settings')`);
    assert.equal(h.run('JSON.stringify(Object.keys(adminSettingsDraft))'),'["admin-rules-textarea"]');
    h.sandbox.target={id:'admin-rules-textarea',value:''};h.run('captureAdminDraft(target)');assert.equal(h.run('hasUnsavedSettings()'),false);
  });
  await test('writes with different keys execute serially',async()=>{
    const h=harness();h.run(`var events=[];var releaseFirst;var first=runAction('first',async function(){events.push('first:start');await new Promise(function(resolve){releaseFirst=resolve;});events.push('first:end');});var second=runAction('second',async function(){events.push('second:start');});`);
    await Promise.resolve();assert.equal(h.run('JSON.stringify(events)'),'["first:start"]');
    h.run('releaseFirst()');await h.run('Promise.all([first,second])');
    assert.equal(h.run('JSON.stringify(events)'),'["first:start","first:end","second:start"]');
  });
  await test('queued work from a previous account is discarded',async()=>{
    const h=harness();h.run(`var ranOld=false;var releaseFirst;var first=runAction('first',async function(){await new Promise(function(resolve){releaseFirst=resolve;});});var second=runAction('second',async function(){ranOld=true;});`);
    await Promise.resolve();h.run('authEpoch++;releaseFirst()');await h.run('Promise.all([first,second])');assert.equal(h.run('ranOld'),false);
  });
  await test('a late authenticated response cannot restore the previous account',async()=>{
    const h=harness();let finish;
    h.sandbox.fetch=()=>new Promise(resolve=>{finish=resolve;});
    const request=h.run(`apiPost({action:'getLeagueData'},{silent:true})`);
    h.run(`authEpoch++;currentUser={email:'new@example.invalid',pin:'SES_new'};`);
    finish({text:async()=>JSON.stringify({success:true,user:{email:'alex@example.invalid',pin:'SES_old'}})});
    const response=await request;assert.equal(response.staleAuth,true);assert.equal(h.run('currentUser.email'),'new@example.invalid');
    h.sandbox.old={_clientAuthEpoch:0,user:{email:'alex@example.invalid'}};assert.equal(h.run('setAppData(old)'),false);
  });
  await test('a signed email-answer timeout is uncertain even without login',async()=>{
    const h=harness();h.run('currentUser=null;');h.sandbox.fetch=async()=>{const e=new Error('timeout');e.name='AbortError';throw e;};
    const r=await h.run(`apiPost({action:'emailAnswer',email:'player@example.invalid',gameId:'G_SAMPLE',status:'Yes',sig:'test-only'},{noAuth:true,silent:true})`);
    assert.equal(r.uncertain,true);assert.doesNotMatch(r.error,/wasn't saved/i);
  });
  await test('signed email links preview only until explicit confirmation',async()=>{
    const h=harness();h.run(`currentUser=null;var actions=[];renderSheetModal=function(markup){globalThis.markup=markup;};
      apiPost=async function(p){actions.push(p.action);return {success:true,email:'player@example.invalid',playerName:'Player Sample',currentStatus:'Maybe',finalStatus:'Yes',guestCount:0,game:{GameID:'G_SAMPLE',Title:'Sample match',Date:'2099-11-04',Time:'12:15 PM',LocationName:'Sample field',yesCount:14,confirmedCount:14,MinThreshold:14,MaxCap:18}};};
      pendingEmailAnswer={email:'player@example.invalid',gameId:'G_SAMPLE',status:'Yes',sig:'test-only'};`);
    await h.run('loadEmailAnswerPreview()');assert.equal(h.run('JSON.stringify(actions)'),'["previewEmailAnswer"]');
    assert.match(h.sandbox.markup,/Review your choice/);assert.match(h.sandbox.markup,/id="ea-confirm"/);
    await h.run('submitEmailAnswer()');assert.equal(h.run('JSON.stringify(actions)'),'["previewEmailAnswer","emailAnswer"]');
    assert.match(h.sandbox.markup,/showEmailAnswerConfirmation/);
    assert.doesNotMatch(main.slice(bootBoundary),/submitEmailAnswer\(\)/);
  });
  await test('unsigned legacy links also wait for explicit confirmation after login',()=>{
    const h=harness();
    h.sandbox.setTimeout=fn=>{fn();return 0;};
    h.run(`var writes=0;renderCurrentTab=function(){};confirmDialog=function(opts){globalThis.confirmation=opts;};closeModal=function(){};onRsvpTap=function(){writes++;};pendingEmailRsvp={email:'alex@example.invalid',gameId:'G_SAMPLE',status:'Yes'};applyEmailRsvpIfReady();`);
    assert.equal(h.run('writes'),0);assert.ok(h.sandbox.confirmation);h.sandbox.confirmation.onOk();assert.equal(h.run('writes'),1);
  });
  await test('invalid signed links offer app access without a retry-only dead end',async()=>{
    const h=harness();h.run(`currentUser=null;var actions=[];renderSheetModal=function(markup){globalThis.markup=markup;};apiPost=async function(p){actions.push(p.action);return {success:false,badLink:true};};pendingEmailAnswer={email:'player@example.invalid',gameId:'G_SAMPLE',status:'Yes',sig:'invalid-test-only'};`);
    await h.run('loadEmailAnswerPreview()');
    assert.equal(h.run('JSON.stringify(actions)'),'["previewEmailAnswer"]');
    assert.match(h.sandbox.markup,/showLoginModal/);assert.match(h.sandbox.markup,/>Close</);assert.doesNotMatch(h.sandbox.markup,/onclick="loadEmailAnswerPreview/);
  });
  await test('changing an email answer reloads its warning and ignores a superseded preview',async()=>{
    const h=harness();h.run(`currentUser=null;var requests=[];renderSheetModal=function(markup){globalThis.markup=markup;};apiPost=function(p){return new Promise(function(resolve){requests.push({payload:p,resolve:resolve});});};pendingEmailAnswer={email:'player@example.invalid',gameId:'G_SAMPLE',status:'Yes',sig:'test-only',preview:{currentStatus:'Yes',lateDropApplies:false,game:appData.games[0]}};showEmailAnswerConfirmation('No');showEmailAnswerConfirmation('Maybe');`);
    assert.equal(h.run('JSON.stringify(requests.map(function(r){return r.payload.status;}))'),'["No","Maybe"]');
    h.run(`requests[1].resolve({success:true,currentStatus:'Yes',lateDropApplies:true,playerName:'Fresh preview',game:appData.games[0]});`);await new Promise(resolve=>setImmediate(resolve));
    assert.match(h.sandbox.markup,/Fresh preview/);assert.match(h.sandbox.markup,/late/i);const newest=h.sandbox.markup;
    h.run(`requests[0].resolve({success:true,currentStatus:'No',lateDropApplies:false,playerName:'Stale preview',game:appData.games[0]});`);await new Promise(resolve=>setImmediate(resolve));
    assert.equal(h.sandbox.markup,newest);assert.doesNotMatch(h.sandbox.markup,/Stale preview/);
  });
  await test('a timed-out reconfirmation stays uncertain while the saved Yes still needs reconfirming',async()=>{
    const h=harness();h.run(`currentUser=null;var actions=[],toasts=[];renderSheetModal=function(markup){globalThis.markup=markup;};showToast=function(message){toasts.push(message);};apiPost=async function(p){actions.push(p.action);return p.action==='emailAnswer'?{success:false,uncertain:true}:{success:true,currentStatus:'Yes',needsReconfirm:true,playerName:'Player Sample',game:appData.games[0]};};pendingEmailAnswer={email:'player@example.invalid',gameId:'G_SAMPLE',status:'Yes',sig:'test-only',preview:{currentStatus:'Yes',needsReconfirm:true,game:appData.games[0]}};`);
    await h.run('submitEmailAnswer()');
    assert.equal(h.run('JSON.stringify(actions)'),'["emailAnswer","previewEmailAnswer"]');assert.equal(h.run('toasts[toasts.length-1]===t("saveUncertain")'),true);assert.equal(h.run('!!pendingEmailAnswer.result'),false);
  });
  await test('attendance marks presence and keeps corrections visible after RSVP changes',()=>{
    const h=harness();h.run(`renderSheetModal=function(markup){globalThis.markup=markup;};appData.games[0].roster.push({email:'present@example.invalid',fullName:'Recorded Present',status:'No',attendanceStatus:'Attended',guestNames:[]},{email:'absent@example.invalid',fullName:'Recorded Absent',status:'No',attendanceStatus:'NoShow',guestNames:[]});openAttendanceSheet('G_SAMPLE');`);
    assert.match(h.sandbox.markup,/Mark present/);assert.match(h.sandbox.markup,/Recorded Present/);assert.match(h.sandbox.markup,/Recorded Absent/);assert.match(h.sandbox.markup,/Mark present instead/);
    assert.match(h.sandbox.markup,/<button type="button" disabled onclick="confirmNoShow/);
    h.run(`hasKickedOff=function(){return true;};openAttendanceSheet('G_SAMPLE');`);assert.doesNotMatch(h.sandbox.markup,/<button type="button" disabled onclick="confirmNoShow/);
    h.run(`var attendancePayload;mutate=function(key,payload){attendancePayload=payload;};submitAttendanceChange('G_SAMPLE','absent@example.invalid','Yes','');`);
    assert.equal(h.run('attendancePayload.action'),'adminUpdateAttendance');assert.equal(h.run('attendancePayload.newStatus'),'Yes');
  });
  await test('future series edits keep their effective date and show scheduled changes',()=>{
    const h=harness();h.run(`getTodayIso=function(){return '2026-11-03';};renderSheetModal=function(markup){globalThis.markup=markup;};appData.recurring=[{SeriesID:'SERIES_SAMPLE',Title:'Lunch soccer',LocationName:'Sample field',DaysOfWeek:'1,2,3,4,5',Time:'12:45 PM',StartDate:'2026-11-03',EndDate:'2027-02-03',MinThreshold:14,MaxCap:18,scheduledChanges:[{effectiveDate:'2026-12-01',config:{DaysOfWeek:'1,2,3,4,5',Time:'12:45 PM',LocationName:'Sample field'}},{effectiveDate:'2026-11-15',config:{DaysOfWeek:'1,3,5',Time:'12:30 PM',LocationName:'Sample field'}}]}];openSeriesForm('SERIES_SAMPLE');`);
    assert.equal(h.run('seriesEditEffectiveDate(appData.recurring[0])'),'2026-12-01');assert.match(h.sandbox.markup,/id="rec-effective"[^>]+value="2026-12-01"/);assert.match(h.sandbox.markup,/replaces later scheduled changes/);
    assert.equal((h.run('renderAdminSeries()').match(/Scheduled from/g)||[]).length,2);
    h.run('appData.recurring[0].scheduledChanges=[];');assert.equal(h.run('seriesEditEffectiveDate(appData.recurring[0])'),'2026-11-03');
  });
  await test('existing guests can be removed when new guests are disabled',async()=>{
    const h=harness();h.run(`appData.settings.maxGuests=0;currentUser.canBringGuests=false;renderSheetModal=function(markup){globalThis.markup=markup;};openMyMatchSheet('G_SAMPLE');`);
    assert.match(h.sandbox.markup,/name="mm-keep-guest"/);assert.match(h.sandbox.markup,/New guests are disabled/);assert.match(h.sandbox.markup,/Remove all guests/);assert.doesNotMatch(h.sandbox.markup,/id="mm-guest-/);
    h.doc.querySelectorAll=selector=>selector==='input[name="mm-keep-guest"]'?[{value:'0',checked:true},{value:'1',checked:false}]:[];
    h.run(`var sent=[];apiPost=async function(p){sent.push(p);return Object.assign({},appData,{success:true,user:currentUser});};closeModal=function(){};renderCurrentTab=function(){};`);
    await h.run(`saveMyMatchSheet('G_SAMPLE')`);
    assert.equal(h.run('JSON.stringify(sent[0].value)'),'["Guest One"]');
  });
  await test('modal has a name, isolates background, traps focus and restores it',()=>{
    const h=harness();const dialog=h.node('dialog'),heading=h.node('heading'),trigger=h.node('trigger'),first=h.node('first'),last=h.node('last'),panel=h.nodes['sheet-panel'];
    h.nodes['modal-root'].querySelector=selector=>selector==='[role="dialog"]'?dialog:selector==='h1,h2,h3'?heading:null;
    h.nodes['modal-root'].contains=x=>[panel,first,last].includes(x);panel.contains=x=>[panel,first,last].includes(x);panel.querySelectorAll=()=>[first,last];
    h.doc.activeElement=trigger;h.run(`renderSheetModal('<h3>Test dialog</h3>')`);
    assert.equal(dialog.getAttribute('aria-labelledby'),'sheet-dialog-title');assert.equal(h.nodes['app-main'].inert,true);
    h.doc.activeElement=last;let prevented=false;
    for(const fn of h.listeners.keydown||[])fn({key:'Tab',shiftKey:false,target:last,preventDefault(){prevented=true;}});
    assert.equal(prevented,true);assert.equal(h.doc.activeElement,first);
    h.run('closeModal(true)');assert.equal(h.nodes['app-main'].inert,false);assert.equal(h.doc.activeElement,trigger);
  });
  await test('finished cards follow future matches but same-day answers remain available',()=>{
    const h=harness();h.run(`getTodayIso=function(){return '2026-11-03';};leagueNowMs=function(){return Date.parse('2026-11-03T23:00:00Z');};
      const future=appData.games[0];future.GameID='FUTURE';future.Date='2026-11-04';future.startAtMs=Date.parse('2026-11-04T20:15:00Z');future.endAtMs=Date.parse('2026-11-04T21:30:00Z');
      const old=JSON.parse(JSON.stringify(future));old.GameID='TODAY';old.Date='2026-11-03';old.startAtMs=Date.parse('2026-11-03T20:15:00Z');old.endAtMs=Date.parse('2026-11-03T21:30:00Z');old.rsvpLocked=false;appData.games=[old,future];`);
    assert.equal(h.run('upcomingGames()[0].GameID'),'FUTURE');
    assert.match(h.run(`renderRsvpBlock(findGame('TODAY'),gameMeta(findGame('TODAY')))`),/Correct your answer for today/);
  });
  console.log('\n'+count+' frontend regression tests passed. No network calls were made.');
})().catch(error=>{console.error(error);process.exitCode=1;});
