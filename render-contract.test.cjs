'use strict';
// Exercises the actual page renderers and parses their HTML without a browser or network.
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const cp = require('node:child_process');
const assert = require('node:assert/strict');
const root = path.resolve(__dirname, '..');
const releaseDir = fs.existsSync(path.join(root, 'release/index.html')) ? path.join(root, 'release') : root;
const html = fs.readFileSync(path.join(releaseDir, 'index.html'), 'utf8');
let code = [...html.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script\s*>/gi)]
  .map(m => m[1]).find(s => s.includes('const I18N ='));
assert.ok(code, 'main app script found');
const boot = code.lastIndexOf('\n    applyThemeAndLanguageUI();');
assert.ok(boot > 0, 'explicit boot boundary found');
code = code.slice(0, boot);
function el(id = '') {
  return { id, value: '', innerHTML: '', textContent: '', dataset: {}, style: {}, checked: false,
    classList: {add(){}, remove(){}, toggle(){}, contains(){return false;}},
    setAttribute(){}, removeAttribute(){}, getAttribute(){return null;},
    addEventListener(){}, removeEventListener(){}, querySelector(){return null;}, querySelectorAll(){return [];},
    appendChild(){}, remove(){}, contains(){return false;}, focus(){}, getBoundingClientRect(){return {top:0,left:0,bottom:10,width:10,height:10};}
  };
}
const elements = Object.fromEntries(['view-container', 'modal-root', 'toast-root', 'app-header', 'mobile-bottom-dock'].map(id => [id, el(id)]));
const location = {origin:'https://example.invalid',pathname:'/soccer/',search:'',href:'https://example.invalid/soccer/'};
const document = {body:el('body'),documentElement:el('html'),activeElement:null,
  addEventListener(){},removeEventListener(){},getElementById(id){return elements[id] || null;},
  querySelector(){return null;},querySelectorAll(){return [];},createElement(){return el();}};
const store = {getItem(){return null;},setItem(){},removeItem(){}};
const sandbox = {console,URL,URLSearchParams,Blob,Intl,Date,JSON,Math,Promise,AbortController,
  navigator:{language:'en',userAgent:'offline regression harness'},document,location,
  localStorage:store,sessionStorage:store,history:{pushState(){},replaceState(){},back(){}},
  setTimeout(){return 0;},clearTimeout(){},setInterval(){return 0;},clearInterval(){},
  fetch(){throw new Error('Network is forbidden in offline tests');},
  requestAnimationFrame(){return 0;},cancelAnimationFrame(){},
  window:{location,matchMedia(){return {matches:false,addEventListener(){}};},addEventListener(){},removeEventListener(){},scrollTo(){},innerWidth:390,innerHeight:844}};
const ctx = vm.createContext(sandbox);
vm.runInContext(code, ctx, {filename:'index.html'});
const run = (s) => vm.runInContext(s, ctx);
let checks = 0;
function check(name, fn) {fn(); ++checks;console.log('PASS ' + name);}
function parsedAttributes(fragments) {
  const python = [
    'import sys,json',
    'from html.parser import HTMLParser',
    'class Parser(HTMLParser):',
    ' def __init__(self): super().__init__(convert_charrefs=True); self.handlers=[]; self.ids=[]',
    ' def handle_starttag(self,tag,attrs):',
    '  for key,value in attrs:',
    '   if key.startswith("on"): self.handlers.append(value or "")',
    '   if key=="id": self.ids.append(value)',
    'out=[]',
    'for h in json.load(sys.stdin):',
    ' p=Parser(); p.feed(h); out.append({"handlers":p.handlers,"ids":p.ids})',
    'print(json.dumps(out))'
  ].join('\n');
  const p = cp.spawnSync('python3', ['-c',python],{input:JSON.stringify(fragments),encoding:'utf8'});
  assert.equal(p.status,0,p.stderr);
  return JSON.parse(p.stdout);
}
check('English and Spanish dictionaries have matching keys', () => {
  const en = JSON.parse(run('JSON.stringify(Object.keys(I18N.en).sort())'));
  const es = JSON.parse(run('JSON.stringify(Object.keys(I18N.es).sort())'));
  assert.deepEqual(es,en);
});
check('inline handler escaping survives HTML entity decoding', () => {
  const values = ["O'Neil & Sons", 'Literal &#39; apostrophe', 'Literal &#x27; entity',
    '<img src=x onerror=probe()>', "\\';probe();//", 'Double " quote &quot;', 'Ángel Pérez', '` ${probe()}`'];
  const fragments = values.map(value => {
    sandbox.inputValue = value;
    return '<button onclick="receive(\'' + run('escJsArg(inputValue)') + '\')">Test</button>';
  });
  const parsed = parsedAttributes(fragments);
  parsed.forEach((p,i) => {
    assert.equal(p.handlers.length,1);
    let received, probed = false;
    vm.runInNewContext(p.handlers[0], {receive(v){received=v;},probe(){probed=true;}});
    assert.equal(probed,false);
    assert.equal(received,values[i]);
  });
});
run(`
  currentUser = {email:'admin@example.invalid',fullName:'Admin Sample',pin:'SES_synthetic',isAdmin:true,position:'MID',notifyPref:'Both'};
  appData = normalizeAppData(JSON.parse(JSON.stringify(EMPTY_APP)));
  appData.leagueTimeZone = 'America/Los_Angeles';
  appData.settings.featTeams = true;
  appData.settings.featScores = true;
  appData.settings.featMotm = true;
  appData.settings.featAttendance = true;
  appData.settings.featLateDrops = true;
  appData.settings.featArrival = true;
  appData.settings.featCaptains = true;
  appData.locations = [{LocationID:'LOC_SAMPLE',Name:'Sample field',Address:'Synthetic address',LocationStatus:'Active'}];
  appData.profiles = {
    'admin@example.invalid':{email:'admin@example.invalid',fullName:'Admin Sample',role:'Admin',position:'MID',status:'Active',notifyPref:'Both',reliability:null,reliabilitySampleCount:0},
    'player@example.invalid':{email:'player@example.invalid',fullName:'Entity &#39; sample',role:'Player',position:'DEF',status:'Active',notifyPref:'Email',reliability:null,reliabilitySampleCount:0,canBringGuests:true}
  };
  appData.system = {emailQuotaLeft:20,notifications:{pending:2,accepted:3,needsReview:1,expired:0,lastRunAt:'',lastError:''}};
  appData.games = [{GameID:'G_SAMPLE',Title:'Lunch soccer',Date:'2099-11-04',Time:'12:15 PM',
    LocationName:'Sample field',MaxCap:4,MinThreshold:2,Status:'Pending',confirmedCount:1,pendingReconfirmCount:1,yesCount:2,
    startAtMs:Date.parse('2099-11-04T20:15:00Z'),endAtMs:Date.parse('2099-11-04T21:30:00Z'),
    ScoreWhite:'',ScoreDark:'',hasScore:false,hasStartedOrPassed:false,ScheduleChangeAt:1,
    locationDetails:{name:'Sample field',address:'Synthetic address',lat:'',lng:'',notes:''},
    roster:[{email:'admin@example.invalid',fullName:'Admin Sample',status:'Yes',position:'MID',guestNames:[],needsReconfirm:false},
      {email:'player@example.invalid',fullName:'Entity &#39; sample',status:'Yes',position:'DEF',guestNames:[],needsReconfirm:true}],
    blasts:[],teams:{white:[],dark:[]},adminActivity:[]}];
  renderSheetModal = function(innerHtml){ globalThis.capturedSheet = innerHtml; };
  closeModal = function(){};
`);
check('core renderers produce valid event handlers in both languages', () => {
  const fragments = [];
  for (const lang of ['en','es']) {
    sandbox.testLang = lang;
    run('currentLang=testLang');
    fragments.push(run('renderGameCard(appData.games[0], true)'));
    fragments.push(run('renderCompactGameCard(appData.games[0])'));
    fragments.push(run('renderAdminPlayers()'));
    fragments.push(run('renderAdminSettings()'));
    fragments.push(run('renderAdminFeatures()'));
    for (const expression of ["openPlayerAdminSheet('player@example.invalid')",'openMyProfileModal()',"openGameForm('')","openSeriesForm('')","openTeamsSheet('G_SAMPLE')","openScoreSheet('G_SAMPLE')","openAttendanceSheet('G_SAMPLE')","openMyMatchSheet('G_SAMPLE')"]) {
      run(expression); fragments.push(sandbox.capturedSheet);
    }
  }
  assert.ok(fragments.every(f => typeof f === 'string' && f.length > 0));
  const parsed = parsedAttributes(fragments);
  let handlerCount = 0;
  parsed.forEach((f,i) => {
    f.handlers.forEach(h => { new vm.Script('(function(event){' + h + '\n})', {filename:'fragment-'+i});++handlerCount; });
    assert.equal(new Set(f.ids).size, f.ids.length, 'unique IDs in fragment ' + i);
  });
  assert.ok(handlerCount > 100, 'representative interactive coverage');
  console.log('  Parsed ' + fragments.length + ' fragments and ' + handlerCount + ' event handlers.');
});
check('under-minimum match does not claim game on after kickoff', () => {
  run(`appData.games[0].startAtMs=Date.now()-10*60000; appData.games[0].endAtMs=Date.now()+65*60000; currentLang='en';`);
  const pill = run('statusPill(appData.games[0],gameMeta(appData.games[0]))');
  assert.doesNotMatch(pill, /Playing now|Game on|Played/i);
});
console.log('\n' + checks + ' render contract checks passed.');
