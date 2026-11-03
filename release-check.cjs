'use strict';
// Offline release checks. Does not connect to the league or Google services.
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const root = path.resolve(__dirname, '..');
const releaseDir = fs.existsSync(path.join(root, 'release/index.html')) ? path.join(root, 'release') : root;
const html = fs.readFileSync(path.join(releaseDir, 'index.html'), 'utf8');
const backend = fs.readFileSync(path.join(releaseDir, 'Code.gs'), 'utf8');
let checks = 0;
function check(name, fn) { fn(); ++checks; console.log('PASS ' + name); }
const scripts = [...html.matchAll(/<script\b([^>]*)>([\s\S]*?)<\/script\s*>/gi)]
  .filter(m => !/\bsrc\s*=/.test(m[1]));
check('all frontend inline scripts parse', () => {
  assert.ok(scripts.length >= 4);
  scripts.forEach((m, i) => new vm.Script(m[2], { filename: 'index.html#script-' + (i + 1) }));
});
check('Apps Script JavaScript parses', () => { new vm.Script(backend, { filename: 'Code.gs' }); });
check('no merge conflict markers', () => {
  assert.doesNotMatch(html + '\n' + backend, /^(?:<<<<<<< |=======\s*$|>>>>>>> )/m);
});
check('both runtime versions are 5.4.0', () => {
  assert.match(html, /const APP_VERSION = ["']5\.4\.0["']/);
  assert.match(backend, /var APP_VERSION = ["']5\.4\.0["']/);
});
check('original deployment and branded assets retained', () => {
  assert.match(html, /const API_URL = "https:\/\/script\.google\.com\/macros\/s\/AKfycbxD2w-OIQ6k59BqRc0jbccjq09qYsKADavM4id6o02hVTKrqdV90j1bnjLb_Fp6hEc2\/exec"/);
  assert.match(html, /pitch-pulse-01-classic-stitched-192\.png/);
  assert.match(html, /OneSignalSDK\.page\.js/);
});
check('all API route handlers are defined', () => {
  const routeBlock = backend.match(/var routes = \{([\s\S]*?)\n\s*\};/);
  assert.ok(routeBlock, 'router not found');
  const names = [...routeBlock[1].matchAll(/^\s*\w+\s*:\s*(\w+)\s*,?\s*$/gm)].map(m => m[1]);
  assert.ok(names.length >= 35);
  names.forEach(n => assert.match(backend, new RegExp('function ' + n + '\\s*\\('), n));
});
check('every frontend API action is supported by the backend', () => {
  const routeBlock = backend.match(/var routes = \{([\s\S]*?)\n\s*\};/);
  const routes = new Set([...routeBlock[1].matchAll(/^\s*(\w+)\s*:/gm)].map(m => m[1]));
  const actions = new Set([...html.matchAll(/\baction\s*:\s*["'](\w+)["']/g)].map(m => m[1]));
  assert.ok(actions.size >= 40);
  for (const action of actions) assert.ok(routes.has(action), 'Unsupported frontend action: ' + action);
});
check('no accidental duplicate global backend function definitions', () => {
  const names = [...backend.matchAll(/^function\s+(\w+)\s*\(/gm)].map(m => m[1]);
  assert.equal(new Set(names).size, names.length);
});
check('HTML shell remains complete', () => {
  assert.match(html, /^<!DOCTYPE html>/i);
  assert.match(html, /<html\b[^>]*lang="en"/);
  assert.match(html, /<\/body>\s*<\/html>\s*$/i);
  for (const id of ['app-header', 'view-container', 'modal-root', 'toast-root', 'mobile-bottom-dock']) {
    const n = [...html.matchAll(new RegExp('\\bid="' + id + '"', 'g'))].length;
    assert.equal(n, 1, 'unique shell ID: ' + id);
  }
});
console.log('\n' + checks + ' release checks passed.');
