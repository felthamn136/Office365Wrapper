'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { test } = require('node:test');
const vm = require('node:vm');

test('dashboard links open the supported Microsoft web apps', () => {
  const html = fs.readFileSync(path.join(__dirname, '../assets/dashboard/dashboard.html'), 'utf8');
  const links = [...html.matchAll(/<a\s+[^>]*href="([^"]+)"/g)].map((match) => match[1]);
  assert.deepEqual(links, [
    'https://www.microsoft365.com/',
    'https://outlook.office365.com/',
    'https://outlook.live.com/mail/',
    'https://go.microsoft.com/fwlink/p/?LinkID=2119709',
    'https://word.office.com/',
    'https://excel.office.com/',
    'https://powerpoint.office.com/',
    'https://onenote.office.com/',
    'https://copilot.cloud.microsoft/',
    'https://to-do.office.com/',
  ]);
});

test('the watermark is available to the dashboard and bundled by Tauri', () => {
  const root = path.join(__dirname, '..');
  const dashboard = path.join(root, 'assets/dashboard');
  const css = fs.readFileSync(path.join(dashboard, 'dashboard.css'), 'utf8');
  const image = css.match(/url\(["']?([^"')]+)["']?\)/)?.[1];
  assert.ok(image, 'dashboard CSS references a watermark image');
  assert.ok(fs.statSync(path.join(dashboard, image)).isFile());
  const config = JSON.parse(fs.readFileSync(path.join(root, 'src-tauri/tauri.conf.json'), 'utf8'));
  assert.ok(config.build.frontendDist.includes(`../assets/dashboard/${image}`));
  assert.match(config.app.security.csp, /img-src 'self'/);
  assert.match(fs.readFileSync(path.join(dashboard, 'dashboard.html'), 'utf8'), /img-src 'self'/);
});

test('Graph search uses Tauri IPC and is disabled outside Tauri', async () => {
  const script = fs.readFileSync(path.join(__dirname, '../assets/dashboard/search.js'), 'utf8');
  const elements = Object.fromEntries(
    ['search-form', 'search-query', 'search-submit', 'search-signout', 'search-status', 'search-results']
      .map((id) => [id, {
        addEventListener(type, handler) { this[type] = handler; },
        replaceChildren() {},
      }]),
  );
  const calls = [];
  const document = { getElementById: (id) => elements[id] };
  const obsoleteBridge = { search() { throw new Error('Electron bridge must not be used'); } };
  vm.runInNewContext(script, {
    document,
    window: { graphSearch: obsoleteBridge, __TAURI__: { core: { invoke: async (...args) => {
      calls.push(args);
      return [];
    } } } },
  });
  elements['search-query'].value = 'project';
  await elements['search-form'].submit({ preventDefault() {} });
  await elements['search-signout'].click();
  assert.deepEqual(JSON.parse(JSON.stringify(calls)), [
    ['graph_search', { query: 'project' }],
    ['graph_sign_out'],
  ]);
  vm.runInNewContext(script, { document, window: { graphSearch: obsoleteBridge } });
  assert.equal(elements['search-submit'].disabled, true);
  assert.equal(elements['search-signout'].disabled, true);
  assert.equal(elements['search-status'].textContent, 'Graph search is available in the desktop app.');
});
