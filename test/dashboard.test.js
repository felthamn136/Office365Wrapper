'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { test } = require('node:test');

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
  const html = fs.readFileSync(path.join(dashboard, 'dashboard.html'), 'utf8');
  assert.match(html, /img-src 'self'/);
  assert.doesNotMatch(config.app.security.csp, /connect-src ipc:/);
  assert.doesNotMatch(html, /connect-src ipc:/);
});

test('dashboard no longer bundles workspace search', () => {
  const root = path.join(__dirname, '..');
  const html = fs.readFileSync(path.join(root, 'assets/dashboard/dashboard.html'), 'utf8');
  const config = JSON.parse(fs.readFileSync(path.join(root, 'src-tauri/tauri.conf.json'), 'utf8'));
  const cargo = fs.readFileSync(path.join(root, 'src-tauri/Cargo.toml'), 'utf8');
  assert.doesNotMatch(html, /search your workspace|search-form|search-results/i);
  assert.ok(!config.build.frontendDist.includes('../assets/dashboard/search.js'));
  assert.equal(config.app.withGlobalTauri, undefined);
  assert.doesNotMatch(cargo, /reqwest|serde_json|base64|sha2|rand/);
});
