'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { test } = require('node:test');
const { destinations } = require('../destinations');

test('dashboard links open the supported Microsoft web apps', () => {
  const html = fs.readFileSync(path.join(__dirname, '../dashboard.html'), 'utf8');
  const links = [...html.matchAll(/<a\s+[^>]*href="([^"]+)"/g)].map((match) => match[1]);
  assert.deepEqual(links, [
    destinations.microsoft365.url,
    destinations.outlook.url,
    destinations.outlookPersonal.url,
    'https://go.microsoft.com/fwlink/p/?LinkID=2119709',
    'https://word.office.com/',
    'https://excel.office.com/',
    'https://powerpoint.office.com/',
    'https://onenote.office.com/',
    'https://copilot.cloud.microsoft/',
    'https://to-do.office.com/',
  ]);
});
