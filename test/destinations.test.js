'use strict';

const assert = require('node:assert/strict');
const { test } = require('node:test');
const { destinationFromArgs } = require('../destinations');

test('selects the dashboard by default and supports direct app destinations', () => {
  assert.equal(destinationFromArgs([]).file, 'dashboard.html');
  assert.equal(destinationFromArgs(['--microsoft365']).url, 'https://www.microsoft365.com/');
  assert.equal(destinationFromArgs(['--outlook']).url, 'https://outlook.office365.com/');
  assert.equal(destinationFromArgs(['--outlook-personal']).url, 'https://outlook.live.com/mail/');
});

test('rejects unknown and multiple arguments', () => {
  assert.throws(() => destinationFromArgs(['--other']), /Unknown argument/);
  assert.throws(
    () => destinationFromArgs(['--outlook', '--outlook-personal']),
    /Only one destination/,
  );
});

test('rejects the removed Teams launch option', () => {
  assert.throws(() => destinationFromArgs(['--teams']), /Unknown argument/);
});
