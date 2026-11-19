'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { test } = require('node:test');
const vm = require('node:vm');

function createHarness(fetchImpl) {
  const calls = { interactive: 0, silent: 0, removed: 0 };
  class InteractionRequiredAuthError extends Error {}
  class PublicClientApplication {
    async acquireTokenInteractive(request) {
      calls.interactive++;
      calls.scopes = request.scopes;
      await request.openBrowser('https://login.microsoftonline.com/organizations/oauth2/v2.0/authorize');
      return { account: { id: 'account' }, accessToken: 'test-token' };
    }

    async acquireTokenSilent() {
      calls.silent++;
      return { accessToken: 'test-token' };
    }

    getTokenCache() {
      return { removeAccount: async () => { calls.removed++; } };
    }
  }
  const sandbox = {
    require: () => ({ PublicClientApplication, InteractionRequiredAuthError }),
    module: { exports: {} },
    URL,
    AbortSignal,
    fetch: fetchImpl,
  };
  vm.runInNewContext(fs.readFileSync(path.join(__dirname, '../graph-search.js'), 'utf8'), sandbox);
  const create = (clientId) => sandbox.module.exports.createGraphSearch({
    clientId,
    openBrowser: async () => {},
    fetchImpl,
  });
  return { calls, create };
}

test('searches each Graph category and reports partial failures without exposing unsafe links', async () => {
  const requests = [];
  const fetchImpl = async (url, options) => {
    const request = JSON.parse(options.body).requests[0];
    requests.push({ url, options, request });
    if (request.entityTypes[0] === 'event') return { ok: false, status: 403 };
    const resource = request.entityTypes[0] === 'message'
      ? { subject: 'A message', webLink: 'javascript:alert(1)' }
      : { name: 'A file', webUrl: 'https://example.sharepoint.com/file' };
    return {
      ok: true,
      json: async () => ({ value: [{ hitsContainers: [{ hits: [{ resource }], moreResultsAvailable: true }] }] }),
    };
  };
  const { calls, create } = createHarness(fetchImpl);
  const graph = create('12345678-1234-1234-1234-123456789abc');
  const results = await graph.search('  report  ');
  assert.equal(calls.interactive, 1);
  assert.deepEqual([...calls.scopes], ['Mail.Read', 'Calendars.Read', 'Files.Read.All']);
  assert.equal(requests.length, 3);
  for (const { url, options, request } of requests) {
    assert.equal(url, 'https://graph.microsoft.com/v1.0/search/query');
    assert.equal(options.headers.Authorization, 'Bearer test-token');
    assert.equal(request.query.queryString, 'report');
    assert.equal(request.size, 25);
  }
  assert.equal(results[0].results[0].url, null);
  assert.equal(results[1].error, 'Microsoft Graph returned HTTP 403.');
  assert.equal(results[2].results[0].url, 'https://example.sharepoint.com/file');
  assert.equal(results[2].more, true);
  await graph.search('again');
  assert.equal(calls.silent, 1);
  await graph.signOut();
  assert.equal(calls.removed, 1);
  await graph.search('again');
  assert.equal(calls.interactive, 2);
});

test('rejects invalid queries and missing app registration before authentication', async () => {
  const { calls, create } = createHarness(async () => { throw new Error('Unexpected request'); });
  const graph = create('');
  await assert.rejects(graph.search('  '), /search term/);
  await assert.rejects(graph.search('x'.repeat(201)), /search term/);
  await assert.rejects(graph.search('hello'), /OFFICE365_CLIENT_ID/);
  assert.equal(calls.interactive, 0);
});
