'use strict';

const assert = require('node:assert/strict');
const { test } = require('node:test');
const { configureNotifications } = require('../notifications');

test('notification permission is opt-in and limited to trusted origins', () => {
  const handlers = {};
  const session = {
    setPermissionRequestHandler: (handler) => { handlers.request = handler; },
    setPermissionCheckHandler: (handler) => { handlers.check = handler; },
  };
  let enabled = false;
  configureNotifications(session, () => enabled);
  const origin = 'https://outlook.office.com';
  const details = { isMainFrame: true, requestingUrl: `${origin}/mail/` };
  const request = (permission, requestDetails) => {
    let result;
    handlers.request(null, permission, (allowed) => { result = allowed; }, requestDetails);
    return result;
  };

  assert.equal(request('notifications', details), false);
  assert.equal(handlers.check(null, 'notifications', origin, details), false);
  enabled = true;
  assert.equal(request('notifications', details), true);
  assert.equal(handlers.check(null, 'notifications', origin, details), true);
  assert.equal(handlers.check(null, 'notifications', 'https://teams.microsoft.com', {}), false);
  assert.equal(handlers.check(null, 'notifications', 'https://teams.cloud.microsoft', {}), false);
  assert.equal(request('media', details), false);
  assert.equal(request('notifications', { ...details, isMainFrame: false }), false);
  assert.equal(request('notifications', { ...details, requestingUrl: 'https://outlook.office.com.evil.example/' }), false);
  assert.equal(request('notifications', { ...details, requestingUrl: 'http://outlook.office.com/' }), false);
  assert.equal(request('notifications', { ...details, requestingUrl: 'invalid' }), false);
  assert.equal(handlers.check(null, 'notifications', origin, { ...details, embeddingOrigin: 'https://evil.example' }), false);
  enabled = false;
  assert.equal(handlers.check(null, 'notifications', origin, details), false);
});
