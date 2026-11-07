'use strict';
const { test } = require('node:test');
const assert = require('node:assert/strict');
const http = require('node:http');
const { createGuiServer } = require('../app/gui-server');
const { createGuiService } = require('../app/gui-service');

async function fixture(t) {
  const service = createGuiService({ initialize: async () => false });
  await service.initialize();
  const server = await createGuiServer(service);
  t.after(() => server.close());
  return { server, base: server.origin, token: new URL(server.url).searchParams.get('token') };
}
test('local GUI requires its session token and rejects foreign origins and hosts', async t => {
  const f = await fixture(t);
  assert.equal((await fetch(f.base + '/api/state')).status, 403);
  assert.equal((await fetch(f.base + '/?token=wrong')).status, 403);
  assert.equal((await fetch(f.base + '/api/state', { headers: { 'X-Session-Token': f.token, Origin: 'https://example.org' } })).status, 403);
  const hostStatus = await new Promise((resolve, reject) => {
    http.get(f.base + '/api/state', { headers: { 'X-Session-Token': f.token, Host: 'example.org' } }, response => {
      response.resume(); resolve(response.statusCode);
    }).on('error', reject);
  });
  assert.equal(hostStatus, 403);
  const response = await fetch(f.base + '/api/state', { headers: { 'X-Session-Token': f.token } });
  assert.equal(response.status, 200); assert.equal((await response.json()).phase, 'ready');
  assert.equal(response.headers.get('access-control-allow-origin'), null);
});
test('GUI serves only explicit assets and never arbitrary local files', async t => {
  const f = await fixture(t);
  const response = await fetch(f.server.url);
  assert.equal(response.status, 200);
  assert((await response.text()).includes(f.token));
  assert(response.headers.get('content-security-policy').includes("default-src 'self'"));
  for (const pathname of ['/package.json', '/app/settings.js', '/.env', '/styles.css/../package.json'])
    assert.equal((await fetch(f.base + pathname)).status, 404);
});
test('write routes require JSON, bound request size, and return actionable operation errors', async t => {
  const f = await fixture(t);
  const headers = { 'X-Session-Token': f.token, 'Content-Type': 'application/json' };
  let response = await fetch(f.base + '/api/start', { method: 'POST', headers, body: '{' });
  assert.equal(response.status, 400);
  response = await fetch(f.base + '/api/start', { method: 'POST', headers, body: JSON.stringify({ names: ['Unknown.xlsx'] }) });
  assert.equal(response.status, 400); assert.match((await response.json()).error, /select/i);
  response = await fetch(f.base + '/api/start', { method: 'POST', headers, body: JSON.stringify({ names: ['x'.repeat(20000)] }) });
  assert.equal(response.status, 413);
});
