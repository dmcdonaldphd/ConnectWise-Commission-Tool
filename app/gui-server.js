'use strict';
const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const { randomBytes, timingSafeEqual } = require('node:crypto');

async function createGuiServer(service) {
  const token = randomBytes(32).toString('hex');
  const clients = new Set();
  let origin;
  function validToken(value) {
    return typeof value === 'string' && /^[a-f0-9]{64}$/.test(value) &&
      timingSafeEqual(Buffer.from(value), Buffer.from(token));
  }
  function json(response, status, data) {
    response.writeHead(status, { 'Content-Type': 'application/json; charset=utf-8' });
    response.end(JSON.stringify(data));
  }
  async function body(request) {
    if (!request.headers['content-type']?.startsWith('application/json'))
      throw Object.assign(new Error('Send a JSON request.'), { status: 400 });
    const chunks = [];
    let size = 0;
    for await (const chunk of request) {
      size += chunk.length;
      if (size <= 16384) chunks.push(chunk);
    }
    if (size > 16384) throw Object.assign(new Error('Request is too large.'), { status: 413 });
    const text = Buffer.concat(chunks).toString('utf8');
    try { return JSON.parse(text || '{}'); }
    catch { throw new Error('The request could not be read.'); }
  }
  const assets = { '/styles.css': ['styles.css', 'text/css'], '/ui.js': ['ui.js', 'text/javascript'] };
  const server = http.createServer(async (request, response) => {
    response.setHeader('Cache-Control', 'no-store');
    response.setHeader('Referrer-Policy', 'no-referrer');
    response.setHeader('X-Content-Type-Options', 'nosniff');
    response.setHeader('Content-Security-Policy', "default-src 'self'; script-src 'self'; style-src 'self'; connect-src 'self'; img-src 'self' data:; object-src 'none'; base-uri 'none'; frame-ancestors 'none'; form-action 'self'");
    try {
      if (request.headers.host !== new URL(origin).host ||
          (request.headers.origin && request.headers.origin !== origin)) return json(response, 403, { error: 'This window is not authorized.' });
      const url = new URL(request.url, origin);
      const api = url.pathname.startsWith('/api/');
      if ((api || url.pathname === '/') && !validToken(request.headers['x-session-token'] || url.searchParams.get('token')))
        return json(response, 403, { error: 'Reopen the Commission Tool to start a new session.' });
      if (request.method === 'GET' && url.pathname === '/') {
        const html = fs.readFileSync(path.join(__dirname, 'ui', 'index.html'), 'utf8').replace('__SESSION_TOKEN__', token);
        response.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' }); return response.end(html);
      }
      if (request.method === 'GET' && assets[url.pathname]) {
        const [file, type] = assets[url.pathname];
        response.writeHead(200, { 'Content-Type': type + '; charset=utf-8' });
        return response.end(fs.readFileSync(path.join(__dirname, 'ui', file)));
      }
      if (request.method === 'GET' && url.pathname === '/api/state') return json(response, 200, service.snapshot());
      if (request.method === 'GET' && url.pathname === '/api/events') {
        response.writeHead(200, { 'Content-Type': 'text/event-stream', Connection: 'keep-alive' });
        response.write(`event: state\ndata: ${JSON.stringify(service.snapshot())}\n\n`);
        clients.add(response);
        const heartbeat = setInterval(() => response.write(': heartbeat\n\n'), 15000);
        request.on('close', () => { clearInterval(heartbeat); clients.delete(response); });
        return;
      }
      if (request.method === 'GET' && url.pathname === '/api/summary') {
        const raw = url.searchParams.get('index');
        if (!/^\d+$/.test(raw || '')) throw new Error('Select a completed run.');
        return json(response, 200, { text: await service.summary(Number(raw)) });
      }
      if (request.method === 'POST' && api) {
        const data = await body(request);
        if (!data || typeof data !== 'object' || Array.isArray(data)) throw new Error('Send a JSON object.');
        let result;
        switch (url.pathname) {
          case '/api/refresh': result = await service.refresh(); break;
          case '/api/folder':
            if (data.path !== undefined && typeof data.path !== 'string') throw new Error('Enter a folder path.');
            result = await service.changeFolder(data.path); break;
          case '/api/start': result = await service.start(data.names); break;
          case '/api/continue': result = await service.continueRun(); break;
          case '/api/cancel-login': result = await service.cancelLogin(); break;
          case '/api/open-folder': result = await service.openDirectory(data.kind); break;
          default: return json(response, 404, { error: 'Unknown action.' });
        }
        return json(response, 200, { ...result, state: service.snapshot() });
      }
      json(response, 404, { error: 'Page not found.' });
    } catch (error) {
      if (!response.headersSent && !response.destroyed) json(response, error.status || 400, { error: error.message });
    }
  });
  const broadcast = (event, data) => {
    for (const client of clients) client.write(`event: ${event}\ndata: ${JSON.stringify(data)}\n\n`);
  };
  const unsubscribe = service.subscribe(state => broadcast('state', state));
  const unsubscribeLogs = service.subscribeLogs(entry => broadcast('log', entry));
  await new Promise((resolve, reject) => { server.once('error', reject); server.listen(0, '127.0.0.1', resolve); });
  origin = `http://127.0.0.1:${server.address().port}`;
  return { origin, url: `${origin}/?token=${token}`,
    async close() {
      unsubscribe(); unsubscribeLogs();
      for (const client of clients) client.end();
      await new Promise(resolve => { server.close(resolve); server.closeIdleConnections(); });
    },
  };
}
module.exports = { createGuiServer };
