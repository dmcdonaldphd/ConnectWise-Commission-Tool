'use strict';
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');

function launchDesktopWindow(chromium, profile, url) {
  return chromium.launchPersistentContext(profile, {
    headless: false, viewport: null,
    args: [`--app=${url}`, '--window-size=1200,900', '--disable-background-mode'],
    ignoreDefaultArgs: ['--enable-automation', 'about:blank'],
  });
}

async function main() {
  const root = path.resolve(__dirname, '..');
  if (fs.existsSync(path.join(root, 'browsers'))) process.env.PLAYWRIGHT_BROWSERS_PATH = path.join(root, 'browsers');
  const { chromium } = require('playwright');
  const { createGuiService } = require('./gui-service');
  const { createGuiServer } = require('./gui-server');
  if (process.argv.includes('--self-check')) {
    for (const file of ['index.html', 'styles.css', 'ui.js']) fs.accessSync(path.join(__dirname, 'ui', file));
    fs.accessSync(chromium.executablePath());
    console.log('Desktop launcher and interface are ready.');
    return;
  }
  const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'phd-commission-ui-'));
  const service = createGuiService();
  let server, window;
  try {
    await service.initialize();
    server = await createGuiServer(service);
    // A separate, temporary Chromium profile holds the desktop UI. The existing
    // persistent ConnectWise profile is used only for authentication/API access.
    window = await launchDesktopWindow(chromium, profile, server.url);
    await new Promise(resolve => {
      window.once('close', resolve);
      const browser = window.browser();
      if (browser && !browser.isConnected()) resolve();
    });
  } finally {
    // Do not interrupt the XLSX writer when the desktop window closes.
    await service.close();
    if (window) await window.close().catch(() => {});
    if (server) await server.close();
    fs.rmSync(profile, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 });
  }
}
if (require.main === module) main().catch(error => { console.error(error.stack || error.message); process.exitCode = 1; });
module.exports = { main, launchDesktopWindow };
