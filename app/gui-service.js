'use strict';
const fs = require('node:fs');
const { EventEmitter } = require('node:events');
const { withLogger } = require('./log');

function defaultAdapters() {
  const paths = require('./paths');
  const settings = require('./settings');
  const store = settings.createSettingsStore();
  return {
    initialize: () => settings.initializeWorkbookFolder({ store, prompt: false }),
    locations: () => ({ workbookDir: paths.WORKBOOK_DIR, logDir: paths.LOG_DIR, backupDir: paths.BACKUP_DIR }),
    discover: require('./commission-menu').findCommissionWorkbooks,
    async chooseFolder(folder) {
      let needsPath = false;
      const changed = await settings.chooseWorkbookFolder({ store, throwOnError: true, pickFolder: () => {
        if (folder !== undefined) return String(folder).trim().replace(/^"(.*)"$/, '$1');
        try { return settings.pickWindowsFolder(paths.WORKBOOK_DIR); }
        catch { needsPath = true; return null; }
      } });
      return { changed, needsPath };
    },
    openDirectory: settings.openFolder,
    async openConnectWise() {
      const { chromium, BASE_URL } = require('./connectwise');
      const context = await chromium.launchPersistentContext(paths.PROFILE_DIR, { headless: false, viewport: null });
      try {
        const page = context.pages()[0] || await context.newPage();
        await page.goto(BASE_URL, { waitUntil: 'domcontentloaded' });
        return context;
      } catch (error) { await context.close().catch(() => {}); throw error; }
    },
    reconcile: require('./app').reconcileWorkbook,
    readSummary: filename => fs.readFileSync(filename, 'utf8'),
  };
}

function createGuiService(adapters) {
  const deps = adapters || defaultAdapters();
  const events = new EventEmitter();
  let activeFolder = false;
  let context = null;
  let selected = [];
  let activeRun = null;
  let operation = null;
  let closed = false;
  const state = { phase: 'starting', workbookDir: null, workbooks: [], skipped: [],
    results: [], progress: null, error: null, logs: [] };
  const snapshot = () => structuredClone(state);
  const publish = () => events.emit('state', snapshot());
  function log(entry) {
    if (!entry.message.trim()) return;
    state.logs.push({ ...entry, time: new Date().toISOString() });
    if (state.logs.length > 400) state.logs.shift();
    events.emit('log', state.logs.at(-1));
  }
  function set(patch) { Object.assign(state, patch); publish(); }
  function requireIdle() {
    if (closed || !['ready', 'complete', 'error'].includes(state.phase))
      throw new Error('An operation is already in progress. Finish it before starting another.');
  }
  async function discover() {
    if (!activeFolder) return;
    const data = await deps.discover();
    set({ workbookDir: deps.locations().workbookDir, workbooks: data.valid, skipped: data.skipped });
  }
  async function closeContext() {
    const current = context; context = null;
    if (current) await current.close().catch(error => log({ level: 'warn', message: error.message }));
  }
  async function tracked(task) {
    operation = withLogger(log, task);
    try { return await operation; }
    finally { operation = null; }
  }
  const service = {
    snapshot,
    subscribe(listener) {
      events.on('state', listener);
      return () => events.off('state', listener);
    },
    subscribeLogs(listener) { events.on('log', listener); return () => events.off('log', listener); },
    async initialize() {
      return tracked(async () => {
        try {
          activeFolder = await deps.initialize();
          await discover(); set({ phase: 'ready' });
        } catch (error) { set({ phase: 'error', error: error.message, workbooks: [] }); }
      });
    },
    async refresh() {
      requireIdle(); set({ phase: 'scanning', error: null });
      return tracked(async () => {
        try { await discover(); set({ phase: 'ready' }); }
        catch (error) { set({ phase: 'error', error: error.message, workbooks: [] }); throw error; }
      });
    },
    async changeFolder(folder) {
      requireIdle(); set({ phase: 'choosing_folder', error: null });
      return tracked(async () => {
        try {
          const outcome = await deps.chooseFolder(folder);
          if (outcome.changed) { activeFolder = true; state.results = []; await discover(); }
          set({ phase: 'ready' }); return outcome;
        } catch (error) { set({ phase: 'error', error: error.message }); throw error; }
      });
    },
    async start(names) {
      requireIdle();
      const allowed = new Set(state.workbooks.map(w => w.fileName));
      if (!activeFolder || !Array.isArray(names) || !names.length || new Set(names).size !== names.length ||
        names.some(name => typeof name !== 'string' || !allowed.has(name)))
        throw new Error('Select one or more workbooks from the current list.');
      selected = [...names];
      set({ phase: 'opening', results: [], logs: [], progress: null, error: null });
      return tracked(async () => {
        try {
          // Recheck the actual folder before opening a session; no silent fallback.
          await discover();
          if (selected.some(name => !state.workbooks.some(w => w.fileName === name)))
            throw new Error('A selected workbook is no longer available. Refresh the list.');
          context = await deps.openConnectWise();
          context.on?.('close', () => {
            if (state.phase === 'awaiting_login') {
              context = null;
              set({ phase: 'error', error: 'The ConnectWise window was closed. Open it again before continuing.' });
            }
          });
          set({ phase: 'awaiting_login' });
        } catch (error) { await closeContext(); set({ phase: 'error', error: error.message }); throw error; }
      });
    },
    async cancelLogin() {
      if (state.phase !== 'awaiting_login') throw new Error('There is no login step to cancel.');
      set({ phase: 'opening' });
      return tracked(async () => { await closeContext(); set({ phase: 'ready' }); });
    },
    async continueRun() {
      if (closed || state.phase !== 'awaiting_login' || !context)
        throw new Error('Open ConnectWise and complete the login step first.');
      set({ phase: 'running', error: null });
      activeRun = withLogger(log, async () => {
        try {
          for (let index = 0; index < selected.length; index++) {
            const filename = selected[index];
            const base = { workbookFilename: filename, workbookIndex: index + 1, workbookTotal: selected.length };
            set({ progress: { ...base, stage: 'reading' } });
            let result;
            try {
              result = await deps.reconcile(context, filename, {
                onProgress: progress => set({ progress: { ...base, ...progress } }),
              });
            } catch (error) {
              result = { ...error.runResult, workbookFilename: filename, status: 'FAILED', error: error.message,
                auditLogPath: error.auditLogPath, summaryLogPath: error.summaryLogPath };
              log({ level: 'error', message: `${filename}: ${error.message}` });
            }
            state.results.push(result); publish();
          }
        } catch (error) { set({ error: error.message }); }
        finally { await closeContext(); set({ phase: 'complete', progress: null }); }
      });
    },
    async waitForRun() { await activeRun; },
    async openDirectory(kind) {
      if (!activeFolder || !['logs', 'backups'].includes(kind)) throw new Error('Select a workbook folder first.');
      await deps.openDirectory(deps.locations()[kind === 'logs' ? 'logDir' : 'backupDir']);
    },
    async summary(index) {
      if (!Number.isInteger(index) || index < 0 || !state.results[index]?.summaryLogPath)
        throw new Error('This run has no saved summary.');
      return deps.readSummary(state.results[index].summaryLogPath);
    },
    async close() {
      closed = true;
      // Closing the UI must never terminate the writer or a selected batch.
      try { await operation; await activeRun; } finally { await closeContext(); }
    },
  };
  return service;
}
module.exports = { createGuiService };
