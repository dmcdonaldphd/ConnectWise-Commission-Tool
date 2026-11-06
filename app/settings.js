'use strict';

const fs = require('node:fs');
const path = require('node:path');
const { randomUUID } = require('node:crypto');
const { spawnSync } = require('node:child_process');
const paths = require('./paths');

function validateWorkbookFolder(folder) {
  try {
    if (typeof folder !== 'string' || !path.isAbsolute(folder) || !fs.statSync(folder).isDirectory()) {
      throw new Error('not an absolute directory');
    }
    fs.accessSync(folder, fs.constants.R_OK | fs.constants.W_OK);
  } catch {
    throw new Error(`Workbook folder is unavailable or not writable: ${folder}`);
  }
}

function createSettingsStore(userDataDir = paths.USER_DATA_DIR) {
  const state = path.resolve(userDataDir);
  const filename = path.join(state, 'settings.json');
  return {
    userDataDir: state,
    filename,
    load() {
      let raw;
      try { raw = fs.readFileSync(filename, 'utf8'); }
      catch (error) { if (error.code === 'ENOENT') return null; throw error; }
      try {
        const settings = JSON.parse(raw);
        if (settings.version !== 1 || typeof settings.workbookDir !== 'string' || !path.isAbsolute(settings.workbookDir)) {
          throw new Error('unsupported settings');
        }
        return { version: 1, workbookDir: settings.workbookDir };
      } catch {
        throw new Error('The saved settings could not be read. Select the workbook folder again to save new settings.');
      }
    },
    save(workbookDir) {
      validateWorkbookFolder(workbookDir);
      const settings = { version: 1, workbookDir: path.resolve(workbookDir) };
      fs.mkdirSync(state, { recursive: true, mode: 0o700 });
      const temporary = `${filename}.${randomUUID()}.tmp`;
      try {
        fs.writeFileSync(temporary, JSON.stringify(settings, null, 2) + '\n', { encoding: 'utf8', flag: 'wx', mode: 0o600 });
        fs.renameSync(temporary, filename);
      } finally {
        if (fs.existsSync(temporary)) fs.unlinkSync(temporary);
      }
      return settings;
    },
  };
}

// Native folder selection only. Workbook writing continues to use JavaScript;
// no Excel COM objects or PowerShell workbook operations are involved.
function pickWindowsFolder(initialFolder = '', run = spawnSync) {
  const script = [
    "$ErrorActionPreference = 'Stop'",
    '[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()',
    'Add-Type -AssemblyName System.Windows.Forms',
    '$dialog = New-Object System.Windows.Forms.FolderBrowserDialog',
    "$dialog.Description = 'Select the folder containing your commission workbooks'",
    '$dialog.ShowNewFolderButton = $false',
    "$dialog.SelectedPath = [Environment]::GetEnvironmentVariable('PHD_COMMISSION_INITIAL_FOLDER')",
    'try { if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { [Console]::WriteLine($dialog.SelectedPath) } } finally { $dialog.Dispose() }',
  ].join('\n');
  const result = run('powershell.exe', ['-NoLogo', '-NoProfile', '-STA', '-Command', script], {
    encoding: 'utf8', windowsHide: true,
    env: { ...process.env, PHD_COMMISSION_INITIAL_FOLDER: initialFolder },
  });
  if (result.error || result.status !== 0) throw new Error('The Windows folder picker is unavailable.');
  return String(result.stdout || '').trim() || null;
}

function activateFolder(folder, store, save = false) {
  validateWorkbookFolder(folder);
  const temporaryDir = path.join(store.userDataDir, 'Temp');
  fs.mkdirSync(temporaryDir, { recursive: true, mode: 0o700 });
  fs.accessSync(temporaryDir, fs.constants.R_OK | fs.constants.W_OK);
  if (save) store.save(folder);
  // Commit runtime paths only after storage and settings are ready. A failed
  // folder change must not redirect the next reconciliation to another folder.
  paths.configureDirectories(folder, store.userDataDir);
}

async function chooseWorkbookFolder({ store = createSettingsStore(), ask,
  initialFolder = paths.WORKBOOK_DIR, pickFolder } = {}) {
  const picker = pickFolder || (process.platform === 'win32' ? pickWindowsFolder : () => {
    throw new Error('Native folder selection is available on Windows.');
  });
  console.log('Choose the folder containing your commission workbooks.');
  let folder;
  try { folder = await picker(initialFolder); }
  catch {
    console.log('The folder picker is unavailable. You can enter the folder path here instead.');
    if (!ask) throw new Error('A folder path prompt is required when the folder picker is unavailable.');
    folder = String(await ask('Full workbook folder path (blank to cancel): ')).trim().replace(/^"(.*)"$/, '$1');
  }
  if (!folder) return false;
  try {
    activateFolder(folder, store, true);
    console.log(`Workbook folder saved: ${paths.WORKBOOK_DIR}`);
    return true;
  } catch (error) {
    console.error(`Folder selection was not completed: ${error.message}`);
    return false;
  }
}

async function initializeWorkbookFolder({ store = createSettingsStore(), ask, pickFolder } = {}) {
  let settings;
  try { settings = store.load(); }
  catch (error) { console.warn(error.message); }
  if (settings) {
    try {
      activateFolder(settings.workbookDir, store);
      return true;
    } catch (error) { console.warn(`Saved workbook folder is unavailable: ${error.message}`); }
  }
  return chooseWorkbookFolder({ store, ask, pickFolder, initialFolder: paths.PROJECT_ROOT });
}

function openFolder(folder) {
  validateWorkbookFolder(path.dirname(folder));
  if (!fs.existsSync(folder)) fs.mkdirSync(folder);
  const command = process.platform === 'win32' ? 'explorer.exe' : process.platform === 'darwin' ? 'open' : 'xdg-open';
  const result = spawnSync(command, [folder], { windowsHide: true, stdio: 'ignore' });
  // Explorer can return a nonzero status even after opening its window.
  if (result.error || (process.platform !== 'win32' && result.status !== 0))
    throw new Error(`Could not open this folder: ${folder}`);
}

module.exports = { createSettingsStore, chooseWorkbookFolder, initializeWorkbookFolder,
  pickWindowsFolder, openFolder };
