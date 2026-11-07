const path = require('path');
const os = require('os');

const PROJECT_ROOT = path.resolve(__dirname, '..');

function getUserDataDirectory(platform = process.platform, env = process.env) {
  const platformPath = platform === 'win32' ? path.win32 : path;
  if (env.PHD_COMMISSION_DATA_DIR) return platformPath.resolve(env.PHD_COMMISSION_DATA_DIR);
  const base = platform === 'win32'
    ? env.LOCALAPPDATA || platformPath.join(os.homedir(), 'AppData', 'Local')
    : env.XDG_DATA_HOME || path.join(os.homedir(), '.local', 'share');
  return platformPath.join(base, 'PHDTech', 'Commission Tool');
}

function configureDirectories(workbookDir, userDataDir = module.exports.USER_DATA_DIR) {
  const selected = path.resolve(workbookDir);
  const state = path.resolve(userDataDir);
  Object.assign(module.exports, {
    WORKBOOK_DIR: selected,
    BACKUP_DIR: path.join(selected, 'Backups'),
    LOG_DIR: path.join(selected, 'Logs'),
    USER_DATA_DIR: state,
    PROFILE_DIR: path.join(state, 'connectwise-profile'),
    TEMP_DIR: path.join(state, 'Temp'),
  });
}

module.exports = {
  PROJECT_ROOT,
  USER_DATA_DIR: getUserDataDirectory(),
  getUserDataDirectory,
  configureDirectories,

  WORKBOOK_DIR: PROJECT_ROOT,

  BACKUP_DIR: path.join(PROJECT_ROOT, 'Backups'),

  LOG_DIR: path.join(PROJECT_ROOT, 'Logs'),

  PROFILE_DIR: path.join(PROJECT_ROOT, 'connectwise-profile'),

  TEMP_DIR: PROJECT_ROOT,

  WRITER_SCRIPT: path.join(__dirname, 'apply-workbook-changes.ps1'),
};
