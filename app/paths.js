const path = require('path');

const PROJECT_ROOT = path.resolve(__dirname, '..');

module.exports = {
  PROJECT_ROOT,

  WORKBOOK_DIR: PROJECT_ROOT,

  BACKUP_DIR: path.join(PROJECT_ROOT, 'Backups'),

  LOG_DIR: path.join(PROJECT_ROOT, 'Logs'),

  PROFILE_DIR: path.join(PROJECT_ROOT, 'connectwise-profile'),

  TEMP_DIR: PROJECT_ROOT,

  WRITER_SCRIPT: path.join(__dirname, 'apply-workbook-changes.ps1'),
};