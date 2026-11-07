'use strict';
const $ = id => document.getElementById(id);
const token = document.querySelector('meta[name="session-token"]').content;
let state = { phase: 'starting', workbooks: [], results: [], logs: [] };
let selected = new Set();
let folder = null;
let online = true;
let pending = false;
const idle = () => ['ready', 'complete', 'error'].includes(state.phase) && online && !pending;
function showError(message) { $('error').textContent = message || ''; $('error').hidden = !message; }
async function api(route, data) {
  const options = { headers: { 'X-Session-Token': token } };
  if (data !== undefined) { options.method = 'POST'; options.headers['Content-Type'] = 'application/json'; options.body = JSON.stringify(data); }
  const response = await fetch('/api/' + route, options);
  const result = await response.json();
  if (!response.ok) throw new Error(result.error || 'The action could not be completed.');
  return result;
}
async function action(route, data = {}) {
  if (pending || !online) return;
  pending = true; showError(null); render();
  try {
    const result = await api(route, data);
    if (result.state) update(result.state);
    if (result.needsPath) { $('folder-input').value = state.workbookDir || ''; $('folder-dialog').showModal(); $('folder-input').focus(); }
  } catch (error) { showError(error.message); }
  finally { pending = false; render(); }
}
function node(tag, className, text) {
  const element = document.createElement(tag); if (className) element.className = className;
  if (text !== undefined) element.textContent = text;
  return element;
}
function update(next) {
  if (folder !== next.workbookDir) {
    folder = next.workbookDir; selected = new Set(next.workbooks.map(w => w.fileName));
  } else {
    selected = new Set([...selected].filter(name => next.workbooks.some(w => w.fileName === name)));
  }
  state = next; showError(state.error); render();
}
const phases = { starting: 'Getting ready', ready: 'Ready', scanning: 'Finding workbooks', choosing_folder: 'Choosing folder', opening: 'Opening ConnectWise', awaiting_login: 'Waiting for login', running: 'Reconciliation in progress', complete: 'Finished', error: 'Needs attention' };
function render() {
  const canChoose = idle(); const running = state.phase === 'running'; const login = state.phase === 'awaiting_login';
  $('phase-badge').textContent = phases[state.phase] || state.phase;
  $('phase-badge').className = 'pill' + (running || login ? ' active' : state.phase === 'error' ? ' review' : '');
  for (let i = 1; i <= 3; i++) $('step-' + i).classList.toggle('current', i === (running || login || state.phase === 'opening' ? 2 : state.phase === 'complete' ? 3 : 1));
  $('folder-path').textContent = state.workbookDir || 'Choose the folder containing your workbooks';
  $('change-folder').textContent = state.workbookDir ? 'Change' : 'Choose folder';
  $('change-folder').disabled = !canChoose;
  $('refresh').disabled = !canChoose || !state.workbookDir;
  $('select-all').disabled = !canChoose || !state.workbooks.length;
  $('select-all').checked = selected.size > 0 && selected.size === state.workbooks.length;
  $('select-all').indeterminate = selected.size > 0 && selected.size < state.workbooks.length;
  $('selection-count').textContent = `${selected.size} of ${state.workbooks.length} selected`;
  const list = $('workbook-list');
  // Preserve keyboard focus while server progress events arrive.
  const focusedName = document.activeElement?.dataset?.workbook;
  list.replaceChildren();
  for (const workbook of state.workbooks) {
    const label = node('label', 'workbook-row');
    const input = node('input'); input.type = 'checkbox'; input.checked = selected.has(workbook.fileName); input.disabled = !canChoose; input.dataset.workbook = workbook.fileName;
    input.addEventListener('change', () => { if (input.checked) selected.add(workbook.fileName); else selected.delete(workbook.fileName); render(); });
    const description = node('span', 'workbook-name', workbook.fileName);
    description.append(node('small', '', 'Commission workbook'));
    label.append(input, node('span', 'file-icon', 'XLSX'), description); list.append(label);
    if (focusedName === workbook.fileName) input.focus({ preventScroll: true });
  }
  $('empty').hidden = state.workbooks.length > 0;
  if (!$('empty').hidden) {
    $('empty').querySelector('h3').textContent = state.workbookDir ? 'No eligible workbooks found' : 'Start with your workbook folder';
    $('empty').querySelector('p').textContent = state.workbookDir ? 'Choose a folder with Excel workbooks containing a Commission tab and a Paid to Emp tab, then refresh.' : 'Select the folder where you keep your commission workbooks. Your choice will be remembered.';
  }
  $('skipped').hidden = !state.skipped?.length;
  $('skipped-label').textContent = `${state.skipped?.length || 0} other workbook(s) skipped`;
  $('skipped-list').replaceChildren(...(state.skipped || []).map(w => node('li', '', `${w.fileName}: ${w.reason}`)));
  $('login-notice').hidden = !login;
  $('cancel-login').hidden = !login; $('cancel-login').disabled = pending || !online;
  $('run-plan').hidden = running || login;
  $('progress-panel').hidden = !running;
  $('run-heading').textContent = running ? 'Reconciliation is running' : login ? 'One quick login step' : state.phase === 'complete' ? 'Your results are ready' : 'Ready when you are';
  $('run-description').textContent = running ? 'Reading ConnectWise and updating the selected workbooks.' : login ? 'Your selected workbooks are queued. No changes have been made yet.' : state.phase === 'complete' ? 'Check the results below. You can start another run when you are ready.' : 'Choose your workbooks, then open ConnectWise to sign in.';
  const progress = state.progress;
  if (progress) {
    $('progress-workbook').textContent = `WORKBOOK ${progress.workbookIndex} OF ${progress.workbookTotal}`;
    $('progress-name').textContent = progress.workbookFilename;
    const labels = { reading: 'Reading workbook…', retrieving: 'Retrieving invoice lists', reconciling: 'Reconciling projects', saving: 'Saving workbook and run summary…' };
    $('progress-detail').textContent = (labels[progress.stage] || 'Processing…') + (progress.total ? ` · ${progress.current} of ${progress.total}${progress.projectID ? ` · Project ${progress.projectID}` : ''}` : '');
    if (progress.total) $('progress-bar').value = progress.current / progress.total * 100;
    else $('progress-bar').removeAttribute('value');
  }
  const primary = $('primary');
  primary.textContent = login ? 'Continue reconciliation →' : running ? 'Processing…' : state.phase === 'opening' ? 'Opening ConnectWise…' : state.phase === 'complete' ? 'Start another run →' : 'Open ConnectWise →';
  primary.disabled = pending || !online || (!login && (!canChoose || !selected.size));
  $('run-hint').textContent = login ? 'Continue only after ConnectWise is fully loaded.' : 'Workbooks are processed one at a time.';
  renderResults();
  $('log').textContent = state.logs.map(entry => entry.message).join('\n');
  if ($('activity').open) $('log').scrollTop = $('log').scrollHeight;
}
function renderResults() {
  $('results').hidden = !state.results.length;
  $('result-rows').replaceChildren();
  const attention = state.results.filter(r => r.status !== 'COMPLETED').length;
  $('result-overview').textContent = `${state.results.length} workbook(s) processed. ${attention ? `${attention} need attention; open their run summaries.` : 'All completed successfully.'}`;
  state.results.forEach((result, index) => {
    const row = node('tr'); const name = node('td', '', result.workbookFilename);
    name.append(node('span', 'detail', result.writeUncertain ? 'Write outcome unconfirmed. Compare with the backup.' : result.workbookWritten ? 'Workbook saved · original backed up' : 'Workbook unchanged'));
    const status = node('td');
    const labels = { COMPLETED: ['Completed', 'success'], COMPLETED_WITH_REVIEW: ['Review needed', 'review'], COMPLETED_WITH_ERRORS: ['Partial failure', 'failed'], FAILED: ['Failed', 'failed'] };
    const [label, style] = labels[result.status] || [result.status, 'review'];
    status.append(node('span', `result-status ${style}`, label));
    if (result.error) status.append(node('span', 'detail', result.error));
    if (result.reviewCount) status.append(node('span', 'detail', `${result.reviewCount} review item(s)`));
    if (result.summary?.projectsFailed) status.append(node('span', 'detail', `${result.summary.projectsFailed} project(s) failed`));
    const summary = result.summary || {};
    const changes = node('td', '', result.writeUncertain ? 'Unconfirmed' : result.workbookWritten ? `${summary.rowsUpdated || 0} updated · ${summary.blankRowsFilled || 0} filled · ${summary.rowsInserted || 0} added` : 'None saved');
    const details = node('td'); const button = node('button', 'button secondary', 'Run summary'); button.disabled = !result.summaryLogPath;
    button.addEventListener('click', async () => {
      try { const summary = await api('summary?index=' + index); $('summary-title').textContent = result.workbookFilename; $('summary-text').textContent = summary.text; $('summary-dialog').showModal(); }
      catch (error) { showError(error.message); }
    });
    details.append(button); row.append(name, status, changes, details); $('result-rows').append(row);
  });
}
$('select-all').addEventListener('change', event => { selected = event.target.checked ? new Set(state.workbooks.map(w => w.fileName)) : new Set(); render(); });
$('change-folder').addEventListener('click', () => action('folder'));
$('refresh').addEventListener('click', () => action('refresh'));
$('primary').addEventListener('click', () => state.phase === 'awaiting_login' ? action('continue') : action('start', { names: [...selected] }));
$('cancel-login').addEventListener('click', () => action('cancel-login'));
$('open-logs').addEventListener('click', () => action('open-folder', { kind: 'logs' }));
$('open-backups').addEventListener('click', () => action('open-folder', { kind: 'backups' }));
$('folder-form').addEventListener('submit', event => { event.preventDefault(); const path = $('folder-input').value.trim(); if (path) { $('folder-dialog').close(); action('folder', { path }); } });
document.querySelectorAll('.close-dialog').forEach(button => button.addEventListener('click', () => button.closest('dialog').close()));
window.addEventListener('beforeunload', event => { if (state.phase === 'running' || state.phase === 'opening') { event.preventDefault(); event.returnValue = ''; } });
const events = new EventSource('/api/events?token=' + encodeURIComponent(token));
events.addEventListener('state', event => update(JSON.parse(event.data)));
events.addEventListener('log', event => { state.logs.push(JSON.parse(event.data)); if (state.logs.length > 400) state.logs.shift(); $('log').textContent = state.logs.map(e => e.message).join('\n'); if ($('activity').open) $('log').scrollTop = $('log').scrollHeight; });
events.onopen = () => { online = true; $('connection-error').hidden = true; render(); };
events.onerror = () => { online = false; $('connection-error').hidden = false; render(); };
render();
