'use strict';
const { AsyncLocalStorage } = require('node:async_hooks');
const { format } = require('node:util');
const storage = new AsyncLocalStorage();
const logger = {};
for (const level of ['log', 'warn', 'error']) {
  logger[level] = (...args) => {
    const sink = storage.getStore();
    if (!sink) return console[level](...args);
    // Display/reporting failures must never interrupt a financial operation.
    try { sink({ level, message: format(...args) }); } catch { /* best effort */ }
  };
}
module.exports = { logger, withLogger: (sink, task) => storage.run(sink, task) };
