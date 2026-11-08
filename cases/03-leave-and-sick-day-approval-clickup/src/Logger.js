/**
 * Structured logger for Google Apps Script.
 *
 * Every call writes one structured entry through console.*, so it appears in
 * Apps Script "Executions" and in Google Cloud Logging with a severity and can
 * be filtered by level, function or any context field.
 *
 * Minimum level: Script Property LOG_LEVEL = DEBUG | INFO | WARN | ERROR (default INFO).
 *
 * Usage:
 *   Log.info('Export started', { entity: 'entity_uk' });
 *   Log.info('%s / %s: +%s rows', entity, job, rows);   // printf-style, like Logger.log
 *   Log.error('Payment upload failed', err, { invoiceId: id });
 */
const Log = (function () {
  const LEVELS = { DEBUG: 10, INFO: 20, WARN: 30, ERROR: 40 };
  const CONSOLE = { DEBUG: 'log', INFO: 'info', WARN: 'warn', ERROR: 'error' };
  let minLevel_ = null;

  function minLevel_get_() {
    if (minLevel_ === null) {
      let name = 'INFO';
      try {
        name = String(PropertiesService.getScriptProperties().getProperty('LOG_LEVEL') || 'INFO').toUpperCase();
      } catch (err) { /* properties unavailable (e.g. custom function context) */ }
      minLevel_ = LEVELS[name] || LEVELS.INFO;
    }
    return minLevel_;
  }

  function toText_(value) {
    if (value instanceof Error) return value.message;
    if (value !== null && typeof value === 'object') {
      try { return JSON.stringify(value); } catch (err) { return String(value); }
    }
    return String(value);
  }

  function isPlainObject_(value) {
    return value !== null && typeof value === 'object' && !(value instanceof Error) && !Array.isArray(value);
  }

  function write_(level, args) {
    if (LEVELS[level] < minLevel_get_()) return;
    const list = Array.prototype.slice.call(args);
    const errors = list.filter(function (a) { return a instanceof Error; });
    let context = null;
    if (list.length > 1 && isPlainObject_(list[list.length - 1]) &&
        !(typeof list[0] === 'string' && list[0].indexOf('%') >= 0)) {
      context = list.pop();
    }
    const rest = list.filter(function (a) { return !(a instanceof Error) || list.length === 1; });

    let message;
    if (typeof rest[0] === 'string' && rest[0].indexOf('%') >= 0 && rest.length > 1) {
      try {
        message = Utilities.formatString.apply(null, [rest[0]].concat(rest.slice(1).map(function (a) {
          return (a !== null && typeof a === 'object') ? toText_(a) : a;
        })));
      } catch (err) {
        message = rest.map(toText_).join(' ');
      }
    } else {
      message = rest.map(toText_).join(' ');
    }

    const entry = { severity: level, message: message };
    if (context) entry.context = context;
    if (errors.length) {
      entry.error = errors[0].message;
      if (errors[0].stack) entry.stack = String(errors[0].stack).split('\n').slice(0, 6).join('\n');
    }
    console[CONSOLE[level]](entry);
  }

  return {
    debug: function () { write_('DEBUG', arguments); },
    info: function () { write_('INFO', arguments); },
    warn: function () { write_('WARN', arguments); },
    error: function () { write_('ERROR', arguments); }
  };
})();
