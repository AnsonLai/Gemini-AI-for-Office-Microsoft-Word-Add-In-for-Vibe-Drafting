/**
 * Logger adapter for reconciliation modules.
 */

let _logger = console;

const LEVELS = Object.freeze({
    silent: 0,
    error: 1,
    warn: 2,
    info: 3
});

const DEFAULT_LOG_LEVEL = (() => {
    const isProd = typeof process !== 'undefined' && process?.env?.NODE_ENV === 'production';
    return isProd ? 'warn' : 'info';
})();

let _logLevel = DEFAULT_LOG_LEVEL;

function normalizeLogLevel(level) {
    const normalized = String(level || '').toLowerCase();
    return Object.prototype.hasOwnProperty.call(LEVELS, normalized) ? normalized : _logLevel;
}

function isEnabled(level) {
    return LEVELS[_logLevel] >= LEVELS[level];
}

/**
 * Configures logger implementation.
 *
 * @param {{log?: Function, warn?: Function, error?: Function}} logger - Logger object
 * @param {{ level?: 'silent'|'error'|'warn'|'info' }} [options={}] - Logger options
 */
export function configureLogger(logger, options = {}) {
    _logger = logger || console;
    if (options.level) {
        _logLevel = normalizeLogLevel(options.level);
    }
}

/**
 * Sets the minimum log level for reconciliation logs.
 *
 * @param {'silent'|'error'|'warn'|'info'} level - Desired log level
 */
export function setLogLevel(level) {
    _logLevel = normalizeLogLevel(level);
}

/**
 * Gets current logger level.
 *
 * @returns {'silent'|'error'|'warn'|'info'}
 */
export function getLogLevel() {
    return _logLevel;
}

/**
 * Log passthrough.
 *
 * @param {...any} args - Log args
 */
export function log(...args) {
    if (!isEnabled('info')) return;
    (_logger.log || (() => { }))(...args);
}

/**
 * Warn passthrough.
 *
 * @param {...any} args - Warn args
 */
export function warn(...args) {
    if (!isEnabled('warn')) return;
    (_logger.warn || (() => { }))(...args);
}

/**
 * Error passthrough.
 *
 * @param {...any} args - Error args
 */
export function error(...args) {
    if (!isEnabled('error')) return;
    (_logger.error || (() => { }))(...args);
}
