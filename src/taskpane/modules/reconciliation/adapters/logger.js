/**
 * Logger adapter for reconciliation modules.
 */

let _logger = console;

/**
 * Configures logger implementation.
 *
 * @param {{log?: Function, warn?: Function, error?: Function}} logger - Logger object
 */
export function configureLogger(logger) {
    _logger = logger || console;
}

/**
 * Log passthrough.
 *
 * @param {...any} args - Log args
 */
export function log(...args) {
    (_logger.log || (() => { }))(...args);
}

/**
 * Warn passthrough.
 *
 * @param {...any} args - Warn args
 */
export function warn(...args) {
    (_logger.warn || (() => { }))(...args);
}

/**
 * Error passthrough.
 *
 * @param {...any} args - Error args
 */
export function error(...args) {
    (_logger.error || (() => { }))(...args);
}
