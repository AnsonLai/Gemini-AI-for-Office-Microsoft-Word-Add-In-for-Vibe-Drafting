/**
 * Loads a module on first use, shares concurrent first loads, initializes it
 * once, and permits a later retry if loading or initialization fails.
 */
export function createLazyModuleLoader(loadModule, initialize = null) {
    if (typeof loadModule !== 'function') throw new TypeError('loadModule must be a function.');
    if (initialize != null && typeof initialize !== 'function') throw new TypeError('initialize must be a function when provided.');

    let pending = null;
    return function loadOnce() {
        if (!pending) {
            pending = Promise.resolve()
                .then(loadModule)
                .then(async module => {
                    if (initialize) await initialize(module);
                    return module;
                })
                .catch(error => {
                    pending = null;
                    throw error;
                });
        }
        return pending;
    };
}
