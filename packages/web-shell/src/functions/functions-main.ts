import { registerGeAsk } from './ge-ask.js';
import { createDefaultGeAskAssist } from './boot.js';
import type { RawEnv } from '../taskpane/config.js';

/**
 * Functions-runtime entry: build the client-direct assist chain once, then associate
 * `=GE.ASK` with Excel's custom-functions registry. A failed boot must not throw at module
 * scope (Excel would mark every function #BUSY forever) — the association simply never
 * happens and cells show a name error, which is the truthful signal that the backend is down.
 */
void (async () => {
  try {
    const env = import.meta.env as unknown as RawEnv;
    const assist = await createDefaultGeAskAssist(env);
    registerGeAsk({ assist });
  } catch {
    // No association → Excel reports an unregistered function name; retry happens on reload.
  }
})();
