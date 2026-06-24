import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { QUICK_ACTIONS } from './quick-actions.js';

/**
 * Doc-drift guard (Workstream J). The prose in README.md / docs/STATUS.md cites the size of the
 * prebuilt-button catalog as "N actions". That number is computable — `QUICK_ACTIONS.length` — so it
 * should never be hand-maintained out of sync. This test fails the build the moment the docs cite a
 * count that no longer matches the catalog, so adding/removing a quick action forces the prose update.
 *
 * Tests run with cwd = repo root (single root vitest.config.ts), so the doc paths are repo-relative.
 */
const DOCS = ['README.md', 'docs/STATUS.md'];

describe('doc metrics stay in sync with the catalog', () => {
  for (const path of DOCS) {
    it(`${path} cites the real QUICK_ACTIONS count`, () => {
      const text = readFileSync(path, 'utf8');
      // Match every "(NN actions" / "NN actions)" style claim referring to the catalog size.
      const claims = [...text.matchAll(/(\d+)\s+actions\b/g)].map((m) => Number(m[1]));
      expect(claims.length).toBeGreaterThan(0); // the doc must actually state a count
      for (const claimed of claims) {
        expect(claimed).toBe(QUICK_ACTIONS.length);
      }
    });
  }
});
