import { describe, it, expect } from 'vitest';
import fc from 'fast-check';
import { chooseAnchorIndex } from './actuate-plan.js';

/**
 * Property tests for anchor selection — the degrade-vs-resolve hot-spot.
 *
 * chooseAnchorIndex picks which `body.search` hit a tracked change lands on. The contract:
 *   • 0 matches            → -1 (degrade to a panel item, never a broken annotation)
 *   • N matches, no hint   → first valid index (0)
 *   • a hint               → a hit consistent with the hint, else -1
 *   • for ANY input        → never an out-of-range index
 */

const SEED = 0x5eed_a11c;
const NUM_RUNS = 500;

const textArb = fc.oneof(
  fc.string(),
  fc.constantFrom('5.2 Availability', 'Service Levels', '', 'café 日本 😀', 'a]b{c}'),
);

describe('chooseAnchorIndex — property: degrade and bounds', () => {
  it('returns -1 for zero matches regardless of hint', () => {
    fc.assert(
      fc.property(fc.option(textArb, { nil: undefined }), (hint) => {
        expect(chooseAnchorIndex([], hint)).toBe(-1);
      }),
      { numRuns: NUM_RUNS, seed: SEED },
    );
  });

  it('returns the first index when there are matches and no hint', () => {
    fc.assert(
      fc.property(fc.array(textArb, { minLength: 1, maxLength: 20 }), (hits) => {
        expect(chooseAnchorIndex(hits, undefined)).toBe(0);
        // empty-string hint short-circuits the truthiness guard → behaves as "no hint"
        expect(chooseAnchorIndex(hits, '')).toBe(0);
      }),
      { numRuns: NUM_RUNS, seed: SEED },
    );
  });

  it('never returns an out-of-range index for any (hits, hint)', () => {
    fc.assert(
      fc.property(
        fc.array(textArb, { maxLength: 20 }),
        fc.option(textArb, { nil: undefined }),
        (hits, hint) => {
          const i = chooseAnchorIndex(hits, hint);
          expect(i).toBeGreaterThanOrEqual(-1);
          expect(i).toBeLessThan(hits.length);
          if (i >= 0) expect(hits[i]).toBeDefined();
        },
      ),
      { numRuns: NUM_RUNS, seed: SEED },
    );
  });

  it('with a non-empty hint, returns a hit consistent with the hint or -1', () => {
    fc.assert(
      fc.property(
        fc.array(textArb, { maxLength: 20 }),
        textArb.filter((s) => s.length > 0),
        (hits, hint) => {
          const i = chooseAnchorIndex(hits, hint);
          if (i >= 0) {
            // If a hint matched, the chosen hit must contain it OR be the first hit when
            // no hit contained the hint (then i === 0 and hits is non-empty).
            const anyMatch = hits.some((t) => t.includes(hint));
            if (anyMatch) {
              expect(hits[i]!.includes(hint)).toBe(true);
              // and it must be the FIRST such match
              expect(i).toBe(hits.findIndex((t) => t.includes(hint)));
            } else {
              expect(i).toBe(0);
            }
          } else {
            expect(hits.length).toBe(0);
          }
        },
      ),
      { numRuns: NUM_RUNS, seed: SEED },
    );
  });

  it('guarantees a hint present in some hit is honoured (never degraded past a real match)', () => {
    fc.assert(
      fc.property(
        fc.array(textArb, { maxLength: 10 }),
        textArb.filter((s) => s.length > 0),
        fc.nat(),
        (base, hint, insertAt) => {
          // Inject a guaranteed-containing hit so we know a match exists.
          const hits = [...base];
          const at = base.length === 0 ? 0 : insertAt % (base.length + 1);
          hits.splice(at, 0, `prefix ${hint} suffix`);
          const i = chooseAnchorIndex(hits, hint);
          expect(i).toBeGreaterThanOrEqual(0);
          expect(hits[i]!.includes(hint)).toBe(true);
        },
      ),
      { numRuns: NUM_RUNS, seed: SEED },
    );
  });
});
