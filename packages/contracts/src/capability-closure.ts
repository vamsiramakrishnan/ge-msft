import type { ActuationKind, CapabilityManifest } from './capability.js';
import type { ReadVerb } from './command-grammar.js';
import { WRITE_VERB_TO_KIND } from './command-grammar.js';

/**
 * ADR-0006 — capability closure: the single, pure definition of whether a surface's
 * *advertised* capability set matches what it can actually *do*.
 *
 * The executable capability set is the intersection
 *
 *   declared manifest ∩ CLI verbs ∩ bridge actuate() ∩ bridge read ports
 *
 * which was never computed, so the independent representations drifted (phantom actuations
 * advertised but unhandled; advertised reads with no bridge port; handled effects with no CLI
 * verb). {@link checkCapabilityClosure} computes the three disagreement sets from the three
 * authoritative inputs so a conformance test can fail the build on a *lie* (phantom / unreached
 * read) and track a *gap* (unreached handler) against an allow-list. The `context` read is a
 * runtime-served strategy probe, so it is not a bridge-port obligation.
 *
 * This is the one definition; surfaces feed it their own `handledKinds`/`readPorts`.
 */

export interface CapabilityClosureInputs {
  /** What the surface advertises (its `getCapabilities()` manifest). */
  manifest: CapabilityManifest;
  /** The `ActuationKind`s the bridge's `actuate()` actually handles. */
  handledKinds: readonly ActuationKind[];
  /** The read verbs the bridge actually serves (has a read port for). */
  readPorts: readonly ReadVerb[];
}

export interface CapabilityClosureReport {
  /**
   * **Phantoms (a lie → hard failure).** Advertised `manifest.actuations[].kind` that the bridge
   * does NOT handle. A surface must never claim a write it cannot perform.
   */
  phantoms: ActuationKind[];
  /**
   * **Unreached reads (a lie → hard failure).** Advertised `manifest.reads` with no matching bridge
   * read port. A surface must never claim a read it cannot serve.
   */
  unreachedReads: ReadVerb[];
  /**
   * **Gaps (unreached → tracked, not fatal).** Bridge-handled kinds reachable by no CLI write verb
   * (per {@link WRITE_VERB_TO_KIND}) — visible so they can be burned down deliberately.
   */
  gaps: ActuationKind[];
}

/** The set of `ActuationKind`s reachable from some CLI write verb (the closure's verb input). */
const VERB_REACHABLE_KINDS: ReadonlySet<ActuationKind> = new Set(Object.values(WRITE_VERB_TO_KIND));
const RUNTIME_SERVED_READS: ReadonlySet<ReadVerb> = new Set([
  'context',
  'list',
  'inspect',
  'properties',
  'comments',
  'attachments',
  'tables',
  'slides',
  'neighbors',
  'open',
  // DocFs coreutils — served by the runtime's command protocol (command-protocol.ts), not by
  // bridge read ports; listed here so the closure helper doesn't false-flag them as unreached.
  'ls',
  'find',
  'tail',
]);

/**
 * Compute the capability-closure report for a surface (pure). No I/O, no Office.js — just set
 * algebra over the three authoritative inputs, so it is trivially unit-testable and can run in a
 * conformance test on every build.
 */
export function checkCapabilityClosure(inputs: CapabilityClosureInputs): CapabilityClosureReport {
  const { manifest, handledKinds, readPorts } = inputs;

  const handled = new Set<ActuationKind>(handledKinds);
  const ports = new Set<ReadVerb>(readPorts);

  // Phantoms: advertised actuation kinds the bridge does not handle. De-dup advertised kinds.
  const advertisedKinds = new Set<ActuationKind>(manifest.actuations.map((a) => a.kind));
  const phantoms = [...advertisedKinds].filter((kind) => !handled.has(kind));

  // Unreached reads: advertised read verbs with no bridge read port.
  const advertisedReads = manifest.reads ?? [];
  const unreachedReads = [...new Set(advertisedReads)].filter(
    (verb) => !ports.has(verb) && !RUNTIME_SERVED_READS.has(verb),
  );

  // Gaps: handled kinds reachable by no CLI write verb. Iterating the `handled` Set de-dups
  // (symmetric with the de-duped phantom set above).
  const gaps = [...handled].filter((kind) => !VERB_REACHABLE_KINDS.has(kind));

  return { phantoms, unreachedReads, gaps };
}
