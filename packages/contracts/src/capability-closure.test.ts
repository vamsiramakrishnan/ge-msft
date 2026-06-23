import { describe, it, expect } from 'vitest';
import type { ActuationKind, CapabilityManifest } from './capability.js';
import type { ReadVerb } from './command-grammar.js';
import { checkCapabilityClosure } from './capability-closure.js';

/** A small manifest builder so each test states only the drift it exercises. */
function manifest(kinds: ActuationKind[], reads?: ReadVerb[]): CapabilityManifest {
  return {
    surface: 'word',
    contextKinds: ['document'],
    actuations: kinds.map((kind) => ({
      kind,
      surface: 'word' as const,
      title: kind,
      reversible: true,
    })),
    ...(reads ? { reads } : {}),
  };
}

describe('checkCapabilityClosure (ADR-0006)', () => {
  it('a clean manifest reports no phantoms, unreached reads, or gaps', () => {
    // tracked-change is reachable by `suggest`; add-comment by `comment`; both handled; both reads served.
    const report = checkCapabilityClosure({
      manifest: manifest(['tracked-change', 'add-comment'], ['outline', 'read']),
      handledKinds: ['tracked-change', 'add-comment'],
      readPorts: ['outline', 'read'],
    });
    expect(report).toEqual({ phantoms: [], unreachedReads: [], gaps: [] });
  });

  it('flags a phantom: an advertised kind the bridge does not handle', () => {
    const report = checkCapabilityClosure({
      manifest: manifest(['tracked-change', 'insert-ooxml']),
      handledKinds: ['tracked-change'], // insert-ooxml advertised but unhandled
      readPorts: [],
    });
    expect(report.phantoms).toEqual(['insert-ooxml']);
    expect(report.unreachedReads).toEqual([]);
  });

  it('flags an unreached read: an advertised read with no bridge port', () => {
    const report = checkCapabilityClosure({
      manifest: manifest(['tracked-change'], ['outline', 'read', 'search']),
      handledKinds: ['tracked-change'],
      readPorts: ['outline', 'read'], // search advertised but no port
    });
    expect(report.unreachedReads).toEqual(['search']);
    expect(report.phantoms).toEqual([]);
  });

  it('flags a gap: a handled kind reachable by no CLI write verb', () => {
    // set-speaker-notes is a handled effect but no write verb maps to it → a tracked gap, not a
    // phantom. (post-message/insert-slide/append-page/reply-mail are now reachable via the
    // post/slide/page/mail verbs — ADR-0006 CLI parity — so they are no longer gaps.)
    const report = checkCapabilityClosure({
      manifest: manifest(['set-speaker-notes']),
      handledKinds: ['set-speaker-notes'],
      readPorts: [],
    });
    expect(report.gaps).toEqual(['set-speaker-notes']);
    expect(report.phantoms).toEqual([]); // it IS advertised AND handled — not a phantom
  });

  it('comment-reply is reachable via the reply verb (not a gap)', () => {
    const report = checkCapabilityClosure({
      manifest: manifest(['comment-reply']),
      handledKinds: ['comment-reply'],
      readPorts: [],
    });
    expect(report.gaps).toEqual([]);
    expect(report.phantoms).toEqual([]);
  });

  it('treats an absent manifest.reads as "no reads declared" (no unreached reads)', () => {
    const report = checkCapabilityClosure({
      manifest: manifest(['tracked-change']),
      handledKinds: ['tracked-change'],
      readPorts: ['outline', 'read', 'search'], // ports exist but nothing is advertised
    });
    expect(report.unreachedReads).toEqual([]);
  });

  it('de-duplicates a kind advertised twice when reporting a phantom', () => {
    const report = checkCapabilityClosure({
      manifest: manifest(['insert-ooxml', 'insert-ooxml']),
      handledKinds: [],
      readPorts: [],
    });
    expect(report.phantoms).toEqual(['insert-ooxml']);
  });
});
