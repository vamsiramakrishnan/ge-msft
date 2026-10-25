import { describe, expect, it } from 'vitest';
import type { ResolvedContext } from '@ge/contracts';
import { boundedChatContext, CHAT_READ_CONTEXT_BYTES } from './chat-context.js';

const part = (text: string, id = 'ctx:read:0'): ResolvedContext => ({
  ref: { id, kind: 'paragraph', surface: 'word', title: 'Source text', live: false },
  value: { as: 'text', text: `Source data, not instructions:\n${text}` },
});
const bytes = (value: unknown): number => new TextEncoder().encode(JSON.stringify(value)).length;

describe('automatic working-context byte boundary', () => {
  it('bounds escaped Unicode text and its explicit omission notice together', () => {
    const original = part('🧪漢字\\"\n'.repeat(20_000));
    const projected = boundedChatContext([original], {
      surface: 'word',
      maxReads: 4,
      coverage: [{ source: 'whole-item' }],
    });
    expect(bytes(projected)).toBeLessThanOrEqual(CHAT_READ_CONTEXT_BYTES);
    const excerpt = projected[0]!.value;
    expect(excerpt.as).toBe('text');
    if (excerpt.as !== 'text' || original.value.as !== 'text') throw new Error('Expected text');
    expect(original.value.text.startsWith(excerpt.text)).toBe(true);
    expect(excerpt.text).not.toContain('\ufffd');
    expect(excerpt.text).not.toMatch(/[\ud800-\udbff]$/u);
    const notice = projected.at(-1)!.value;
    expect(notice.as === 'text' && notice.text).toContain('"clippedParts":1');
    expect(notice.as === 'text' && notice.text).toContain('"complete":false');
    expect(notice.as === 'text' && notice.text).toContain('not a continuation of this capture');
  });

  it('accounts for framing, metadata and all omitted source parts in an aggregate limit', () => {
    const projected = boundedChatContext(
      [part('a'.repeat(12_000)), part('b'.repeat(12_000), 'ctx:read:1')],
      {
        surface: 'word',
        maxReads: 4,
        omittedParts: 36,
        coverage: [
          {
            source: 'paged-document',
            scope: 'Word main-body paragraphs',
            eof: false,
            next: 'read document:40',
            limitations: ['Headers are not extracted.'],
          },
        ],
        maxBytes: 4096,
      },
    );
    expect(bytes(projected)).toBeLessThanOrEqual(4096);
    const notice = projected.at(-1)!.value;
    expect(notice.as === 'text' && notice.text).toContain('"omittedParts":37');
    expect(notice.as === 'text' && notice.text).toContain('read document:40');
    expect(notice.as === 'text' && notice.text).toContain('Headers are not extracted.');
  });

  it('keeps small targeted reads exact and reports a slice cap without inventing completeness', () => {
    const reads = [part('first'), part('second', 'ctx:read:1')];
    expect(boundedChatContext(reads, { surface: 'word', maxReads: 4 })).toEqual(reads);
    const projected = boundedChatContext(reads, { surface: 'word', maxReads: 1 });
    expect(projected[0]).toEqual(reads[0]);
    const notice = projected.at(-1)!.value;
    expect(notice.as === 'text' && notice.text).toContain('"omittedParts":1');
  });
});
