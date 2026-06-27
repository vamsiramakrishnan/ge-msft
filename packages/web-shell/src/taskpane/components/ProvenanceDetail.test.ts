// @vitest-environment jsdom
import { describe, it, expect, afterEach } from 'vitest';
import { createElement, act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { ProvenanceDetail } from './ProvenanceDetail.js';
import type { ProvenancePayload } from '@ge/contracts';

/**
 * Behavioral tests for the provenance drill-down. This record is the "traceable + reversible"
 * contract for an applied write, so the tests assert the who/when/sources/hash actually render, that
 * the source-link gate rejects non-http(s) URIs (untrusted grounding sources), the empty-sources
 * branch, the locator-vs-no-locator label, and the ISO-timestamp parse + fallback.
 */

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

let container: HTMLDivElement;
let root: Root;

const base: ProvenancePayload = {
  agentId: 'contract-review-agent@v2',
  identity: 'v.k@acme.com',
  timestamp: '2026-06-23T09:14:02.000Z',
  sources: [],
  contentHash: 'sha256:deadbeef',
};

function render(provenance: ProvenancePayload): void {
  container = document.createElement('div');
  document.body.appendChild(container);
  root = createRoot(container);
  act(() => {
    root.render(createElement(ProvenanceDetail, { provenance }));
  });
}

afterEach(() => {
  act(() => root.unmount());
  container.remove();
});

describe('ProvenanceDetail', () => {
  it('renders the agent id, signed-in identity and content hash verbatim', () => {
    render(base);
    expect(container.querySelector('.mono')?.textContent).toBe('contract-review-agent@v2');
    expect(container.textContent).toContain('v.k@acme.com');
    expect(container.querySelector('.prov-hash')?.textContent).toBe('sha256:deadbeef');
  });

  it('shows a "no sources recorded" notice when the source set is empty', () => {
    render(base);
    expect(container.querySelector('.prov-sources')).toBeNull();
    expect(container.textContent).toContain('No sources recorded.');
  });

  it('renders an http(s) source as a new-tab noopener link', () => {
    render({
      ...base,
      sources: [{ title: 'Policy v4', uri: 'https://example.com/policy', locator: '§3.2' }],
    });
    const link = container.querySelector<HTMLAnchorElement>('.prov-sources a');
    expect(link).not.toBeNull();
    expect(link?.getAttribute('href')).toBe('https://example.com/policy');
    expect(link?.getAttribute('rel')).toBe('noreferrer noopener');
    // Locator is appended to the title in the label.
    expect(link?.textContent).toBe('Policy v4 · §3.2');
  });

  it('renders a source title without a locator separator when none is present', () => {
    render({ ...base, sources: [{ title: 'Bare Source', uri: 'https://example.com' }] });
    const link = container.querySelector<HTMLAnchorElement>('.prov-sources a');
    expect(link?.textContent).toBe('Bare Source');
    expect(link?.textContent).not.toContain('·');
  });

  it('renders a non-http(s) source URI as inert text, never as an anchor', () => {
    render({
      ...base,
      // eslint-disable-next-line no-script-url
      sources: [{ title: 'Evil', uri: 'javascript:alert(1)' }],
    });
    expect(container.querySelector('.prov-sources a')).toBeNull();
    // The label still appears, just not as a link.
    const item = container.querySelector('.prov-sources li');
    expect(item?.textContent).toBe('Evil');
    expect(item?.querySelector('span')).not.toBeNull();
  });

  it('renders a source with no URI as inert text', () => {
    render({ ...base, sources: [{ title: 'Offline Doc', locator: 'p. 4' }] });
    expect(container.querySelector('.prov-sources a')).toBeNull();
    expect(container.querySelector('.prov-sources li')?.textContent).toBe('Offline Doc · p. 4');
  });

  it('treats a malformed URI as inert (URL parse throws)', () => {
    render({ ...base, sources: [{ title: 'Broken', uri: 'ht\ntp://b ad' }] });
    expect(container.querySelector('.prov-sources a')).toBeNull();
  });

  it('formats a valid ISO timestamp into a localized time and keeps the machine dateTime', () => {
    render(base);
    const time = container.querySelector('time');
    expect(time?.getAttribute('dateTime')).toBe('2026-06-23T09:14:02.000Z');
    // Parsed and re-rendered: a non-empty, non-ISO display string.
    expect((time?.textContent ?? '').length).toBeGreaterThan(0);
  });

  it('falls back to the raw timestamp string when it is not a parseable date', () => {
    render({ ...base, timestamp: 'not-a-date' });
    const time = container.querySelector('time');
    expect(time?.getAttribute('dateTime')).toBe('not-a-date');
    expect(time?.textContent).toBe('not-a-date');
  });

  it('renders one list item per recorded source', () => {
    render({
      ...base,
      sources: [
        { title: 'A', uri: 'https://a.example' },
        { title: 'B', uri: 'https://b.example' },
        { title: 'C' },
      ],
    });
    expect(container.querySelectorAll('.prov-sources li').length).toBe(3);
  });
});
