import { describe, it, expect } from 'vitest';
import { asChangeId } from '@ge/contracts';
import { planAppendPage } from './actuate-plan.js';

describe('onenote append-page planning (pure)', () => {
  it('uses prebuilt html verbatim and a caller-chosen title', () => {
    const plan = planAppendPage({
      changeId: asChangeId('c1'),
      kind: 'append-page',
      surface: 'onenote',
      params: {
        html: '<p>Already tagged <span data-ge-cite="1">[x]</span></p>',
        target: { matchText: 'Risk synthesis' },
      },
    });
    expect(plan.title).toBe('Risk synthesis');
    expect(plan.html).toBe('<p>Already tagged <span data-ge-cite="1">[x]</span></p>');
  });

  it('renders plain text + the first source into a citation-tagged paragraph', () => {
    const plan = planAppendPage({
      changeId: asChangeId('c2'),
      kind: 'append-page',
      surface: 'onenote',
      params: {
        text: 'The SLA sits below standard.',
        sources: [{ title: 'Risk Policy', locator: '§3.2' }],
      },
    });
    expect(plan.title).toBe('Synthesis');
    expect(plan.html).toBe(
      '<p>The SLA sits below standard. <span data-ge-cite="1">[Risk Policy · §3.2]</span></p>',
    );
  });

  it('appends extra citation tags when several sources ground one claim', () => {
    const plan = planAppendPage({
      changeId: asChangeId('c3'),
      kind: 'append-page',
      surface: 'onenote',
      params: {
        text: 'Claim',
        sources: [{ title: 'A' }, { title: 'B', locator: '§5' }],
      },
    });
    expect(plan.html).toContain('[A]');
    expect(plan.html).toContain('[B · §5]');
  });

  it('escapes injected markup in plain-text synthesis', () => {
    const plan = planAppendPage({
      changeId: asChangeId('c4'),
      kind: 'append-page',
      surface: 'onenote',
      params: { text: '<img src=x onerror=1>' },
    });
    expect(plan.html).toBe('<p>&lt;img src=x onerror=1&gt;</p>');
  });

  it('yields empty html for an empty synthesis', () => {
    const plan = planAppendPage({
      changeId: asChangeId('c5'),
      kind: 'append-page',
      surface: 'onenote',
      params: {},
    });
    expect(plan.html).toBe('');
  });
});
