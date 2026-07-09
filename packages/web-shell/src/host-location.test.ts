import { describe, expect, it } from 'vitest';
import { canRenderHostLocation, synthesizeLocationRef } from './host-location.js';

describe('host-location', () => {
  it('synthesizes Excel range refs', () => {
    expect(synthesizeLocationRef('excel', "'Daily schedule'!K6:L18")).toMatchObject({
      id: "xl:'Daily schedule'!K6:L18",
      kind: 'range',
      surface: 'excel',
      hostRef: { type: 'excel.range', address: "'Daily schedule'!K6:L18" },
    });
    expect(synthesizeLocationRef('excel', "<citation:'Daily schedule'!K6:L18>")).toMatchObject({
      id: "xl:'Daily schedule'!K6:L18",
      kind: 'range',
      hostRef: { type: 'excel.range', address: "'Daily schedule'!K6:L18" },
    });
  });

  it('synthesizes Word selection, comment, content-control, and text-anchor refs', () => {
    expect(synthesizeLocationRef('word', 'word:selection')).toMatchObject({
      id: 'word:selection',
      kind: 'selection',
    });
    expect(synthesizeLocationRef('word', 'comment:c1')).toMatchObject({
      id: 'word:comment:c1',
      kind: 'comment',
      hostRef: { type: 'word.comment', commentId: 'c1' },
    });
    expect(synthesizeLocationRef('word', 'content-control:42')).toMatchObject({
      id: 'word:cc:42',
      hostRef: { type: 'word.contentControl', contentControlId: '42' },
      anchor: { locator: 'content-control:42' },
    });
    expect(synthesizeLocationRef('word', 'citation:paragraph:7')).toMatchObject({
      id: 'word:paragraph:7',
      kind: 'paragraph',
      anchor: { locator: 'word:paragraph:7' },
      hostRef: { type: 'word.range' },
    });
    expect(synthesizeLocationRef('word', 'citation:heading:Service availability')).toMatchObject({
      kind: 'selection',
      anchor: { matchText: 'Service availability' },
    });
    expect(synthesizeLocationRef('word', 'text:Service availability')).toMatchObject({
      kind: 'selection',
      anchor: { matchText: 'Service availability' },
    });
  });

  it('synthesizes PowerPoint slide and shape refs', () => {
    expect(synthesizeLocationRef('powerpoint', 'slide:s2')).toMatchObject({
      id: 'pp:slide:s2',
      kind: 'slide',
      hostRef: { type: 'powerpoint.slide', slideId: 's2' },
      anchor: { locator: 'slide:s2' },
    });
    expect(synthesizeLocationRef('powerpoint', 'shape:s2:sh7')).toMatchObject({
      id: 'pp:shape:s2:sh7',
      kind: 'shape',
      hostRef: { type: 'powerpoint.shape', slideId: 's2', shapeId: 'sh7' },
      anchor: { locator: 'shape:s2:sh7' },
    });
  });

  it('synthesizes OneNote, Outlook, and Teams refs', () => {
    expect(synthesizeLocationRef('onenote', 'page:p1')).toMatchObject({
      id: 'on:page:p1',
      kind: 'page',
      hostRef: { type: 'onenote.page', pageId: 'p1' },
    });
    expect(synthesizeLocationRef('outlook', 'outlook:item:msg-1')).toMatchObject({
      id: 'msg-1',
      kind: 'mail-item',
      hostRef: { type: 'outlook.item', itemId: 'msg-1' },
    });
    expect(synthesizeLocationRef('teams', 'teams:transcript')).toMatchObject({
      id: 'teams:transcript',
      kind: 'transcript',
    });
    expect(
      synthesizeLocationRef('teams', 'teams:link:https://teams.microsoft.com/l/message/19:abc/123'),
    ).toMatchObject({
      kind: 'transcript',
      hostRef: { type: 'teams.deepLink' },
    });
  });

  it('does not render a target under the wrong active surface', () => {
    expect(canRenderHostLocation('excel', 'K6:L18')).toBe(true);
    expect(canRenderHostLocation('word', 'K6:L18')).toBe(false);
    expect(canRenderHostLocation('teams', 'javascript:alert(1)')).toBe(false);
  });
});
