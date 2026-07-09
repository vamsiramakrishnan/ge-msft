import type { ContextRef, Surface } from '@ge/contracts';

/**
 * Shared navigation-only host target helpers.
 *
 * UI affordances can surface locations from many places: context chips, plan previews, citations,
 * assistant prose, and future right-click/context-menu surfaces. They should all converge here so
 * Office.js-specific reveal semantics stay behind `ContextLister.revealContext`.
 */

export function canRenderHostLocation(surface: Surface, location: string): boolean {
  return synthesizeLocationRef(surface, location) !== undefined;
}

export function synthesizeLocationRef(surface: Surface, location: string): ContextRef | undefined {
  const target = normalizeHostLocationInput(surface, location);
  if (!target) return undefined;

  switch (surface) {
    case 'excel':
      return excelLocationRef(target);
    case 'word':
      return wordLocationRef(target);
    case 'powerpoint':
      return powerpointLocationRef(target);
    case 'onenote':
      return onenoteLocationRef(target);
    case 'outlook':
      return outlookLocationRef(target);
    case 'teams':
      return teamsLocationRef(target);
  }
}

export function findContextRefForLocation(
  refs: Iterable<ContextRef>,
  surface: Surface,
  location: string,
): ContextRef | undefined {
  const needle = normalizeLocation(location);
  for (const ref of refs) {
    if (ref.surface !== surface) continue;
    const candidates = [
      ref.id,
      ref.title,
      ref.anchor?.locator,
      ref.hostRef?.type === 'excel.range' ? ref.hostRef.address : undefined,
    ];
    if (candidates.some((candidate) => normalizeLocation(candidate ?? '') === needle)) return ref;
  }
  return undefined;
}

export function normalizeLocation(value: string): string {
  return value
    .trim()
    .replace(/^<(.+)>$/, '$1')
    .replace(/^citation:/i, '')
    .replace(/^range:/i, '')
    .replace(/^xl:/i, '')
    .replace(/\$/g, '')
    .toLowerCase();
}

function normalizeHostLocationInput(surface: Surface, value: string): string {
  let target = value.trim();
  if (target.startsWith('<') && target.endsWith('>')) target = target.slice(1, -1).trim();
  try {
    target = decodeURIComponent(target);
  } catch {
    // Keep the raw host location if it is not URI-encoded.
  }
  if (target.toLowerCase().startsWith('citation:')) target = target.slice('citation:'.length);

  const surfacePrefixes: Record<Surface, string[]> = {
    excel: ['excel:', 'xl:'],
    word: ['word:'],
    powerpoint: ['powerpoint:', 'ppt:', 'pp:'],
    onenote: ['onenote:', 'on:'],
    outlook: ['outlook:', 'mail:'],
    teams: [],
  };
  for (const prefix of surfacePrefixes[surface]) {
    if (target.toLowerCase().startsWith(prefix)) return target.slice(prefix.length).trim();
  }
  return target;
}

function looksLikeExcelTarget(value: string): boolean {
  const quotedSheet = "'[^']+'";
  const bareSheet = '[A-Za-z_][A-Za-z0-9_ .-]*';
  const sheet = `(?:${quotedSheet}|${bareSheet})!`;
  const cell = '\\$?[A-Za-z]{1,3}\\$?\\d{1,7}';
  const range = `${cell}(?::${cell})?`;
  return new RegExp(`^(?:${sheet})?${range}$`).test(value.trim());
}

function excelLocationRef(target: string): ContextRef | undefined {
  if (!looksLikeExcelTarget(target)) return undefined;
  return {
    id: `xl:${target}`,
    kind: 'range',
    surface: 'excel',
    title: target,
    hostRef: { type: 'excel.range', address: target },
  };
}

function wordLocationRef(target: string): ContextRef | undefined {
  if (target === 'word:selection' || target === 'selection') {
    return { id: 'word:selection', kind: 'selection', surface: 'word', title: 'Selection' };
  }
  const paragraphId = prefixedValue(target, 'word:paragraph:', 'paragraph:', 'para:');
  if (paragraphId) {
    return {
      id: `word:paragraph:${paragraphId}`,
      kind: 'paragraph',
      surface: 'word',
      title: `Paragraph ${paragraphId}`,
      anchor: { matchText: '', locator: `word:paragraph:${paragraphId}` },
      hostRef: {
        type: 'word.range',
        anchor: { matchText: '', locator: `word:paragraph:${paragraphId}` },
      },
    };
  }
  const commentId = prefixedValue(target, 'word:comment:', 'comment:');
  if (commentId) {
    return {
      id: `word:comment:${commentId}`,
      kind: 'comment',
      surface: 'word',
      title: `Comment ${commentId}`,
      hostRef: { type: 'word.comment', commentId },
    };
  }
  const contentControlId = prefixedValue(
    target,
    'word:cc:',
    'word:content-control:',
    'content-control:',
    'cc:',
  );
  if (contentControlId) {
    return {
      id: `word:cc:${contentControlId}`,
      kind: 'selection',
      surface: 'word',
      title: `Content control ${contentControlId}`,
      hostRef: { type: 'word.contentControl', contentControlId },
      anchor: { matchText: '', locator: `content-control:${contentControlId}` },
    };
  }
  const matchText = prefixedValue(target, 'word:text:', 'text:', 'heading:', 'word:heading:');
  if (matchText) {
    return {
      id: `word:text:${hashId(matchText)}`,
      kind: 'selection',
      surface: 'word',
      title: truncate(matchText),
      anchor: { matchText },
    };
  }
  return undefined;
}

function powerpointLocationRef(target: string): ContextRef | undefined {
  const slideId = prefixedValue(target, 'pp:slide:', 'powerpoint:slide:', 'slide:');
  if (slideId) {
    return {
      id: `pp:slide:${slideId}`,
      kind: 'slide',
      surface: 'powerpoint',
      title: `Slide ${slideId}`,
      hostRef: { type: 'powerpoint.slide', slideId },
      anchor: { matchText: '', locator: `slide:${slideId}` },
    };
  }
  const shape = prefixedValue(target, 'pp:shape:', 'powerpoint:shape:', 'shape:');
  if (shape) {
    const [shapeSlideId, shapeId] = shape.split(':');
    if (!shapeSlideId || !shapeId) return undefined;
    return {
      id: `pp:shape:${shapeSlideId}:${shapeId}`,
      kind: 'shape',
      surface: 'powerpoint',
      title: `Shape ${shapeId}`,
      hostRef: { type: 'powerpoint.shape', slideId: shapeSlideId, shapeId },
      anchor: { matchText: '', locator: `shape:${shapeSlideId}:${shapeId}` },
    };
  }
  return undefined;
}

function onenoteLocationRef(target: string): ContextRef | undefined {
  const pageId = prefixedValue(target, 'on:page:', 'onenote:page:', 'page:');
  if (!pageId) return undefined;
  return {
    id: `on:page:${pageId}`,
    kind: 'page',
    surface: 'onenote',
    title: `Page ${pageId}`,
    hostRef: { type: 'onenote.page', pageId },
  };
}

function outlookLocationRef(target: string): ContextRef | undefined {
  const itemId = prefixedValue(target, 'outlook:item:', 'mail:item:', 'message:', 'item:');
  if (!itemId) return undefined;
  return {
    id: itemId,
    kind: 'mail-item',
    surface: 'outlook',
    title: `Message ${truncate(itemId, 24)}`,
    hostRef: { type: 'outlook.item', itemId },
  };
}

function teamsLocationRef(target: string): ContextRef | undefined {
  if (target === 'teams:transcript' || target === 'transcript') {
    return {
      id: 'teams:transcript',
      kind: 'transcript',
      surface: 'teams',
      title: 'Transcript',
    };
  }
  const url = prefixedValue(target, 'teams:link:', 'teams:deeplink:');
  if (!url || !safeHttpUrl(url)) return undefined;
  return {
    id: `teams:link:${hashId(url)}`,
    kind: 'transcript',
    surface: 'teams',
    title: 'Teams link',
    hostRef: { type: 'teams.deepLink', url },
  };
}

function prefixedValue(value: string, ...prefixes: string[]): string | undefined {
  const trimmed = value.trim();
  for (const prefix of prefixes) {
    if (trimmed.startsWith(prefix)) {
      const rest = trimmed.slice(prefix.length).trim();
      return rest || undefined;
    }
  }
  return undefined;
}

function safeHttpUrl(value: string): boolean {
  try {
    const parsed = new URL(value);
    return parsed.protocol === 'https:' || parsed.protocol === 'http:';
  } catch {
    return false;
  }
}

function truncate(value: string, max = 48): string {
  return value.length <= max ? value : `${value.slice(0, max - 1)}…`;
}

function hashId(value: string): string {
  let hash = 0;
  for (let i = 0; i < value.length; i++) {
    hash = (hash * 31 + value.charCodeAt(i)) >>> 0;
  }
  return hash.toString(36);
}
