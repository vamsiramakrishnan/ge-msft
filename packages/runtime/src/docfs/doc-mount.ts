// packages/runtime/src/docfs/doc-mount.ts
import type {
  DirEntry,
  FileView,
  ReadOpts,
  ResolvedContext,
  SearchMatch,
  SearchOpts,
} from '@ge/contracts';
import type { DocBridge } from '../bridge.js';
import { byteLength, truncateToBytes } from './bytes.js';
import type { Mount } from './mount.js';

const DEFAULT_MAX_BYTES = 64 * 1024;

/** `/doc` — the live document as a lazy, read-only index: readdir = the doc-state inventory/outline,
 *  readFile/search = on-demand reads via the bridge. Surface-agnostic (DocBridge only). The caller
 *  wraps this with `readOnly()` (see mount.ts) when assembling the DocFs; this mount itself has no
 *  write methods to begin with. */
export function docMount(bridge: DocBridge): Mount {
  return {
    prefix: 'doc',
    async readdir(): Promise<DirEntry[]> {
      const snapshot = await bridge.captureDocState?.();
      const entries: DirEntry[] = [{ name: 'outline.md', kind: 'file' }];
      if (!snapshot) return entries;
      if (snapshot.selection) entries.push({ name: 'selection', kind: 'file' });
      if (snapshot.namedRanges?.length) entries.push({ name: 'named-ranges.json', kind: 'file' });
      for (const entry of snapshot.inventory) entries.push({ name: entry.id, kind: 'file' });
      return entries;
    },
    async stat(rel) {
      const entries = await this.readdir('');
      const entry = entries.find((e) => e.name === rel);
      return entry ? { path: '', kind: 'file' as const, size: entry.size ?? 0 } : null;
    },
    async readFile(rel, opts?: ReadOpts): Promise<FileView | null> {
      const max = opts?.maxBytes ?? DEFAULT_MAX_BYTES;
      const snapshot = await bridge.captureDocState?.();

      if (rel === 'outline.md') {
        const outline = snapshot?.outline ?? [];
        const text = outline.map((o) => `${'#'.repeat(Math.max(1, o.level))} ${o.text}`).join('\n');
        return cap(text, max);
      }

      // Everything below needs a real snapshot to have anything to serve.
      if (!snapshot) return null;

      if (rel === 'selection') {
        if (!snapshot.selection) return null;
        return cap(snapshot.selection.preview ?? snapshot.selection.title, max);
      }

      if (rel === 'named-ranges.json') {
        if (!snapshot.namedRanges) return null;
        return cap(JSON.stringify(snapshot.namedRanges, null, 2), max);
      }

      // An inventory id (Phase 1 simplification: pass rel straight through as the address).
      const parts = (await bridge.readRange?.(rel)) ?? [];
      if (parts.length === 0) return null;
      return cap(parts.map(renderResolved).join('\n'), max);
    },
    async search(_rel, pattern, opts?: SearchOpts): Promise<SearchMatch[]> {
      const parts = (await bridge.searchDocument?.(pattern)) ?? [];
      const max = opts?.max ?? 50;
      return parts.slice(0, max).map(
        (part, i): SearchMatch => ({
          path: part.ref?.id ?? `hit:${i}`,
          line: 1,
          text: renderResolved(part),
        }),
      );
    },
  };
}

function cap(text: string, max: number): FileView {
  const { text: truncated, truncated: wasTruncated } = truncateToBytes(text, max);
  return {
    path: '',
    text: truncated,
    bytes: byteLength(truncated),
    truncated: wasTruncated,
  };
}

/** Render a ResolvedContext value to a text line for DocFs views (data only, never instructions). */
function renderResolved(part: ResolvedContext): string {
  return part.value.as === 'text' ? part.value.text : JSON.stringify(part.value).slice(0, 500);
}
