import type { FileView } from '@ge/contracts';
import { byteLength, truncateToBytes } from '../docfs/bytes.js';
import type { Mount } from '../docfs/mount.js';
import { EvidenceWorkspace } from './workspace.js';

/** Existing DocFs verbs reach immutable evidence without a feature-specific command grammar. */
export function evidenceMount(workspace: EvidenceWorkspace): Mount {
  const read = async (
    rel: string,
    maxBytes = 64 * 1024,
  ): Promise<(FileView & { sourceBytes: number }) | null> => {
    if (!Number.isSafeInteger(maxBytes) || maxBytes < 1 || maxBytes > 96 * 1024)
      throw new Error('Invalid evidence read budget.');
    const [id, kind, group, leaf] = rel.split('/');
    if (!id || !(await workspace.list()).includes(id)) return null;
    const receipt = await workspace.receipt(id);
    let text: string;
    if (kind === 'manifest.json' && group === undefined) text = JSON.stringify(receipt);
    else if (kind === 'windows') {
      if (
        !group ||
        !/^\d+$/.test(group) ||
        !leaf ||
        !/^\d+\.json$/.test(leaf) ||
        rel.split('/').length !== 4
      )
        return null;
      if (Number(group) >= receipt.unitCount) return null;
      const unit = await workspace.readAt(id, Number(group));
      if (!unit) return null;
      const window = await workspace.readText(id, unit.id, {
        offset: Number(leaf.slice(0, -5)),
        maxBytes: 4096,
      });
      text = JSON.stringify({
        ...window,
        nextPath:
          window.nextOffset === undefined
            ? undefined
            : `/evidence/${id}/windows/${group}/${window.nextOffset}.json`,
      });
    } else {
      if (
        kind !== 'units' ||
        !group ||
        !/^\d+$/.test(group) ||
        !leaf ||
        !/^\d+\.json$/.test(leaf) ||
        rel.split('/').length !== 4
      )
        return null;
      const index = Number(leaf.slice(0, -5));
      if (Math.floor(index / 100) !== Number(group)) return null;
      if (index >= receipt.unitCount) return null;
      const unit = await workspace.readAt(id, index);
      if (!unit) return null;
      text = JSON.stringify({ textWindowPath: `/evidence/${id}/windows/${index}/0.json`, ...unit });
    }
    const bounded = truncateToBytes(text, maxBytes);
    return {
      path: '',
      text: bounded.text,
      bytes: byteLength(bounded.text),
      truncated: bounded.truncated,
      sourceBytes: byteLength(text),
    };
  };
  return {
    prefix: 'evidence',
    readonly: true,
    async readdir(rel) {
      if (!rel) return (await workspace.list()).map((name) => ({ name, kind: 'dir' }));
      const [id, kind, group] = rel.split('/');
      if (!id || !(await workspace.list()).includes(id)) return [];
      const snapshot = await workspace.receipt(id);
      if (!kind)
        return [
          { name: 'manifest.json', kind: 'file' },
          { name: 'units', kind: 'dir' },
        ];
      if (kind !== 'units') return [];
      if (group === undefined)
        return Array.from({ length: Math.ceil(snapshot.unitCount / 100) }, (_, i) => ({
          name: String(i),
          kind: 'dir',
        }));
      if (!/^\d+$/.test(group) || rel.split('/').length !== 3) return [];
      const start = Number(group) * 100;
      return Array.from(
        { length: Math.max(0, Math.min(100, snapshot.unitCount - start)) },
        (_, i) => ({ name: `${start + i}.json`, kind: 'file' }),
      );
    },
    async stat(rel) {
      const value = await read(rel);
      return value ? { path: '', kind: 'file', size: value.sourceBytes } : null;
    },
    async readFile(rel, options) {
      const value = await read(rel, options?.maxBytes);
      if (!value) return null;
      const { sourceBytes: _sourceBytes, ...view } = value;
      return view;
    },
    async search(rel, pattern, options) {
      const max = options?.max ?? 20;
      if (!Number.isSafeInteger(max) || max < 1 || max > 100)
        throw new Error('Invalid evidence search limit.');
      const id = rel.split('/')[0];
      if (!id || !(await workspace.list()).includes(id) || !pattern.trim()) return [];
      const prefix = rel.replace(/\/$/, '');
      const hits: Array<{ path: string; line: number; text: string }> = [];
      let cursor: string | undefined;
      do {
        const page = await workspace.query(
          id,
          { kind: 'grep', literal: pattern, caseSensitive: true },
          { cursor, limit: 100 },
        );
        for (const node of page.nodes) {
          const path = node.path.replace(/^\/evidence\//, '');
          if (path !== prefix && !path.startsWith(prefix + '/')) continue;
          const unit = await workspace.read(id, node.id);
          const offset = node.matchOffset ?? 0;
          hits.push({
            path,
            line: unit.text.slice(0, offset).split('\n').length,
            text: truncateToBytes(unit.text.slice(offset), 512).text,
          });
          if (hits.length >= max) return hits;
        }
        cursor = page.nextCursor;
      } while (cursor);
      return hits;
    },
  };
}
