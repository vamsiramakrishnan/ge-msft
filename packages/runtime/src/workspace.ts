import { renderValue, type Value } from './compose.js';

export type WorkspaceArtifactKind = 'text' | 'table' | 'json' | 'markdown' | 'tsv' | 'handoff';

export interface WorkspaceArtifact {
  id: string;
  name: string;
  kind: WorkspaceArtifactKind;
  mimeType: string;
  text: string;
  sourceLabel: string;
  createdAt: string;
  bytes: number;
  lineCount: number;
  truncated: boolean;
}

export interface WorkspaceArtifactSummary {
  id: string;
  name: string;
  kind: WorkspaceArtifactKind;
  mimeType: string;
  sourceLabel: string;
  createdAt: string;
  bytes: number;
  lineCount: number;
  truncated: boolean;
}

export type WorkspaceResult =
  | { workspace: 'list'; artifacts: WorkspaceArtifactSummary[] }
  | { workspace: 'summary'; artifact: WorkspaceArtifactSummary; preview: string }
  | { workspace: 'save'; artifact: WorkspaceArtifactSummary; preview: string }
  | { workspace: 'cat'; artifact: WorkspaceArtifactSummary; head: number; preview: string }
  | {
      workspace: 'grep';
      artifact: WorkspaceArtifactSummary;
      pattern: string;
      context: number;
      matches: Array<{ line: number; text: string }>;
      truncated: boolean;
    }
  | { workspace: 'cp'; artifact: WorkspaceArtifactSummary }
  | { workspace: 'mv'; artifact: WorkspaceArtifactSummary }
  | { workspace: 'rm'; name: string }
  | { workspace: 'share'; name: string; bytes: number }
  | { workspace: 'error'; error: string };

export interface SaveWorkspaceInput {
  name: string;
  sourceLabel: string;
  content: string | Value;
  kind?: WorkspaceArtifactKind;
}

const MAX_ARTIFACTS = 32;
const MAX_ARTIFACT_BYTES = 256 * 1024;
const PREVIEW_LINES = 24;
const MAX_CAT_LINES = 200;
const GREP_MATCH_CAP = 50;

export class WorkspaceStore {
  private readonly artifacts = new Map<string, WorkspaceArtifact>();
  private readonly aliases = new Map<string, string>();
  private nextId = 1;

  list(): WorkspaceArtifactSummary[] {
    return [...this.artifacts.values()].map(summaryOf);
  }

  get(ref: string): WorkspaceArtifact | undefined {
    const key = ref.trim();
    const id = this.artifacts.has(key) ? key : this.aliases.get(key);
    return id ? this.artifacts.get(id) : undefined;
  }

  save(input: SaveWorkspaceInput): WorkspaceResult {
    const rendered = renderContent(input.content);
    const capped = capText(rendered.text, MAX_ARTIFACT_BYTES);
    const id = `ws:${this.nextId++}`;
    const artifact: WorkspaceArtifact = {
      id,
      name: input.name,
      kind: input.kind ?? rendered.kind,
      mimeType: mimeTypeFor(input.kind ?? rendered.kind),
      text: capped.text,
      sourceLabel: input.sourceLabel,
      createdAt: new Date().toISOString(),
      bytes: byteLength(capped.text),
      lineCount: lineCount(capped.text),
      truncated: capped.truncated,
    };
    this.artifacts.set(id, artifact);
    this.aliases.set(input.name, id);
    this.evictOldest();
    return { workspace: 'save', artifact: summaryOf(artifact), preview: preview(artifact.text) };
  }

  summary(ref: string): WorkspaceResult {
    const artifact = this.get(ref);
    if (!artifact) return { workspace: 'error', error: `workspace artifact not found: ${ref}` };
    return { workspace: 'summary', artifact: summaryOf(artifact), preview: preview(artifact.text) };
  }

  cat(ref: string, head = PREVIEW_LINES): WorkspaceResult {
    const artifact = this.get(ref);
    if (!artifact) return { workspace: 'error', error: `workspace artifact not found: ${ref}` };
    const boundedHead = Math.max(1, Math.min(head, MAX_CAT_LINES));
    return {
      workspace: 'cat',
      artifact: summaryOf(artifact),
      head: boundedHead,
      preview: artifact.text.split(/\r?\n/).slice(0, boundedHead).join('\n'),
    };
  }

  grep(ref: string, pattern: string, context = 0): WorkspaceResult {
    const artifact = this.get(ref);
    if (!artifact) return { workspace: 'error', error: `workspace artifact not found: ${ref}` };
    const lines = artifact.text.split(/\r?\n/);
    const needle = pattern.toLowerCase();
    const selected = new Map<number, string>();
    for (let i = 0; i < lines.length; i++) {
      if (!lines[i]!.toLowerCase().includes(needle)) continue;
      for (let j = Math.max(0, i - context); j <= Math.min(lines.length - 1, i + context); j++) {
        selected.set(j + 1, lines[j]!);
      }
      if (selected.size >= GREP_MATCH_CAP) break;
    }
    const matches = [...selected.entries()]
      .slice(0, GREP_MATCH_CAP)
      .map(([line, text]) => ({ line, text }));
    return {
      workspace: 'grep',
      artifact: summaryOf(artifact),
      pattern,
      context,
      matches,
      truncated: selected.size > GREP_MATCH_CAP,
    };
  }

  /**
   * Duplicate an artifact under a new name/alias with a FRESH id — a copy is a distinct artifact
   * for eviction purposes (see `evictOldest`), so it is inserted like any `save()`d artifact.
   * `dst` is last-write-wins, matching `save()`'s `this.aliases.set(input.name, id)` semantics: an
   * existing `dst` alias is silently overwritten to point at the new copy, orphaning whatever it
   * previously named (recoverable only by its own `ws:id`, exactly like re-saving an existing name).
   */
  cp(src: string, dst: string): WorkspaceResult {
    const source = this.get(src);
    if (!source) return { workspace: 'error', error: `workspace artifact not found: ${src}` };
    const id = `ws:${this.nextId++}`;
    const copy: WorkspaceArtifact = {
      ...source,
      id,
      name: dst,
      createdAt: new Date().toISOString(),
    };
    this.artifacts.set(id, copy);
    this.aliases.set(dst, id);
    this.evictOldest();
    return { workspace: 'cp', artifact: summaryOf(copy) };
  }

  /**
   * Rename an artifact IN PLACE — same id, no new `artifacts` entry (so it keeps its original
   * insertion-order age for `evictOldest`). `get()` resolves `src` by id OR alias, so the alias to
   * drop is the artifact's OWN current `name` field, not necessarily the literal `src` string (which
   * may itself have been a `ws:id`). Only drop that alias if it still points at this artifact — a
   * `save()` that reused the same name for a DIFFERENT artifact since would have already repointed
   * it, and this rename must not un-repoint someone else's alias. `dst` is last-write-wins, exactly
   * like `save()`/`cp()`.
   */
  mv(src: string, dst: string): WorkspaceResult {
    const artifact = this.get(src);
    if (!artifact) return { workspace: 'error', error: `workspace artifact not found: ${src}` };
    if (this.aliases.get(artifact.name) === artifact.id) {
      this.aliases.delete(artifact.name);
    }
    artifact.name = dst;
    this.aliases.set(dst, artifact.id);
    return { workspace: 'mv', artifact: summaryOf(artifact) };
  }

  /** Delete an artifact so it no longer resolves by name OR id. Same alias-ownership care as `mv`. */
  rm(name: string): WorkspaceResult {
    const artifact = this.get(name);
    if (!artifact) return { workspace: 'error', error: `workspace artifact not found: ${name}` };
    this.artifacts.delete(artifact.id);
    if (this.aliases.get(artifact.name) === artifact.id) {
      this.aliases.delete(artifact.name);
    }
    return { workspace: 'rm', name: artifact.name };
  }

  private evictOldest(): void {
    while (this.artifacts.size > MAX_ARTIFACTS) {
      const oldest = this.artifacts.values().next().value as WorkspaceArtifact | undefined;
      if (!oldest) return;
      this.artifacts.delete(oldest.id);
      if (this.aliases.get(oldest.name) === oldest.id) this.aliases.delete(oldest.name);
    }
  }
}

function renderContent(content: string | Value): { text: string; kind: WorkspaceArtifactKind } {
  if (typeof content === 'string') return { text: content, kind: guessKindFromText(content) };
  if (content.kind === 'table') return { text: renderValue(content), kind: 'table' };
  if (content.kind === 'number') return { text: String(content.value), kind: 'text' };
  return { text: content.value, kind: guessKindFromText(content.value) };
}

function guessKindFromText(text: string): WorkspaceArtifactKind {
  const trimmed = text.trimStart();
  if (trimmed.startsWith('{') || trimmed.startsWith('[')) return 'json';
  if (/^\|.*\|\s*$/m.test(text)) return 'markdown';
  if (text.includes('\t')) return 'tsv';
  return 'text';
}

function mimeTypeFor(kind: WorkspaceArtifactKind): string {
  switch (kind) {
    case 'json':
      return 'application/json';
    case 'markdown':
    case 'table':
      return 'text/markdown';
    case 'tsv':
      return 'text/tab-separated-values';
    case 'handoff':
      return 'application/vnd.ge.handoff+json';
    case 'text':
      return 'text/plain';
  }
}

function summaryOf(artifact: WorkspaceArtifact): WorkspaceArtifactSummary {
  const { id, name, kind, mimeType, sourceLabel, createdAt, bytes, lineCount, truncated } =
    artifact;
  return { id, name, kind, mimeType, sourceLabel, createdAt, bytes, lineCount, truncated };
}

function capText(text: string, maxBytes: number): { text: string; truncated: boolean } {
  if (byteLength(text) <= maxBytes) return { text, truncated: false };
  const chars = [...text];
  let lo = 0;
  let hi = chars.length;
  while (lo < hi) {
    const mid = Math.ceil((lo + hi) / 2);
    if (byteLength(chars.slice(0, mid).join('')) <= maxBytes) lo = mid;
    else hi = mid - 1;
  }
  return {
    text: `${chars.slice(0, lo).join('')}\n[workspace artifact truncated at ${maxBytes} bytes]`,
    truncated: true,
  };
}

function preview(text: string): string {
  const lines = text.split(/\r?\n/);
  const shown = lines.slice(0, PREVIEW_LINES).join('\n');
  return lines.length > PREVIEW_LINES
    ? `${shown}\n... ${lines.length - PREVIEW_LINES} more line(s)`
    : shown;
}

function byteLength(text: string): number {
  return new TextEncoder().encode(text).length;
}

function lineCount(text: string): number {
  if (text === '') return 0;
  return text.split(/\r?\n/).length;
}
