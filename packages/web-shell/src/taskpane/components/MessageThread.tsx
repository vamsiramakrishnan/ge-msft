import { parseEmphasis, type EmphasisInput, type EmphasisNode } from './emphasis.js';
import { useState } from 'react';
import type { ReactNode } from 'react';
import type { ChatMessage, ResultChange } from '../../controller.js';
import { uiStrings, type UiLanguage } from '../../ui-language.js';
import {
  ALL_VERBS,
  CONFIRMED_PLAN_OPEN,
  isCommandParseError,
  parseCommandLine,
  type SourceRef,
  type Surface,
} from '@ge/contracts';
import type { InsertableArtifact } from '../insert-artifact.js';
import { canRenderHostLocation } from '../../host-location.js';
import { copyText } from './RunSteps.js';
import { answerForClipboard } from '../answer-copy.js';

export interface MessageThreadProps {
  messages: ChatMessage[];
  showEmpty?: boolean;
  disabled?: boolean;
  surface?: Surface;
  onInsertArtifact?: (artifact: InsertableArtifact) => void;
  onRevealLocation?: (location: string) => void;
  insertArtifactDisabledReason?: string;
  /** The conversation's language, for the pane's own text. */
  language?: UiLanguage;
  /** Undo what a result card lists (through the approval card). */
  onUndoChanges?: (messageId: string) => void;
  /** Send the last request again (shown on the latest failed answer only). */
  onRetry?: () => void;
}

/**
 * Citation URIs come from grounded, untrusted source material. Only render an http(s) link;
 * anything else (e.g. `javascript:` / `data:`) is rendered as inert text so a malicious source
 * cannot smuggle an executable href into the panel. Host content is data, never instructions.
 */
function safeHttpUri(uri: string | undefined): string | undefined {
  if (!uri) return undefined;
  try {
    const parsed = new URL(uri, window.location.origin);
    return parsed.protocol === 'http:' || parsed.protocol === 'https:' ? parsed.href : undefined;
  } catch {
    return undefined;
  }
}

function citationLocation(uri: string | undefined): string | undefined {
  let target = uri?.trim();
  if (!target) return undefined;
  if (target.startsWith('<') && target.endsWith('>')) target = target.slice(1, -1).trim();
  return target.toLowerCase().startsWith('citation:') ? target : undefined;
}

function splitTableRow(line: string): string[] {
  const trimmed = line.trim().replace(/^\|/, '').replace(/\|$/, '');
  return trimmed.split('|').map((cell) => cell.trim());
}

function isTableSeparator(line: string): boolean {
  const cells = splitTableRow(line);
  return cells.length > 1 && cells.every((cell) => /^:?-{3,}:?$/.test(cell));
}

function isTableStart(lines: string[], index: number): boolean {
  const current = lines[index];
  const next = lines[index + 1];
  return Boolean(current && next && current.includes('|') && isTableSeparator(next));
}

function isBlockStart(lines: string[], index: number): boolean {
  const line = lines[index] ?? '';
  const trimmed = line.trim();
  return (
    trimmed === '' ||
    trimmed.startsWith('```') ||
    isTableStart(lines, index) ||
    /^(#{1,4})\s+/.test(trimmed) ||
    /^([-*_])\1\1+$/.test(trimmed) ||
    /^[-*+]\s+/.test(trimmed) ||
    /^\d+\.\s+/.test(trimmed)
  );
}

function findNextInlineMarker(text: string, start: number): number {
  const markers = ['`', '['];
  const indexes = markers.map((marker) => text.indexOf(marker, start)).filter((idx) => idx >= 0);
  return indexes.length ? Math.min(...indexes) : text.length;
}

function renderInline(
  text: string,
  keyPrefix: string,
  surface: Surface,
  onRevealLocation?: (location: string) => void,
): ReactNode[] {
  // Code spans and links first (atoms), then emphasis by the CommonMark rules over what is left.
  const pieces: EmphasisInput<ReactNode> = [];
  let i = 0;
  let key = 0;
  const pushText = (value: string): void => {
    if (!value) return;
    const last = pieces[pieces.length - 1];
    if (last?.kind === 'text') last.text += value;
    else pieces.push({ kind: 'text', text: value });
  };

  while (i < text.length) {
    const atoms: ReactNode[] = [];
    if (text[i] === '`') {
      const end = text.indexOf('`', i + 1);
      if (end > i + 1) {
        const codeText = text.slice(i + 1, end);
        atoms.push(
          onRevealLocation && canRenderHostLocation(surface, codeText) ? (
            <button
              key={`${keyPrefix}-loc-${key++}`}
              type="button"
              className="md-host-location"
              onClick={() => onRevealLocation(codeText)}
              title="Open this location in the host"
            >
              {codeText}
            </button>
          ) : (
            <code key={`${keyPrefix}-c-${key++}`}>{codeText}</code>
          ),
        );
        pieces.push({ kind: 'atom', atom: atoms[0] });
        i = end + 1;
        continue;
      }
    }

    if (text[i] === '[') {
      const labelEnd = text.indexOf(']', i + 1);
      const urlStart = labelEnd >= 0 && text[labelEnd + 1] === '(' ? labelEnd + 2 : -1;
      const urlEnd = urlStart >= 0 ? text.indexOf(')', urlStart) : -1;
      if (labelEnd > i + 1 && urlStart >= 0 && urlEnd > urlStart) {
        const label = text.slice(i + 1, labelEnd);
        const rawTarget = text.slice(urlStart, urlEnd).trim();
        const hostLocation = citationLocation(rawTarget);
        const href = safeHttpUri(rawTarget);
        atoms.push(
          hostLocation && onRevealLocation && canRenderHostLocation(surface, hostLocation) ? (
            <button
              key={`${keyPrefix}-host-a-${key++}`}
              type="button"
              className="md-host-location"
              onClick={() => onRevealLocation(hostLocation)}
              title="Open this location in the host"
            >
              {renderInline(label, `${keyPrefix}-host-a-${key}`, surface, onRevealLocation)}
            </button>
          ) : href ? (
            <a
              key={`${keyPrefix}-a-${key++}`}
              className="md-link"
              href={href}
              target="_blank"
              rel="noreferrer noopener"
            >
              {renderInline(label, `${keyPrefix}-a-${key}`, surface, onRevealLocation)}
            </a>
          ) : (
            <span key={`${keyPrefix}-bad-a-${key++}`}>
              {label} ({rawTarget})
            </span>
          ),
        );
        pieces.push({ kind: 'atom', atom: atoms[0] });
        i = urlEnd + 1;
        continue;
      }
    }

    const next = findNextInlineMarker(text, i + 1);
    pushText(text.slice(i, next));
    i = next;
  }

  const render = (nodes: EmphasisNode<ReactNode>[]): ReactNode[] =>
    nodes.map((node) => {
      if (node.kind === 'text') return node.text;
      if (node.kind === 'atom') return node.atom;
      const Tag = node.kind === 'em' ? 'em' : 'strong';
      return <Tag key={`${keyPrefix}-${node.kind}-${key++}`}>{render(node.children)}</Tag>;
    });
  return render(parseEmphasis(pieces));
}

const USER_TOKEN_RE =
  /(^|\s)(\/[a-z][\w-]*|@[a-z][\w-]*(?::[^\s]+)?|scope:[a-z-]+(?:\([^)]+\))?)/gi;

function userTokenClass(token: string): string {
  if (token.startsWith('/')) return 'msg-token msg-token-command';
  if (token.startsWith('@')) return 'msg-token msg-token-mention';
  return 'msg-token msg-token-scope';
}

function userTokenLabel(token: string): string {
  if (token.startsWith('/')) return `Command ${token}`;
  if (token.startsWith('@')) return `Grounding ${token}`;
  return `Scope ${token.replace(/^scope:/, '')}`;
}

interface UserTextSegment {
  kind: 'text';
  text: string;
}

interface UserCommandSegment {
  kind: 'command';
  verb: string;
  target?: string;
  summary: string;
  preview?: string[][];
  hiddenDetail?: string;
}

interface InternalPlanSegment {
  kind: 'internal-plan';
  intent?: string;
  surface?: string;
  scope?: string;
  steps: string[];
}

type UserSegment = UserTextSegment | UserCommandSegment | InternalPlanSegment;

const COMMAND_LINE_RE =
  /^(outline|read|search|context|list|inspect|properties|comments|attachments|tables|slides|neighbors|open|set|grid|suggest|comment|format|reply|slide|page|mail|post|compose|table|chart|cf|spill|done)\b/i;

function UserText({ text }: { text: string }): JSX.Element {
  const nodes: ReactNode[] = [];
  let last = 0;
  let key = 0;
  USER_TOKEN_RE.lastIndex = 0;

  for (const match of text.matchAll(USER_TOKEN_RE)) {
    const whole = match[0] ?? '';
    const prefix = match[1] ?? '';
    const token = match[2] ?? '';
    const tokenStart = (match.index ?? 0) + prefix.length;
    if (tokenStart > last) nodes.push(text.slice(last, tokenStart));
    nodes.push(
      <span key={`tok-${key++}`} className={userTokenClass(token)} title={userTokenLabel(token)}>
        {token}
      </span>,
    );
    last = tokenStart + token.length;
    // Preserve any whitespace prefix in the normal text flow; the token starts after it.
    if (whole.length > prefix.length + token.length) nodes.push(whole.slice(prefix.length));
  }

  if (last < text.length) nodes.push(text.slice(last));
  return <>{nodes}</>;
}

function unescapedQuoteCount(text: string): number {
  let count = 0;
  let escaped = false;
  for (const ch of text) {
    if (escaped) {
      escaped = false;
      continue;
    }
    if (ch === '\\') {
      escaped = true;
      continue;
    }
    if (ch === '"') count++;
  }
  return count;
}

function decodeCommandString(value: string): string {
  return value
    .replace(/\\t/g, '\t')
    .replace(/\\n/g, '\n')
    .replace(/\\"/g, '"')
    .replace(/\\\\/g, '\\');
}

function extractQuotedBody(text: string): { body: string; rest: string } | undefined {
  const start = text.indexOf('"');
  if (start < 0) return undefined;
  let escaped = false;
  let body = '';
  for (let i = start + 1; i < text.length; i++) {
    const ch = text[i] ?? '';
    if (escaped) {
      body += `\\${ch}`;
      escaped = false;
      continue;
    }
    if (ch === '\\') {
      escaped = true;
      continue;
    }
    if (ch === '"') return { body: decodeCommandString(body), rest: text.slice(i + 1).trim() };
    body += ch;
  }
  return { body: decodeCommandString(body), rest: '' };
}

function gridSegment(raw: string): UserCommandSegment | undefined {
  const match = /^grid\s+(.+?)\s*=\s*([\s\S]+)$/i.exec(raw.trim());
  if (!match) return undefined;
  const target = match[1]?.trim();
  const quoted = extractQuotedBody(match[2] ?? '');
  const body = quoted?.body ?? '';
  const rows = body
    .split('\n')
    .map((row) => row.split('\t'))
    .filter((row) => row.some((cell) => cell.trim()));
  const cols = rows.reduce((max, row) => Math.max(max, row.length), 0);
  const hiddenDetail =
    rows.length > 3 || cols > 4
      ? `${Math.max(rows.length - 3, 0)} more rows · ${Math.max(cols - 4, 0)} more columns hidden`
      : undefined;
  return {
    kind: 'command',
    verb: 'grid',
    ...(target ? { target } : {}),
    summary: `${rows.length} x ${cols} cells`,
    preview: rows.slice(0, 3).map((row) => row.slice(0, 4)),
    ...(hiddenDetail ? { hiddenDetail } : {}),
  };
}

function commandSegment(raw: string): UserCommandSegment | undefined {
  const trimmed = raw.trim();
  const verb = COMMAND_LINE_RE.exec(trimmed)?.[1]?.toLowerCase();
  if (!verb) return undefined;
  if (verb === 'grid') return gridSegment(trimmed);
  if (verb === 'read') {
    const target = trimmed.replace(/^read\s+/i, '').trim();
    return {
      kind: 'command',
      verb: 'read',
      ...(target ? { target } : {}),
      summary: 'host read',
    };
  }
  if (verb === 'done') {
    return { kind: 'command', verb, summary: 'finish command loop' };
  }
  return {
    kind: 'command',
    verb,
    summary: trimmed.length > 120 ? `${trimmed.slice(0, 117)}...` : trimmed,
  };
}

function internalPlanSegment(text: string): InternalPlanSegment | undefined {
  if (!text.includes(CONFIRMED_PLAN_OPEN)) return undefined;
  const lines = text.replace(/\r\n/g, '\n').split('\n');
  const value = (key: string): string | undefined => {
    const prefix = `${key}:`;
    return lines
      .find((line) => line.trim().toLowerCase().startsWith(prefix))
      ?.trim()
      .slice(prefix.length)
      .trim();
  };
  const steps = lines
    .map((line) => /^step\s+\d+:\s*(.+)$/i.exec(line.trim())?.[1]?.trim())
    .filter((line): line is string => Boolean(line));
  const intent = value('intent');
  const surface = value('surface');
  const scope = value('scope');
  return {
    kind: 'internal-plan',
    ...(intent ? { intent } : {}),
    ...(surface ? { surface } : {}),
    ...(scope ? { scope } : {}),
    steps,
  };
}

function parseUserSegments(text: string): UserSegment[] {
  const internal = internalPlanSegment(text);
  if (internal) return [internal];

  const segments: UserSegment[] = [];
  const lines = text.replace(/\r\n/g, '\n').split('\n');
  const textBuffer: string[] = [];

  const flushText = (): void => {
    const value = textBuffer.join('\n').trim();
    textBuffer.length = 0;
    if (value) segments.push({ kind: 'text', text: value });
  };

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i] ?? '';
    const trimmed = line.trim();
    if (!COMMAND_LINE_RE.test(trimmed)) {
      textBuffer.push(line);
      continue;
    }

    flushText();
    let raw = line;
    if (/^grid\b/i.test(trimmed) && unescapedQuoteCount(raw) % 2 === 1) {
      while (i + 1 < lines.length && unescapedQuoteCount(raw) % 2 === 1) {
        i++;
        raw += `\n${lines[i] ?? ''}`;
      }
    }
    const command = commandSegment(raw);
    if (command) segments.push(command);
    else textBuffer.push(raw);
  }

  flushText();
  return segments.length > 0 ? segments : [{ kind: 'text', text }];
}

function CommandCard({ segment }: { segment: UserCommandSegment }): JSX.Element {
  return (
    <div
      className={`cmd-card cmd-card-${segment.verb}`}
      role="group"
      aria-label={`${segment.verb} command`}
    >
      <div className="cmd-card-head">
        <span className="cmd-card-verb">{segment.verb}</span>
        <span className="cmd-card-summary">{segment.summary}</span>
      </div>
      {segment.target && <div className="cmd-card-target">{segment.target}</div>}
      {segment.preview && segment.preview.length > 0 && (
        <div className="cmd-card-grid-preview" aria-label="Grid preview">
          {segment.preview.map((row, rowIdx) => (
            <div key={rowIdx} className="cmd-card-grid-row">
              {row.map((cell, cellIdx) => (
                <span key={cellIdx} className="cmd-card-cell">
                  {cell || '\u00a0'}
                </span>
              ))}
            </div>
          ))}
        </div>
      )}
      {segment.hiddenDetail && <div className="cmd-card-muted">{segment.hiddenDetail}</div>}
    </div>
  );
}

function InternalPlanCard({ segment }: { segment: InternalPlanSegment }): JSX.Element {
  const title = [
    segment.intent ? `${segment.intent} plan` : 'approved plan',
    segment.surface,
    segment.scope,
  ]
    .filter(Boolean)
    .join(' · ');
  return (
    <div className="cmd-card cmd-card-internal" role="group" aria-label="Approved plan execution">
      <div className="cmd-card-head">
        <span className="cmd-card-verb">execute</span>
        <span className="cmd-card-summary">{title}</span>
      </div>
      {segment.steps.length > 0 && (
        <ol className="cmd-card-steps">
          {segment.steps.slice(0, 3).map((step, idx) => (
            <li key={idx}>{step}</li>
          ))}
        </ol>
      )}
      {segment.steps.length > 3 && (
        <div className="cmd-card-muted">{segment.steps.length - 3} more steps hidden</div>
      )}
    </div>
  );
}

function UserMessageContent({ text }: { text: string }): JSX.Element {
  const segments = parseUserSegments(text);
  return (
    <>
      {segments.map((segment, idx) => {
        if (segment.kind === 'text') {
          return (
            <span key={idx} className="user-text-fragment">
              <UserText text={segment.text} />
            </span>
          );
        }
        if (segment.kind === 'internal-plan') {
          return <InternalPlanCard key={idx} segment={segment} />;
        }
        return <CommandCard key={idx} segment={segment} />;
      })}
    </>
  );
}

function ArtifactActions({
  artifact,
  onInsertArtifact,
  disabledReason,
}: {
  artifact: InsertableArtifact;
  onInsertArtifact?: (artifact: InsertableArtifact) => void;
  disabledReason?: string;
}): JSX.Element | null {
  if (!onInsertArtifact) return null;
  const disabled = Boolean(disabledReason);
  return (
    <div className="md-artifact-actions">
      <button
        type="button"
        className="md-artifact-insert"
        disabled={disabled}
        title={disabledReason ?? 'Insert into the current Microsoft 365 surface'}
        onClick={() => {
          if (!disabled) onInsertArtifact(artifact);
        }}
      >
        Insert
      </button>
    </div>
  );
}

function MarkdownContent({
  text,
  surface,
  onInsertArtifact,
  onRevealLocation,
  insertArtifactDisabledReason,
}: {
  text: string;
  surface: Surface;
  onInsertArtifact?: (artifact: InsertableArtifact) => void;
  onRevealLocation?: (location: string) => void;
  insertArtifactDisabledReason?: string;
}): JSX.Element {
  const lines = text.replace(/\r\n/g, '\n').split('\n');
  const blocks: ReactNode[] = [];
  let i = 0;
  let block = 0;

  while (i < lines.length) {
    const raw = lines[i] ?? '';
    const trimmed = raw.trim();
    if (!trimmed) {
      i++;
      continue;
    }

    if (trimmed.startsWith('```')) {
      const code: string[] = [];
      i++;
      while (i < lines.length && !(lines[i] ?? '').trim().startsWith('```')) {
        code.push(lines[i] ?? '');
        i++;
      }
      if (i < lines.length) i++;
      const artifact: InsertableArtifact = {
        kind: 'code-block',
        title: 'Inserted content',
        code: code.join('\n'),
      };
      blocks.push(
        <div key={`code-${block++}`} className="md-artifact md-code-artifact">
          <ArtifactActions
            artifact={artifact}
            onInsertArtifact={onInsertArtifact}
            disabledReason={insertArtifactDisabledReason}
          />
          <pre className="md-code">
            <code>{code.join('\n')}</code>
          </pre>
        </div>,
      );
      continue;
    }

    if (isTableStart(lines, i)) {
      const headerLine = lines[i] ?? '';
      const headers = splitTableRow(headerLine);
      i += 2;
      const rows: string[][] = [];
      while (i < lines.length && (lines[i] ?? '').includes('|') && (lines[i] ?? '').trim()) {
        const row = splitTableRow(lines[i] ?? '');
        if (row.length === headers.length) rows.push(row);
        i++;
      }
      const artifact: InsertableArtifact = {
        kind: 'markdown-table',
        title: 'Inserted table',
        headers,
        rows,
      };
      blocks.push(
        <div key={`table-wrap-${block++}`} className="md-artifact md-table-artifact">
          <ArtifactActions
            artifact={artifact}
            onInsertArtifact={onInsertArtifact}
            disabledReason={insertArtifactDisabledReason}
          />
          <div className="md-table-wrap">
            <table className="md-table">
              <thead>
                <tr>
                  {headers.map((cell, idx) => (
                    <th key={idx}>
                      {renderInline(cell, `th-${block}-${idx}`, surface, onRevealLocation)}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {rows.map((row, rowIdx) => (
                  <tr key={rowIdx}>
                    {row.map((cell, cellIdx) => (
                      <td key={cellIdx}>
                        {renderInline(
                          cell,
                          `td-${block}-${rowIdx}-${cellIdx}`,
                          surface,
                          onRevealLocation,
                        )}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>,
      );
      continue;
    }

    const heading = /^(#{1,4})\s+(.+)$/.exec(trimmed);
    if (heading) {
      const marker = heading[1] ?? '#';
      const content = heading[2] ?? '';
      const level = Math.min(marker.length + 1, 4);
      const Tag = `h${level}` as 'h2' | 'h3' | 'h4';
      blocks.push(
        <Tag key={`h-${block++}`}>
          {renderInline(content, `h-${block}`, surface, onRevealLocation)}
        </Tag>,
      );
      i++;
      continue;
    }

    if (/^([-*_])\1\1+$/.test(trimmed)) {
      blocks.push(<hr key={`hr-${block++}`} />);
      i++;
      continue;
    }

    const unordered = /^[-*+]\s+/.test(trimmed);
    const ordered = /^\d+\.\s+/.test(trimmed);
    if (unordered || ordered) {
      const items: string[] = [];
      const itemRegex = unordered ? /^[-*+]\s+/ : /^\d+\.\s+/;
      while (i < lines.length && itemRegex.test((lines[i] ?? '').trim())) {
        items.push((lines[i] ?? '').trim().replace(itemRegex, ''));
        i++;
      }
      const List = ordered ? 'ol' : 'ul';
      blocks.push(
        <List key={`list-${block++}`}>
          {items.map((item, idx) => (
            <li key={idx}>{renderInline(item, `li-${block}-${idx}`, surface, onRevealLocation)}</li>
          ))}
        </List>,
      );
      continue;
    }

    const paragraph: string[] = [];
    while (i < lines.length && !isBlockStart(lines, i)) {
      paragraph.push((lines[i] ?? '').trim());
      i++;
    }
    if (paragraph.length > 0) {
      blocks.push(
        <p key={`p-${block++}`}>
          {renderInline(paragraph.join(' '), `p-${block}`, surface, onRevealLocation)}
        </p>,
      );
      continue;
    }

    blocks.push(
      <p key={`fallback-${block++}`}>
        {renderInline(trimmed, `fallback-${block}`, surface, onRevealLocation)}
      </p>,
    );
    i++;
  }

  return <div className="md-content">{blocks}</div>;
}

/**
 * A citation pill that opens a source-detail popover (title + uri + locator) so the user can drill
 * into a grounding source without leaving the thread. The pill is a `button` controlling the detail;
 * the detail's link is http(s)-only (untrusted source URIs are otherwise rendered inert), and any
 * external link opens in a new tab. Keyboard-operable and labelled.
 */
function Citation({ source, id }: { source: SourceRef; id: string }): JSX.Element {
  const [open, setOpen] = useState(false);
  const label = source.locator ? `${source.title} · ${source.locator}` : source.title;
  const href = safeHttpUri(source.uri);
  const detailId = `cite-detail-${id}`;
  return (
    <span className="cite-wrap">
      <button
        type="button"
        className="cite cite-btn"
        aria-expanded={open}
        aria-controls={detailId}
        onClick={() => setOpen((v) => !v)}
        title={label}
      >
        <span className="cite-d" aria-hidden="true" />
        <span className="cite-label">{label}</span>
      </button>
      {open && (
        <div id={detailId} className="cite-detail" role="group" aria-label={`Source: ${label}`}>
          <div className="cite-detail-title">{source.title}</div>
          {source.locator && <div className="cite-detail-loc muted small">{source.locator}</div>}
          {source.excerpt && <blockquote className="cite-excerpt">{source.excerpt}</blockquote>}
          {href ? (
            <a className="cite-detail-link" href={href} target="_blank" rel="noreferrer noopener">
              {href}
            </a>
          ) : (
            <div className="muted small">{source.uri ?? 'No link available for this source.'}</div>
          )}
        </div>
      )}
    </span>
  );
}

function AnswerActions({ text }: { text: string }): JSX.Element {
  const [copyState, setCopyState] = useState('');
  const copy = async (): Promise<void> => {
    const { plain, html } = answerForClipboard(text);
    setCopyState((await copyText(plain, html)) ? 'Copied' : 'Select the answer to copy it.');
  };
  return (
    <div className="answer-actions" aria-label="Use this answer">
      <button type="button" onClick={() => void copy()}>
        <svg viewBox="0 0 24 24" aria-hidden="true">
          <rect x="9" y="9" width="11" height="11" rx="2" />
          <path d="M5 15V5a1 1 0 0 1 1-1h10" />
        </svg>
        Copy answer
      </button>
      {copyState && <span role="status">{copyState}</span>}
    </div>
  );
}

/**
 * A step that sends nothing visible for a while (planning, reading the document) must not look
 * frozen: three bouncing dots, beside the activity label when there is one.
 */
function WorkingIndicator({ activity }: { activity?: string }): JSX.Element {
  return (
    <span
      className="message-activity"
      role="status"
      aria-live="polite"
      aria-label={activity ? undefined : 'Working'}
    >
      <span className="typing-dots" aria-hidden="true">
        <i />
        <i />
        <i />
      </span>
      {activity ? <span>{activity}</span> : null}
    </span>
  );
}

/** Command structure a sentence rarely has: a sheet/cell reference, an `=` or a quoted value. */
const COMMAND_STRUCTURE = /["=!]|\b[A-Z]{1,3}\d+\b/;

/** Words that start a command or program line (`finish when=verified`, `let $x = …`). */
const COMMAND_WORDS = new Set<string>([...ALL_VERBS, 'finish', 'let']);

/**
 * A line the model wrote as a command rather than prose: lowercase (sentences start with a capital),
 * parsed by the real command grammar, and carrying command structure, so "Read the summary below." or
 * "Done! The chart is ready." stay visible. A bare `done` closes a narrated program.
 */
function isNarratedCommand(line: string, inFence = false): boolean {
  const t = line.trim();
  if (t === 'done') return true;
  // A command the grammar rejects is still a command (UAT 2026-10-16: `set I2 ""`,
  // `finish when=verified`, a bare `save` in a fence showed in an Excel reply).
  const word = /^\/?([a-z][\w-]*)/.exec(t)?.[1];
  if (word && COMMAND_WORDS.has(word) && (inFence || COMMAND_STRUCTURE.test(t))) return true;
  // A binding (`let $data = read 'Sheet'!A1:Q49`) is program syntax the line parser does not take.
  const binding = /^let \$\w+ = (.+)$/.exec(t);
  if (binding) return isNarratedCommand(binding[1] ?? '') || /^read\b/.test(binding[1] ?? '');
  // A slash form (`/add-shape slide=2 …`) is the same grammar (live 2026-10-13: shown in a reply).
  return (
    /^\/?[a-z]/.test(t) && COMMAND_STRUCTURE.test(t) && !isCommandParseError(parseCommandLine(t))
  );
}

/** Rejoin a command whose quoted argument runs onto the next line (`slide "T" "• a⏎• b"`). */
function joinQuotedLines(lines: string[]): string[] {
  const out: string[] = [];
  let open = '';
  for (const line of lines) {
    open = open ? `${open} ${line}` : line;
    if ((open.match(/"/g)?.length ?? 0) % 2 === 0) {
      out.push(open);
      open = '';
    }
  }
  if (open) out.push(open);
  return out;
}

/**
 * Remove commands from a reply. They are how the model drives the add-in, not part of the answer; the
 * approval card and Activity show what they do, and the reply says how many blocks were hidden.
 * Hidden: ```cmd / ```plan fences (as the runtime parses them), any other fence whose every line is a
 * command, and bare narrated command lines. An unclosed ```cmd / ```plan fence is hidden whole while
 * the reply is still streaming; in a finished reply its leading command lines are hidden and anything
 * after them stays visible, so prose after a stray fence is never lost.
 */
export function stripCommandBlocks(
  text: string,
  streaming = false,
): { text: string; hidden: number } {
  const lines = text.split('\n');
  const out: string[] = [];
  let hidden = 0;
  let inNarration = false;
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i] ?? '';
    const fence = /^[ \t]*```([\w-]*)[ \t]*$/.exec(line);
    if (fence) {
      let end = i + 1;
      while (end < lines.length && !/^[ \t]*```[ \t]*$/.test(lines[end] ?? '')) end++;
      const closed = end < lines.length;
      const body = lines.slice(i + 1, end);
      const label = (fence[1] ?? '').toLowerCase();
      const protocol = label === 'cmd' || label === 'plan';
      const allCommands =
        body.some((l) => l.trim()) &&
        joinQuotedLines(body).every((l) => !l.trim() || isNarratedCommand(l, true));
      if ((protocol && (closed || streaming)) || (closed && allCommands)) {
        hidden++;
        inNarration = false;
        i = end;
        continue;
      }
      if (protocol && !closed) {
        // Finished reply, fence never closed (live 2026-10-13: the command showed as a code block).
        let j = i + 1;
        while (
          j < lines.length &&
          (!(lines[j] ?? '').trim() || isNarratedCommand(lines[j] ?? '', true))
        )
          j++;
        if (j > i + 1) {
          hidden++;
          inNarration = false;
          i = j - 1;
          continue;
        }
      }
      out.push(...lines.slice(i, closed ? end + 1 : lines.length));
      inNarration = false;
      i = closed ? end : lines.length;
      continue;
    }
    if (isNarratedCommand(line)) {
      if (!inNarration) hidden++;
      inNarration = true;
      continue;
    }
    inNarration = false;
    out.push(line);
  }
  return {
    text: out
      .join('\n')
      .replace(/\n{3,}/g, '\n\n')
      .trim(),
    hidden,
  };
}

function Message({
  message,
  disabled,
  surface,
  onInsertArtifact,
  onRevealLocation,
  insertArtifactDisabledReason,
  language = 'en',
  onUndoChanges,
  onRetry,
}: {
  message: ChatMessage;
  disabled: boolean;
  surface: Surface;
  onInsertArtifact?: (artifact: InsertableArtifact) => void;
  onRevealLocation?: (location: string) => void;
  insertArtifactDisabledReason?: string;
  language?: UiLanguage;
  onUndoChanges?: (messageId: string) => void;
  onRetry?: () => void;
}): JSX.Element | null {
  const isUser = message.role === 'user';
  const stripped = isUser
    ? { text: message.text, hidden: 0 }
    : stripCommandBlocks(message.text, Boolean(message.streaming));
  const text = stripped.text;
  // Only a finished reply that ended up empty is dropped; a streaming one keeps its bubble.
  if (
    !isUser &&
    !text &&
    !message.streaming &&
    stripped.hidden === 0 &&
    !message.activity &&
    !message.error &&
    !message.cancelled &&
    !message.note &&
    !message.changes?.length &&
    !(message.sources && message.sources.length > 0)
  )
    return null;
  return (
    <div className={`m ${isUser ? 'u' : 'a'}`}>
      {!isUser && <span className="ic" aria-hidden="true" />}
      <div className="c">
        <div className="text">
          {isUser ? (
            <UserMessageContent text={message.text} />
          ) : (
            <MarkdownContent
              text={text}
              surface={surface}
              onInsertArtifact={
                message.streaming || message.cancelled || message.error
                  ? undefined
                  : onInsertArtifact
              }
              onRevealLocation={onRevealLocation}
              insertArtifactDisabledReason={insertArtifactDisabledReason}
            />
          )}
          {/* While waiting (a working step, or no text yet) show only the loader; the caret starts
              with the text. */}
          {message.streaming && (message.activity || !text) ? (
            <WorkingIndicator activity={message.activity} />
          ) : null}
          {/* Only when the bubble would otherwise be empty: next to an answer or a change card the
              note was technical noise (UAT 2026-10-15, "chat by Contoso-AI in the bubble"). */}
          {stripped.hidden > 0 && !text.trim() && !message.changes?.length && (
            <span className="cmd-hidden-note muted small">
              {stripped.hidden === 1
                ? '1 command block sent to the add-in. See Activity.'
                : `${stripped.hidden} command blocks sent to the add-in. See Activity.`}
            </span>
          )}
          {message.streaming && text && !message.activity && (
            <span className="caret" aria-label="streaming" />
          )}
        </div>
        {message.error && (
          <ErrorCard
            blocked={message.blocked === true}
            made={message.changes?.filter((c) => !c.undone).length ?? 0}
            language={language}
            onRetry={onRetry}
          />
        )}
        {message.changes?.length && !message.streaming ? (
          <ResultCard
            changes={message.changes}
            surface={surface}
            language={language}
            disabled={disabled}
            onShow={onRevealLocation}
            onUndo={onUndoChanges ? () => onUndoChanges(message.id) : undefined}
          />
        ) : null}
        {message.cancelled && <div className="muted small">Cancelled.</div>}
        {message.note && !message.streaming && (
          <div className="muted small" role="note">
            {message.note}
          </div>
        )}
        {message.sources && message.sources.length > 0 && (
          <div className="cites" aria-label="Citations">
            <span className="cites-h eyebrow">Sources</span>
            {message.sources.map((s, i) => (
              <Citation key={`${s.title}-${i}`} source={s} id={`${message.id}-${i}`} />
            ))}
          </div>
        )}
        {!isUser &&
          text &&
          !message.streaming &&
          !message.cancelled &&
          !message.error &&
          !message.clarifying && <AnswerActions text={text} />}
      </div>
    </div>
  );
}

const EMPTY_COPY: Record<Surface, string> = {
  word: "Ask about this document or selection. I'll ground answers on attached context and citations.",
  excel:
    'Ask about this workbook, sheet, or range. I can summarize, explain formulas, and stage gated changes.',
  powerpoint:
    'Ask about this deck or slide. I can summarize, critique, and draft reviewable slide content.',
  onenote: 'Ask about this page and its sources. I can synthesize grounded notes with citations.',
  outlook:
    'Ask about this message or thread. I can summarize, extract actions, and draft reviewable replies.',
  teams: 'Ask about this meeting or transcript. I can recap decisions and action items.',
};

/** The grounded conversation: user/assistant bubbles, streamed answer + citation pills. */
export function MessageThread({
  messages,
  showEmpty = true,
  disabled = false,
  surface = 'word',
  onInsertArtifact,
  onRevealLocation,
  insertArtifactDisabledReason,
  language,
  onUndoChanges,
  onRetry,
}: MessageThreadProps): JSX.Element {
  return (
    <div className="thread" role="log" aria-live="polite" aria-label="Conversation">
      {showEmpty && messages.length === 0 && (
        <div className="thread-empty">
          <div className="thread-empty-plate">
            <span className="thread-empty-mark" aria-hidden="true" />
            <p className="thread-empty-copy">{EMPTY_COPY[surface]}</p>
            <p className="thread-empty-next" aria-hidden="true">
              type <span className="te-key">/</span> for a verb, <span className="te-key">@</span>{' '}
              to ground — or attach a source above
            </p>
          </div>
        </div>
      )}
      {messages.map((m) => (
        <Message
          key={m.id}
          message={m}
          disabled={disabled}
          surface={surface}
          onInsertArtifact={onInsertArtifact}
          onRevealLocation={onRevealLocation}
          insertArtifactDisabledReason={insertArtifactDisabledReason}
          language={language}
          onUndoChanges={onUndoChanges}
          onRetry={m === messages[messages.length - 1] ? onRetry : undefined}
        />
      ))}
    </div>
  );
}

/**
 * What a request changed, in plain words, with Show me (reveal in the document) and Undo. Undo is
 * offered only where a safe restore exists; otherwise the card points to Office's own Undo.
 */
function ResultCard({
  changes,
  surface,
  language,
  disabled,
  onShow,
  onUndo,
}: {
  changes: ResultChange[];
  surface: Surface;
  language: UiLanguage;
  disabled: boolean;
  onShow?: (location: string) => void;
  onUndo?: () => void;
}): JSX.Element {
  const t = uiStrings(language);
  const allUndone = changes.every((c) => c.undone);
  const shown = changes.find(
    (c) => c.location && !c.undone && canRenderHostLocation(surface, c.location),
  );
  const undoable = changes.filter((c) => c.undoable && !c.undone);
  return (
    <div
      className={`result-card${allUndone ? ' is-undone' : ''}`}
      role="group"
      aria-label={allUndone ? t.undone : t.doneOne}
    >
      <div className="result-card-title">
        {allUndone ? t.undone : changes.length === 1 ? t.doneOne : t.doneMany(changes.length)}
      </div>
      <ul className="result-card-list">
        {changes.map((c) => (
          <li key={c.changeId} className={c.undone ? 'is-undone' : undefined}>
            {c.label}
          </li>
        ))}
      </ul>
      {!allUndone && (
        <div className="result-card-actions">
          {shown && onShow ? (
            <button
              type="button"
              className="btn sm"
              disabled={disabled}
              onClick={() => onShow(shown.location!)}
            >
              {t.showMe}
            </button>
          ) : null}
          {undoable.length > 0 && onUndo ? (
            <button type="button" className="btn sm" disabled={disabled} onClick={onUndo}>
              {changes.length > 1 ? t.undoAll : t.undo}
            </button>
          ) : (
            <span className="result-card-hint">{t.undoInOffice}</span>
          )}
        </div>
      )}
    </div>
  );
}

/**
 * A failed answer, in plain words and the conversation's language. The error itself is never shown
 * here: it is in Copy diagnostics (redacted).
 */
function ErrorCard({
  blocked,
  made,
  language,
  onRetry,
}: {
  blocked: boolean;
  made: number;
  language: UiLanguage;
  onRetry?: () => void;
}): JSX.Element {
  const t = uiStrings(language);
  return (
    <div className="error-card" role="alert">
      <div className="error-card-title">{blocked ? t.blockedTitle : t.failedTitle}</div>
      {blocked ? (
        <div className="error-card-body">{t.blockedBody}</div>
      ) : made > 0 ? (
        <div className="error-card-body">{t.failedPartial(made)}</div>
      ) : null}
      {!blocked && onRetry ? (
        <button type="button" className="btn sm pr error-card-retry" onClick={onRetry}>
          {t.tryAgain}
        </button>
      ) : null}
    </div>
  );
}
