import { useState } from 'react';
import type { ReactNode } from 'react';
import type { ChatMessage } from '../../controller.js';
import type { SourceRef, Surface } from '@ge/contracts';

export interface MessageThreadProps {
  messages: ChatMessage[];
  surface?: Surface;
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
  const markers = ['**', '`', '[', '*'];
  const indexes = markers.map((marker) => text.indexOf(marker, start)).filter((idx) => idx >= 0);
  return indexes.length ? Math.min(...indexes) : text.length;
}

function renderInline(text: string, keyPrefix: string): ReactNode[] {
  const nodes: ReactNode[] = [];
  let i = 0;
  let key = 0;
  const pushText = (value: string): void => {
    if (value) nodes.push(value);
  };

  while (i < text.length) {
    if (text.startsWith('**', i)) {
      const end = text.indexOf('**', i + 2);
      if (end > i + 2) {
        nodes.push(
          <strong key={`${keyPrefix}-b-${key++}`}>
            {renderInline(text.slice(i + 2, end), `${keyPrefix}-b-${key}`)}
          </strong>,
        );
        i = end + 2;
        continue;
      }
    }

    if (text[i] === '`') {
      const end = text.indexOf('`', i + 1);
      if (end > i + 1) {
        nodes.push(<code key={`${keyPrefix}-c-${key++}`}>{text.slice(i + 1, end)}</code>);
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
        const href = safeHttpUri(text.slice(urlStart, urlEnd).trim());
        nodes.push(
          href ? (
            <a
              key={`${keyPrefix}-a-${key++}`}
              className="md-link"
              href={href}
              target="_blank"
              rel="noreferrer noopener"
            >
              {renderInline(label, `${keyPrefix}-a-${key}`)}
            </a>
          ) : (
            <span key={`${keyPrefix}-bad-a-${key++}`}>
              {label} ({text.slice(urlStart, urlEnd).trim()})
            </span>
          ),
        );
        i = urlEnd + 1;
        continue;
      }
    }

    if (text[i] === '*' && !text.startsWith('**', i)) {
      const end = text.indexOf('*', i + 1);
      if (end > i + 1) {
        nodes.push(
          <em key={`${keyPrefix}-i-${key++}`}>
            {renderInline(text.slice(i + 1, end), `${keyPrefix}-i-${key}`)}
          </em>,
        );
        i = end + 1;
        continue;
      }
    }

    const next = findNextInlineMarker(text, i + 1);
    pushText(text.slice(i, next));
    i = next;
  }

  return nodes;
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

function UserMessageContent({ text }: { text: string }): JSX.Element {
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

function MarkdownContent({ text }: { text: string }): JSX.Element {
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
      blocks.push(
        <pre key={`code-${block++}`} className="md-code">
          <code>{code.join('\n')}</code>
        </pre>,
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
      blocks.push(
        <div key={`table-wrap-${block++}`} className="md-table-wrap">
          <table className="md-table">
            <thead>
              <tr>
                {headers.map((cell, idx) => (
                  <th key={idx}>{renderInline(cell, `th-${block}-${idx}`)}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.map((row, rowIdx) => (
                <tr key={rowIdx}>
                  {row.map((cell, cellIdx) => (
                    <td key={cellIdx}>{renderInline(cell, `td-${block}-${rowIdx}-${cellIdx}`)}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
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
      blocks.push(<Tag key={`h-${block++}`}>{renderInline(content, `h-${block}`)}</Tag>);
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
            <li key={idx}>{renderInline(item, `li-${block}-${idx}`)}</li>
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
      blocks.push(<p key={`p-${block++}`}>{renderInline(paragraph.join(' '), `p-${block}`)}</p>);
      continue;
    }

    blocks.push(<p key={`fallback-${block++}`}>{renderInline(trimmed, `fallback-${block}`)}</p>);
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

function Message({ message }: { message: ChatMessage }): JSX.Element {
  const isUser = message.role === 'user';
  return (
    <div className={`m ${isUser ? 'u' : 'a'}`}>
      {!isUser && <span className="ic" aria-hidden="true" />}
      <div className="c">
        <div className="text">
          {isUser ? (
            <UserMessageContent text={message.text} />
          ) : (
            <MarkdownContent text={message.text} />
          )}
          {message.streaming && <span className="caret" aria-label="streaming" />}
        </div>
        {message.error && (
          <div className="msg-error" role="alert">
            {message.error}
          </div>
        )}
        {message.cancelled && <div className="muted small">Cancelled.</div>}
        {message.sources && message.sources.length > 0 && (
          <div className="cites" aria-label="Citations">
            <span className="cites-h eyebrow">Sources</span>
            {message.sources.map((s, i) => (
              <Citation key={`${s.title}-${i}`} source={s} id={`${message.id}-${i}`} />
            ))}
          </div>
        )}
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
export function MessageThread({ messages, surface = 'word' }: MessageThreadProps): JSX.Element {
  return (
    <div className="thread" role="log" aria-live="polite" aria-label="Conversation">
      {messages.length === 0 && (
        <div className="m a">
          <span className="ic" aria-hidden="true" />
          <div className="c">
            <div className="text">{EMPTY_COPY[surface]}</div>
          </div>
        </div>
      )}
      {messages.map((m) => (
        <Message key={m.id} message={m} />
      ))}
    </div>
  );
}
