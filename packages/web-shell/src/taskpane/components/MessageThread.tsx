import { useState } from 'react';
import type { ChatMessage } from '../../controller.js';
import type { SourceRef } from '@ge/contracts';

export interface MessageThreadProps {
  messages: ChatMessage[];
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
          {message.text}
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

/** The grounded conversation: user/assistant bubbles, streamed answer + citation pills. */
export function MessageThread({ messages }: MessageThreadProps): JSX.Element {
  return (
    <div className="thread" role="log" aria-live="polite" aria-label="Conversation">
      {messages.length === 0 && (
        <div className="m a">
          <span className="ic" aria-hidden="true" />
          <div className="c">
            <div className="text">
              Ask me anything about your document — I&apos;ll ground the answer on your research
              unit and attached context, with citations.
            </div>
          </div>
        </div>
      )}
      {messages.map((m) => (
        <Message key={m.id} message={m} />
      ))}
    </div>
  );
}
