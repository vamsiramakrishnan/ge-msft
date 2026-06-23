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

function Citation({ source }: { source: SourceRef }): JSX.Element {
  const label = source.locator ? `${source.title} · ${source.locator}` : source.title;
  const href = safeHttpUri(source.uri);
  const body = (
    <>
      <span className="cite-d" aria-hidden="true" />
      <span className="cite-label">{label}</span>
    </>
  );
  return href ? (
    <a
      className="cite"
      href={href}
      target="_blank"
      rel="noreferrer noopener"
      title={`${label} — opens in a new tab`}
    >
      {body}
    </a>
  ) : (
    <span className="cite" title={label}>
      {body}
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
              <Citation key={`${s.title}-${i}`} source={s} />
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
