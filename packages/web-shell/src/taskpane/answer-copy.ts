/**
 * Copy answer puts the answer on the clipboard twice: as HTML, so Word and Outlook paste headings,
 * bold, lists and tables, and as plain text without markdown symbols (`##`, `**`, `|---|`) for
 * everything else. Answer text is model output: every piece is HTML-escaped and links are kept only
 * for http(s).
 */

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

const LINK = /\[([^\]]+)\]\(([^)\s]+)\)/g;

function inlineHtml(text: string): string {
  return escapeHtml(text)
    .replace(/`([^`]+)`/g, '<code>$1</code>')
    .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
    .replace(/(^|[^*])\*([^*\s][^*]*)\*/g, '$1<em>$2</em>')
    .replace(LINK, (_m, label: string, url: string) =>
      /^https?:\/\//i.test(url) ? `<a href="${url}">${label}</a>` : label,
    );
}

function inlinePlain(text: string): string {
  return text
    .replace(/`([^`]+)`/g, '$1')
    .replace(/\*\*([^*]+)\*\*/g, '$1')
    .replace(/(^|[^*])\*([^*\s][^*]*)\*/g, '$1$2')
    .replace(LINK, (_m, label: string, url: string) =>
      /^https?:\/\//i.test(url) ? `${label} (${url})` : label,
    );
}

const cells = (line: string): string[] =>
  line
    .trim()
    .replace(/^\|/, '')
    .replace(/\|$/, '')
    .split('|')
    .map((c) => c.trim());
const isSeparator = (line: string): boolean =>
  cells(line).length > 1 && cells(line).every((c) => /^:?-{3,}:?$/.test(c));

/** The answer as HTML for pasting into Office, and as plain text. */
export function answerForClipboard(markdown: string): { html: string; plain: string } {
  const lines = markdown.replace(/\r\n/g, '\n').split('\n');
  const html: string[] = [];
  const plain: string[] = [];
  let list: 'ul' | 'ol' | undefined;
  const closeList = (): void => {
    if (list) html.push(`</${list}>`);
    list = undefined;
  };

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i] ?? '';
    const t = line.trim();
    if (t.startsWith('```')) {
      closeList();
      const code: string[] = [];
      while (++i < lines.length && !(lines[i] ?? '').trim().startsWith('```')) code.push(lines[i]!);
      html.push(`<pre>${escapeHtml(code.join('\n'))}</pre>`);
      plain.push(...code);
      continue;
    }
    if (t.includes('|') && isSeparator(lines[i + 1] ?? '')) {
      closeList();
      const head = cells(t);
      const rows: string[][] = [];
      i += 2;
      while (i < lines.length && (lines[i] ?? '').includes('|')) rows.push(cells(lines[i++]!));
      i--;
      html.push(
        '<table border="1" style="border-collapse:collapse">' +
          `<tr>${head.map((c) => `<th>${inlineHtml(c)}</th>`).join('')}</tr>` +
          rows
            .map((r) => `<tr>${r.map((c) => `<td>${inlineHtml(c)}</td>`).join('')}</tr>`)
            .join('') +
          '</table>',
      );
      plain.push(
        head.map(inlinePlain).join('\t'),
        ...rows.map((r) => r.map(inlinePlain).join('\t')),
      );
      continue;
    }
    const heading = /^(#{1,4})\s+(.*)$/.exec(t);
    const bullet = /^[-*+]\s+(.*)$/.exec(t);
    const numbered = /^(\d+)\.\s+(.*)$/.exec(t);
    if (bullet || numbered) {
      const kind = bullet ? 'ul' : 'ol';
      if (list !== kind) {
        closeList();
        html.push(`<${kind}>`);
        list = kind;
      }
      const item = bullet ? bullet[1]! : numbered![2]!;
      html.push(`<li>${inlineHtml(item)}</li>`);
      plain.push(bullet ? `• ${inlinePlain(item)}` : `${numbered![1]}. ${inlinePlain(item)}`);
      continue;
    }
    closeList();
    if (!t) {
      if (plain.length && plain[plain.length - 1] !== '') plain.push('');
      continue;
    }
    if (/^([-*_])\1\1+$/.test(t)) {
      html.push('<hr>');
      continue;
    }
    if (heading) {
      const level = Math.min(heading[1]!.length + 1, 4);
      html.push(`<h${level}>${inlineHtml(heading[2]!)}</h${level}>`);
      plain.push(inlinePlain(heading[2]!));
      continue;
    }
    html.push(`<p>${inlineHtml(t)}</p>`);
    plain.push(inlinePlain(t));
  }
  closeList();
  return { html: html.join(''), plain: plain.join('\n').trim() };
}
