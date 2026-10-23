import type { ActuationRequest } from '@ge/contracts';

/**
 * Render a compiled `ActuationRequest` as the single, verbatim command line the approval card
 * shows the user (ADR-0004 §3.1: "the approval card renders the command verbatim"). This is the
 * inverse of the runtime's parse/compile step — it reconstructs the flat `cmd`-grammar line the
 * model emitted, so the human approves exactly what will actuate.
 *
 * Pure and total: unknown / partially-specified kinds degrade to a best-effort label rather than
 * throwing, because a missing label must never block (or crash) the fail-closed approval path.
 */
export function renderCommandLine(request: ActuationRequest): string {
  const p = request.params;
  switch (request.kind) {
    case 'write-cells': {
      const range = p.target?.range ?? '<range>';
      const cells = p.cells ?? [];
      if (cells.length > 1 || (cells[0]?.length ?? 0) > 1) {
        return `grid ${range} (${cells.length}x${cells[0]?.length ?? 0})`;
      }
      const value = firstCell(cells) ?? p.text ?? '';
      return `set ${range} ${value}`.trimEnd();
    }
    case 'tracked-change': {
      // `suggest "<old>" => "<new>"` — content-anchored Word edit.
      const oldText = p.target?.matchText ?? '';
      const newText = p.text ?? p.ooxml ?? p.html ?? '';
      return `suggest ${quote(oldText)} => ${quote(newText)}`;
    }
    case 'add-comment':
    case 'comment-reply': {
      // `comment <selector> "text"`.
      const selector = p.target?.matchText ?? p.target?.range ?? p.target?.commentId ?? '<sel>';
      return `comment ${selector} ${quote(p.text ?? '')}`;
    }
    case 'format-cells': {
      // `format <range> k=v ...`.
      const range = p.target?.range ?? '<range>';
      const pairs = formatPairs(p.format);
      return `format ${range}${pairs ? ` ${pairs}` : ''}`;
    }
    case 'insert-slide': {
      // `slide "Title" "bullet" … [at=N]` — the composed slide, so the card shows what it will say.
      // A request carrying a deck inserts that deck (the bridge prefers it), so never label it a slide.
      if (!p.slide || p.deck?.base64 || p.ooxml) return labelFallback(request);
      const at = p.target?.slideIndex;
      return [
        'slide',
        quote(p.slide.title),
        ...p.slide.bullets.map(quote),
        ...(typeof at === 'number' ? [`at=${at + 1}`] : []),
      ].join(' ');
    }
    case 'reply-mail': {
      // The file and the full text it holds: a reply leaves the mailbox, so nothing is hidden.
      const file = p.attachment;
      return [
        'mail',
        quote(p.mail?.body ?? p.text ?? ''),
        ...(file ? [`attach=${quote(file.name ?? 'a file')}`] : []),
        ...(file?.text !== undefined ? [`text=${quote(file.text)}`] : []),
      ].join(' ');
    }
    case 'add-attachment': {
      const file = p.attachment;
      return [
        '/add-attachment',
        `name=${quote(file?.name ?? file?.uri ?? 'a file')}`,
        ...(file?.text !== undefined ? [`text=${quote(file.text)}`] : []),
      ].join(' ');
    }
    case 'sort-range': {
      // The pinned range, key and order: the user approves exactly which cells move (review H3).
      const sort = p.sortRange;
      if (!sort) return labelFallback(request);
      return [
        '/sort-range',
        `range=${sort.range}`,
        `key=${quote(sort.key)}`,
        `order=${sort.order ?? 'ascending'}`,
        `headers=${sort.hasHeaders ?? true}`,
      ].join(' ');
    }
    case 'insert-table': {
      // Size and first row; and, in capitals, that it REPLACES paragraphs (review H1).
      const grid = p.tableGrid;
      if (!grid) return labelFallback(request);
      const size = `${grid.rows.length}x${grid.rows[0]?.length ?? 0} table`;
      const first = `(first row: ${(grid.rows[0] ?? []).join(' | ')})`;
      const anchor = p.target?.matchText;
      if (grid.replaceThrough && anchor)
        return `/insert-table ${size} ${first} REPLACES the paragraphs from ${quote(anchor)} through ${quote(grid.replaceThrough)}`;
      return `/insert-table ${size} ${first}${anchor ? ` after ${quote(anchor)}` : ''}`;
    }
    default:
      // Best-effort fallback for any kind without a dedicated renderer (e.g. compose-deck).
      return labelFallback(request);
  }
}

function firstCell(cells: readonly (readonly string[])[] | undefined): string | undefined {
  return cells?.[0]?.[0];
}

function formatPairs(format: ActuationRequest['params']['format']): string {
  if (!format) return '';
  const parts: string[] = [];
  if (format.bold !== undefined) parts.push(`bold=${format.bold}`);
  if (format.italic !== undefined) parts.push(`italic=${format.italic}`);
  if (format.fill !== undefined) parts.push(`fill=${format.fill}`);
  if (format.numberFormat !== undefined) parts.push(`numberFormat=${quote(format.numberFormat)}`);
  if (format.fontColor !== undefined) parts.push(`fontColor=${format.fontColor}`);
  if (format.fontSize !== undefined) parts.push(`fontSize=${format.fontSize}`);
  if (format.align !== undefined) parts.push(`align=${format.align}`);
  if (format.valign !== undefined) parts.push(`valign=${format.valign}`);
  if (format.wrap !== undefined) parts.push(`wrap=${format.wrap}`);
  if (format.border !== undefined) parts.push(`border=${format.border}`);
  if (format.borderOutside !== undefined) parts.push(`borderOutside=${format.borderOutside}`);
  if (format.borderInside !== undefined) parts.push(`borderInside=${format.borderInside}`);
  if (format.borderColor !== undefined) parts.push(`borderColor=${format.borderColor}`);
  if (format.autofit) parts.push('autofit=true');
  return parts.join(' ');
}

/** A double-quoted, escape-safe rendering of a value (matches the grammar's quoted operands). */
function quote(value: string): string {
  return `"${value.replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`;
}

function labelFallback(request: ActuationRequest): string {
  const p = request.params;
  const hint = p.text ?? p.target?.matchText ?? p.target?.range ?? '';
  return hint ? `${request.kind} ${quote(hint)}` : request.kind;
}
