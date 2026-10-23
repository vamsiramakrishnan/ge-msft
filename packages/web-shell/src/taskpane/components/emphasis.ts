/**
 * Markdown emphasis (`*`, `_`) by the CommonMark delimiter rules, so nested and triple markers
 * render instead of showing as stray asterisks (real Outlook, 2026-10-17: "*Gemini Enterprise >
 * … data store*" and "(Active*)" in an answer). Pure: the caller turns the pieces into React.
 *
 * A delimiter run opens when it is left-flanking (a word follows) and closes when it is
 * right-flanking (a word precedes); `_` also may not open or close inside a word (`snake_case`).
 * Runs pair with the nearest earlier opener of the same character, two characters for strong and
 * one for em, and unmatched characters stay literal, as in CommonMark.
 */

/** A piece the caller renders: plain text, an atom it made earlier (code, link), or emphasis. */
export type EmphasisNode<A> =
  | { kind: 'text'; text: string }
  | { kind: 'atom'; atom: A }
  | { kind: 'em' | 'strong'; children: EmphasisNode<A>[] };

/** Input: text with atoms (code spans, links) already cut out in order. */
export type EmphasisInput<A> = Array<{ kind: 'text'; text: string } | { kind: 'atom'; atom: A }>;

interface Delimiter {
  char: '*' | '_';
  count: number;
  readonly original: number;
  canOpen: boolean;
  canClose: boolean;
  active: boolean;
  /** Matched openings, in match order (the first match is innermost). */
  opens: Array<'em' | 'strong'>;
  /** Matched closings, in match order (the first match is innermost). */
  closes: Array<'em' | 'strong'>;
}

type Token<A> =
  | { kind: 'text'; text: string }
  | { kind: 'atom'; atom: A }
  | { kind: 'delim'; delim: Delimiter };

const WHITESPACE = /\s/u;
const PUNCTUATION = /[\p{P}\p{S}]/u;

function flanking(before: string, after: string): { left: boolean; right: boolean } {
  const beforeSpace = before === '' || WHITESPACE.test(before);
  const afterSpace = after === '' || WHITESPACE.test(after);
  const beforePunct = PUNCTUATION.test(before);
  const afterPunct = PUNCTUATION.test(after);
  return {
    left: !afterSpace && (!afterPunct || beforeSpace || beforePunct),
    right: !beforeSpace && (!beforePunct || afterSpace || afterPunct),
  };
}

/** Split text pieces into text and delimiter runs; atoms count as word characters around them. */
function tokenize<A>(input: EmphasisInput<A>): Token<A>[] {
  const tokens: Token<A>[] = [];
  const charAt = (piece: number, index: number): string => {
    // The character next to a run, looking across piece boundaries (an atom reads as a letter).
    let p = piece;
    let i = index;
    while (p >= 0 && p < input.length) {
      const item = input[p]!;
      if (item.kind === 'atom') return 'a';
      if (i >= 0 && i < item.text.length) return item.text[i]!;
      if (i < 0) {
        p -= 1;
        const prev = input[p];
        i = prev && prev.kind === 'text' ? prev.text.length - 1 : 0;
      } else {
        p += 1;
        i = 0;
      }
    }
    return '';
  };
  input.forEach((item, piece) => {
    if (item.kind === 'atom') {
      tokens.push(item);
      return;
    }
    const text = item.text;
    let start = 0;
    let i = 0;
    while (i < text.length) {
      const c = text[i];
      if (c !== '*' && c !== '_') {
        i += 1;
        continue;
      }
      let end = i;
      while (end < text.length && text[end] === c) end += 1;
      if (i > start) tokens.push({ kind: 'text', text: text.slice(start, i) });
      const before = charAt(piece, i - 1);
      const after = charAt(piece, end);
      const { left, right } = flanking(before, after);
      const count = end - i;
      const underscore = c === '_';
      tokens.push({
        kind: 'delim',
        delim: {
          char: c as '*' | '_',
          count,
          original: count,
          canOpen: underscore ? left && (!right || PUNCTUATION.test(before)) : left,
          canClose: underscore ? right && (!left || PUNCTUATION.test(after)) : right,
          active: true,
          opens: [],
          closes: [],
        },
      });
      start = end;
      i = end;
    }
    if (start < text.length) tokens.push({ kind: 'text', text: text.slice(start) });
  });
  return tokens;
}

/** CommonMark "process emphasis": pair closers with the nearest usable opener. */
function pair<A>(tokens: Token<A>[]): void {
  const delims = tokens.flatMap((t) => (t.kind === 'delim' ? [t.delim] : []));
  for (let c = 0; c < delims.length; c++) {
    const closer = delims[c]!;
    if (!closer.canClose) continue;
    let o = c - 1;
    while (closer.count > 0 && o >= 0) {
      const opener = delims[o]!;
      const usable =
        opener.active &&
        opener.canOpen &&
        opener.count > 0 &&
        opener.char === closer.char &&
        // The "rule of 3": a run that can both open and close does not pair across a multiple of 3.
        !(
          (opener.canClose || closer.canOpen) &&
          (opener.original + closer.original) % 3 === 0 &&
          !(opener.original % 3 === 0 && closer.original % 3 === 0)
        );
      if (!usable) {
        o -= 1;
        continue;
      }
      const use = opener.count >= 2 && closer.count >= 2 ? 2 : 1;
      const kind = use === 2 ? 'strong' : 'em';
      opener.count -= use;
      closer.count -= use;
      opener.opens.push(kind);
      closer.closes.push(kind);
      // Delimiters between the pair can no longer open or close anything.
      for (let between = o + 1; between < c; between++) delims[between]!.active = false;
      if (opener.count === 0) o -= 1;
    }
  }
}

/** Turn text with atoms into nested emphasis; malformed markup falls back to plain text. */
export function parseEmphasis<A>(input: EmphasisInput<A>): EmphasisNode<A>[] {
  const tokens = tokenize(input);
  pair(tokens);
  type Frame = { kind: 'em' | 'strong' | 'root'; children: EmphasisNode<A>[] };
  const root: Frame = { kind: 'root', children: [] };
  const stack: Frame[] = [root];
  const top = (): Frame => stack[stack.length - 1]!;
  const pushText = (text: string): void => {
    if (!text) return;
    const children = top().children;
    const last = children[children.length - 1];
    if (last?.kind === 'text') last.text += text;
    else children.push({ kind: 'text', text });
  };
  for (const token of tokens) {
    if (token.kind === 'text') pushText(token.text);
    else if (token.kind === 'atom') top().children.push(token);
    else {
      const d = token.delim;
      // Closings first (innermost first), then unmatched characters, then openings (outermost first).
      for (const kind of d.closes) {
        const frame = stack.pop();
        if (!frame || frame.kind !== kind) return input.map((p) => ({ ...p }));
        top().children.push({ kind, children: frame.children });
      }
      pushText(d.char.repeat(d.count));
      for (const kind of [...d.opens].reverse()) stack.push({ kind, children: [] });
    }
  }
  return stack.length === 1 ? root.children : input.map((p) => ({ ...p }));
}
