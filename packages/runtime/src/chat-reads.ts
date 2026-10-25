/**
 * What a chat turn reads from the document before the model answers (docs/COMMAND-RELIABILITY.md,
 * fix E). Every bridge's `searchDocument` matches its query as one substring, so passing the whole
 * question ("Which region had the highest total revenue in A1:J11?") found nothing and the model
 * answered without data. These helpers turn a question into the probes a bridge can serve: a few
 * distinctive words for `searchDocument`, and the references it names for `readRange`. Selector
 * syntax stays the bridge's business: `readRange` returns `[]` for anything it cannot address.
 */

/** Maximum distinct words searched per turn (each is one host round-trip on the web). */
export const MAX_SEARCH_TERMS = 3;
/** Maximum references read per turn. */
export const MAX_REFERENCE_READS = 3;
/** Only the start of a very long question is probed, so a pasted blob stays cheap to scan. */
const MAX_PROBED_CHARS = 2000;

const STOP_WORDS = new Set(
  (
    'about above after again all also and any are because been before being below between both but ' +
    'can could did does doing down during each few for from further had has have having her here ' +
    'hers how into its itself just more most much must not now off once only other our out over own ' +
    'please same should some such than that the their them then there these they this those through ' +
    'too under until very was were what when where which while who whom why will with would your ' +
    'tell show give find list make says said exactly mention mentions document sheet slide slides ' +
    'deck page email mail message thread paragraph paragraphs range cell cells table'
  ).split(' '),
);

/** Up to {@link MAX_SEARCH_TERMS} distinctive words of the question, longest first. */
export function searchTerms(query: string): string[] {
  const seen = new Set<string>();
  const terms: string[] = [];
  for (const word of query.slice(0, MAX_PROBED_CHARS).match(/[\p{L}\p{N}][\p{L}\p{N}'-]*/gu) ??
    []) {
    const key = word.toLowerCase();
    if (key.length < 4 || STOP_WORDS.has(key) || /^\d+$/.test(key) || seen.has(key)) continue;
    seen.add(key);
    terms.push(word);
  }
  return terms.sort((a, b) => b.length - a.length).slice(0, MAX_SEARCH_TERMS);
}

/**
 * The references a question names: cell addresses (`A1:J11`, `Sheet2!B4`, `'Q3 Sales'!A1:C9`) and
 * slide numbers (`slide 2`). Up to {@link MAX_REFERENCE_READS}, in the order they appear.
 */
export function referenceCandidates(query: string): string[] {
  const found: string[] = [];
  const text = query.slice(0, MAX_PROBED_CHARS);
  // Sheet names are at most 31 characters, which also keeps the optional prefix linear.
  const address =
    /(?:(?:'[^']{1,31}'|[\p{L}\p{N}_.]{1,31})!)?\$?[A-Z]{1,3}\$?\d{1,7}(?::\$?[A-Z]{1,3}\$?\d{1,7})?(?![\p{L}\p{N}])/gu;
  for (const match of text.matchAll(address)) {
    const value = match[0];
    const start = match.index ?? 0;
    // A bare word such as "Q3" or "FY26" is not a reference unless it is sheet-qualified or a range.
    if (!value.includes('!') && !value.includes(':')) continue;
    if (start > 0 && /[\p{L}\p{N}]/u.test(text[start - 1]!)) continue;
    found.push(value);
  }
  for (const match of text.matchAll(/\bslide\s+(\d{1,3})\b/gi)) found.push(`slide ${match[1]}`);
  return [...new Set(found)].slice(0, MAX_REFERENCE_READS);
}
