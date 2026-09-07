/** A deliberately restricted analytical SQL dialect layered over a read-only, external-I/O-disabled engine.
 * Tokens are parsed before dispatch; strings cannot hide statement separators or disallowed calls.
 * Engine restrictions remain authoritative even when a query is syntactically a SELECT.
 */
const FORBIDDEN = new Set(
  'attach detach copy export import install load pragma set reset create alter drop insert update delete merge call execute prepare deallocate vacuum checkpoint force use grant revoke recursive into secret secrets read_csv read_parquet read_json glob query query_table'.split(
    ' ',
  ),
);
const FUNCTIONS = new Set(
  'abs avg count sum min max round floor ceil ceiling coalesce nullif if ifnull lower upper trim ltrim rtrim length char_length substring substr replace regexp_replace regexp_full_match concat concat_ws greatest least cast try_cast decimal numeric varchar integer bigint double boolean date timestamp extract date_part date_trunc strftime row_number rank dense_rank lag lead first_value last_value ntile percent_rank cume_dist over filter in as'.split(
    ' ',
  ),
);
export function validateQuery(sql: string): string {
  if (!sql.trim() || sql.length > 32_768) throw new Error('Query must contain 1–32768 characters.');
  const tokens: Array<{ text: string; kind: 'word' | 'literal' | 'symbol' }> = [];
  let i = 0;
  while (i < sql.length) {
    const c = sql[i]!;
    if (/\s/.test(c)) {
      i++;
      continue;
    }
    if (c === "'") {
      let closed = false;
      const start = i++;
      while (i < sql.length) {
        if (sql[i++] === "'") {
          if (sql[i] === "'") {
            i++;
            continue;
          }
          closed = true;
          break;
        }
      }
      if (!closed) throw new Error('Unclosed SQL string.');
      tokens.push({ text: sql.slice(start, i), kind: 'literal' });
      continue;
    }
    if (/[A-Za-z_]/.test(c)) {
      const start = i++;
      while (i < sql.length && /[A-Za-z0-9_]/.test(sql[i]!)) i++;
      tokens.push({ text: sql.slice(start, i).toLowerCase(), kind: 'word' });
      continue;
    }
    if (/[0-9]/.test(c)) {
      const start = i++;
      while (i < sql.length && /[0-9.]/.test(sql[i]!)) i++;
      tokens.push({ text: sql.slice(start, i), kind: 'literal' });
      continue;
    }
    if (
      !'(),.+-*/%=<>!|'.includes(c) ||
      sql.slice(i, i + 2) === '--' ||
      sql.slice(i, i + 2) === '/*'
    )
      throw new Error(
        'Only a single analytical SELECT is supported; comments, quoted identifiers and statement separators are unavailable.',
      );
    tokens.push({ text: c, kind: 'symbol' });
    i++;
  }
  if (!['select', 'with'].includes(tokens[0]?.text ?? ''))
    throw new Error('Only SELECT or non-recursive WITH queries are supported.');
  for (let j = 0; j < tokens.length; j++) {
    const token = tokens[j]!;
    if (token.kind !== 'word') continue;
    if (FORBIDDEN.has(token.text))
      throw new Error(`SQL operation ${token.text} is unavailable in this workspace.`);
    if (
      tokens[j + 1]?.text === '(' &&
      !FUNCTIONS.has(token.text) &&
      !['select', 'from', 'join', 'where', 'and', 'or', 'not', 'exists'].includes(token.text)
    ) {
      // CTE alias AS (...) is permitted; table functions and arbitrary code-bearing functions are not.
      throw new Error(`SQL function ${token.text} is not in the analytical allowlist.`);
    }
  }
  return sql.trim();
}
