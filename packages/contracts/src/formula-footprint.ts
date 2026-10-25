/** Shape inference only. Excel remains the evaluator; unfamiliar functions are staged/refused. */
const aggregate = new Set(
  'SUM AVERAGE MIN MAX COUNT COUNTA PRODUCT SUMPRODUCT MEDIAN STDEV STDEV.S STDEV.P VAR VAR.S VAR.P AND OR XOR CONCAT'.split(
    ' ',
  ),
);
const elementwise = new Set(
  'ABS ROUND ROUNDUP ROUNDDOWN INT TRUNC SIGN SQRT POWER MOD IF IFERROR IFNA IFS CHOOSE SWITCH TEXT VALUE NUMBERVALUE LEN LEFT RIGHT MID LOWER UPPER PROPER TRIM CLEAN SUBSTITUTE REPLACE FIND SEARCH DATE TIME YEAR MONTH DAY HOUR MINUTE SECOND TODAY NOW DAYS NETWORKDAYS WORKDAY EDATE EOMONTH ISBLANK ISNUMBER ISTEXT ISERROR ISNA NOT CONCATENATE EXP LN LOG LOG10 SIN COS TAN ASIN ACOS ATAN PI RAND RANDBETWEEN TRUE FALSE ROW COLUMN'.split(
    ' ',
  ),
);
const arrayFunctions = new Set(
  'FILTER SORT SORTBY UNIQUE SEQUENCE RANDARRAY TRANSPOSE CHOOSECOLS CHOOSEROWS TAKE DROP EXPAND HSTACK VSTACK TOCOL TOROW WRAPCOLS WRAPROWS INDIRECT OFFSET'.split(
    ' ',
  ),
);
type Frame = { name?: string; args: boolean[]; array: boolean; start: number; raw: string[] };
const literalRange =
  /^(?:(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(?:\$?[A-Za-z]{1,3}\$?[1-9]\d*(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[1-9]\d*:\$?[1-9]\d*)$/;
function arrayResult(frame: Frame): boolean {
  const args = [...frame.args, frame.array];
  const name = frame.name;
  if (!name) return args.some(Boolean);
  if (aggregate.has(name)) return false;
  if (arrayFunctions.has(name)) return true;
  const literal = (index: number): boolean => literalRange.test(frame.raw[index] ?? '');
  if (name === 'COUNTBLANK' || name === 'ROWS' || name === 'COLUMNS') return !literal(0);
  if (name === 'SUMIF' || name === 'COUNTIF' || name === 'AVERAGEIF')
    return !literal(0) || (frame.raw.length > 2 && !literal(2)) || (args[1] ?? false);
  if (name === 'SUMIFS' || name === 'AVERAGEIFS' || name === 'MINIFS' || name === 'MAXIFS')
    return (
      !literal(0) ||
      frame.raw.some((_value, index) => index % 2 === 1 && !literal(index)) ||
      args.some((value, index) => index >= 2 && index % 2 === 0 && value)
    );
  if (name === 'COUNTIFS')
    return (
      frame.raw.some((_value, index) => index % 2 === 0 && !literal(index)) ||
      args.some((value, index) => index % 2 === 1 && value)
    );
  if (name === 'MATCH' || name === 'XMATCH')
    return !literal(1) || args.some((value, index) => index !== 1 && value);
  if (name === 'LOOKUP')
    return !literal(1) || (frame.raw.length > 2 && !literal(2)) || (args[0] ?? false);
  if (name === 'SUBTOTAL')
    return frame.raw.some((_value, index) => index > 0 && !literal(index)) || (args[0] ?? false);
  if (name === 'VLOOKUP' || name === 'HLOOKUP')
    return !literal(1) || args.some((value, index) => index !== 1 && value);
  if (name === 'TEXTJOIN') return Boolean(args[0] || args[1]);
  return args.some(Boolean);
}
const specialized = new Set(
  'SUMIF COUNTIF AVERAGEIF SUMIFS AVERAGEIFS MINIFS MAXIFS COUNTIFS MATCH XMATCH LOOKUP SUBTOTAL VLOOKUP HLOOKUP TEXTJOIN COUNTBLANK ROWS COLUMNS'.split(
    ' ',
  ),
);

/** Proves a one-cell result for a conservative, extensible set of native expression shapes. */
export function hasScalarFormulaFootprint(formula: string): boolean {
  if (!formula.startsWith('=') || formula.length > 32768) return false;
  const stack: Frame[] = [{ args: [], array: false, start: 1, raw: [] }];
  let index = 1;
  let tokens = 0;
  const current = (): Frame => stack[stack.length - 1]!;
  while (index < formula.length && ++tokens < 4096) {
    const char = formula[index]!;
    if (/\s/.test(char)) {
      index++;
      continue;
    }
    if (char === '"' || char === "'") {
      const quote = char;
      let ended = false;
      for (index++; index < formula.length; index++) {
        if (formula[index] !== quote) continue;
        if (formula[index + 1] === quote) {
          index++;
          continue;
        }
        index++;
        ended = true;
        break;
      }
      if (!ended) return false;
      if (quote === "'") {
        if (formula[index] !== '!') return false;
        index++;
      }
      continue;
    }
    if (char === '{' || char === '[') {
      const close = char === '{' ? '}' : ']';
      let depth = 1;
      for (index++; index < formula.length && depth; index++) {
        if (formula[index] === char) depth++;
        if (formula[index] === close) depth--;
      }
      if (depth) return false;
      current().array = true;
      continue;
    }
    const number = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/.exec(formula.slice(index));
    if (number) {
      index += number[0].length;
      continue;
    }
    const identifier = /^\$?[A-Za-z_][A-Za-z0-9_.$]*/.exec(formula.slice(index));
    if (identifier) {
      const raw = identifier[0];
      index += raw.length;
      if (formula[index] === '!') {
        index++;
        continue;
      }
      while (/\s/.test(formula[index] ?? '') && index < formula.length) index++;
      if (formula[index] === '(') {
        const name = raw
          .toUpperCase()
          .replace(/^_XLFN\./, '')
          .replace(/^_XLWS\./, '');
        if (
          !aggregate.has(name) &&
          !elementwise.has(name) &&
          !arrayFunctions.has(name) &&
          !specialized.has(name)
        )
          return false;
        stack.push({ name, args: [], array: false, start: index + 1, raw: [] });
        index++;
        if (stack.length > 64) return false;
      } else if (!/^\$?[A-Za-z]{1,3}\$?[1-9]\d*$/.test(raw) && !/^(TRUE|FALSE)$/i.test(raw))
        current().array = true;
      continue;
    }
    if (char === '(') {
      stack.push({ args: [], array: false, start: index + 1, raw: [] });
      if (stack.length > 64) return false;
    } else if (char === ')') {
      if (stack.length < 2) return false;
      const frame = stack.pop()!;
      frame.raw.push(formula.slice(frame.start, index).trim());
      current().array ||= arrayResult(frame);
    } else if (char === ',') {
      if (current().name) {
        current().args.push(current().array);
        current().array = false;
        current().raw.push(formula.slice(current().start, index).trim());
        current().start = index + 1;
      } else current().array = true;
    } else if (char === ':' || char === '#') current().array = true;
    else if (!'+-*/^&=<>%'.includes(char)) return false;
    index++;
  }
  return tokens < 4096 && stack.length === 1 && !arrayResult(stack[0]!);
}
