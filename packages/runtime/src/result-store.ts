/** Task-scoped, memory-only tool results. References identify data, never authority. */
export interface CommandResultStoreOptions {
  turnBytes?: number;
  inlineBytes?: number;
  itemBytes?: number;
  totalBytes?: number;
  maxItems?: number;
  maxDepth?: number;
  maxNodes?: number;
  maxResults?: number;
}

export interface EncodedCommandResults {
  text: string;
  /** Serialized input bytes observed; a lower bound when inputBytesComplete is false. */
  inputBytes: number;
  inputBytesComplete: boolean;
  outputBytes: number;
  retained: number;
  errors: number;
}

type Json = null | boolean | number | string | Json[] | { [key: string]: Json };
type RecordValue = { value: Json; bytes: number; ref: string };
type StringCursor = { length: number; offset: number; position: number };
type Limits = Required<CommandResultStoreOptions>;
const encoder = new TextEncoder();
const forbidden = new Set(['__proto__', 'prototype', 'constructor']);
let nextScope = 0;

class ResultDataError extends Error {
  constructor(
    readonly code: string,
    message: string,
  ) {
    super(message);
  }
}

const byteLength = (text: string): number => encoder.encode(text).byteLength;
const typeOf = (value: Json): string =>
  value === null ? 'null' : Array.isArray(value) ? 'array' : typeof value;

/** Path hints disclose structure only; values require deliberate inspection. */
function shapeOf(value: Json, stringLength: (value: string) => number): { [key: string]: Json } {
  if (Array.isArray(value)) return { length: value.length };
  if (typeof value === 'string') return { length: stringLength(value) };
  if (value && typeof value === 'object') {
    const keys: string[] = [];
    let keyCount = 0;
    let keyBytes = 0;
    for (const key in value) {
      if (!Object.hasOwn(value, key)) continue;
      keyCount++;
      // Omitted long keys are explicit. Do not truncate a key into a nonexistent pointer.
      const bytes = byteLength(key);
      if (keys.length < 12 && keyBytes + bytes <= 256) {
        keys.push(key);
        keyBytes += bytes;
      }
    }
    return { keys, keyCount, ...(keyCount > keys.length ? { keysComplete: false } : {}) };
  }
  return {};
}

/** Never invokes toJSON/getters. Rejects unsafe shapes before JSON.stringify sees them. */
function snapshot(
  input: unknown,
  limits: Limits,
  nodeBudget = { remaining: limits.maxNodes },
): RecordValue {
  const seen = new Set<object>();
  let bytes = 0;
  const add = (text: string): string => {
    bytes += byteLength(text);
    if (bytes > limits.itemBytes)
      throw new ResultDataError('item-budget', 'Result exceeds the retained item byte limit.');
    return text;
  };
  const quoted = (value: string): string => {
    if (value.length > limits.itemBytes - bytes)
      throw new ResultDataError('item-budget', 'Result exceeds the retained item byte limit.');
    return add(JSON.stringify(value));
  };
  const visit = (value: unknown, depth: number): string => {
    if (--nodeBudget.remaining < 0)
      throw new ResultDataError('node-budget', 'Result exceeds the traversal limit.');
    if (depth > limits.maxDepth)
      throw new ResultDataError('depth-budget', 'Result exceeds the nesting limit.');
    if (value === null || value === undefined) return add('null');
    if (typeof value === 'string') return quoted(value);
    if (typeof value === 'boolean') return add(String(value));
    if (typeof value === 'number' && Number.isFinite(value)) return add(String(value));
    if (typeof value !== 'object')
      throw new ResultDataError('unsupported-value', 'Result contains a non-JSON value.');
    if (value instanceof Error) {
      const message = Object.getOwnPropertyDescriptor(value, 'message');
      return visit(
        {
          error:
            typeof message?.value === 'string'
              ? message.value.slice(0, 512)
              : 'Tool operation failed.',
        },
        depth + 1,
      );
    }
    if (seen.has(value)) throw new ResultDataError('cyclic-value', 'Result contains a cycle.');
    if (
      !Array.isArray(value) &&
      Object.getPrototypeOf(value) !== Object.prototype &&
      Object.getPrototypeOf(value) !== null
    ) {
      throw new ResultDataError('unsupported-value', 'Result must contain plain JSON objects.');
    }
    seen.add(value);
    const fields: string[] = [];
    if (Array.isArray(value)) {
      if (value.length > nodeBudget.remaining)
        throw new ResultDataError('node-budget', 'Result exceeds the traversal limit.');
      add('[');
      for (let index = 0; index < value.length; index++) {
        if (index) add(',');
        const property = Object.getOwnPropertyDescriptor(value, String(index));
        if (property && !('value' in property))
          throw new ResultDataError('accessor-value', 'Result contains an accessor.');
        fields.push(visit(property?.value, depth + 1));
      }
      add(']');
      seen.delete(value);
      return `[${fields.join(',')}]`;
    }
    add('{');
    for (const key in value) {
      const property = Object.getOwnPropertyDescriptor(value, key);
      if (!property?.enumerable) continue;
      if (!('value' in property))
        throw new ResultDataError('accessor-value', 'Result contains an accessor.');
      if (property.value === undefined) continue;
      if (fields.length) add(',');
      const name = quoted(key);
      add(':');
      fields.push(`${name}:${visit(property.value, depth + 1)}`);
    }
    add('}');
    seen.delete(value);
    return `{${fields.join(',')}}`;
  };
  const json = visit(input, 0);
  return { value: JSON.parse(json) as Json, bytes, ref: '' };
}

function failure(code: string, message: string): Json {
  return { complete: false, storageError: { code, message } };
}

/** Carry effect outcome flags independently of large payloads; never copy exception stacks. */
function metadata(value: unknown): { [key: string]: Json } {
  if (!value || typeof value !== 'object') return {};
  const result: { [key: string]: Json } = {};
  for (const key of ['ok', 'error', 'verification', 'recoveryPending', 'truncated', 'status']) {
    const property = Object.getOwnPropertyDescriptor(value, key);
    if (!property || !('value' in property)) continue;
    try {
      const item = snapshot(property.value, {
        turnBytes: 1024,
        inlineBytes: 256,
        itemBytes: 512,
        totalBytes: 1024,
        maxItems: 1,
        maxDepth: 4,
        maxNodes: 32,
        maxResults: 1,
      });
      result[key] = item.value;
    } catch {
      if (typeof property.value === 'string') result[key] = property.value.slice(0, 80);
      else if (property.value && typeof property.value === 'object') {
        const status = Object.getOwnPropertyDescriptor(property.value, 'status');
        if (typeof status?.value === 'string') result[key] = { status: status.value.slice(0, 80) };
      }
      result.metadataPreview = true;
    }
  }
  return result;
}

export class CommandResultStore {
  private readonly limits: Limits;
  private readonly records = new Map<string, RecordValue>();
  private readonly stringCursors = new Map<string, StringCursor>();
  private scope = (++nextScope).toString(36);
  private sequence = 0;
  private bytes = 0;

  constructor(options: CommandResultStoreOptions = {}) {
    this.limits = {
      turnBytes: 16 * 1024,
      inlineBytes: 2 * 1024,
      itemBytes: 8 * 1024 * 1024,
      totalBytes: 16 * 1024 * 1024,
      maxItems: 128,
      maxDepth: 32,
      maxNodes: 200_000,
      maxResults: 1024,
      ...options,
    };
    for (const [key, value] of Object.entries(this.limits)) {
      if (!Number.isSafeInteger(value) || value < 1)
        throw new RangeError(`${key} must be a positive safe integer.`);
    }
    if (this.limits.turnBytes < 1024) throw new RangeError('turnBytes must be at least 1024.');
    this.limits.inlineBytes = Math.min(
      this.limits.inlineBytes,
      Math.floor(this.limits.turnBytes / 2),
    );
  }

  clear(): void {
    this.records.clear();
    this.stringCursors.clear();
    this.bytes = 0;
    this.sequence = 0;
    this.scope = (++nextScope).toString(36);
  }

  private retain(record: RecordValue): Json {
    if (
      this.records.size >= this.limits.maxItems ||
      this.bytes + record.bytes > this.limits.totalBytes
    ) {
      return {
        ...metadata(record.value),
        ...(failure(
          'store-budget',
          'Result storage is full. Narrow the read or start a new task.',
        ) as object),
      } as Json;
    }
    record.ref = `result:${this.scope}:${++this.sequence}`;
    this.records.set(record.ref, record);
    this.bytes += record.bytes;
    return {
      ...metadata(record.value),
      ref: record.ref,
      type: typeOf(record.value),
      ...shapeOf(record.value, (value) => this.stringCursor(value).length),
      bytes: record.bytes,
      complete: false,
    };
  }

  encode(results: readonly unknown[]): EncodedCommandResults {
    let inputBytes = 2;
    let inputBytesComplete = true;
    let errors = 0;
    const initialRecords = this.records.size;
    const nodes = { remaining: this.limits.maxNodes };
    const raw: Array<RecordValue | undefined> = [];
    let output: Json[] = [];
    if (results.length > Math.min(this.limits.maxNodes, this.limits.maxResults)) {
      output = [
        failure(
          'node-budget',
          `The batch contains ${results.length} results and exceeds the traversal limit. No result was retained.`,
        ),
      ];
      errors = results.length;
      inputBytesComplete = false;
    } else
      for (let index = 0; index < results.length; index++) {
        let record: RecordValue;
        try {
          record = snapshot(results[index], this.limits, nodes);
          inputBytes += record.bytes + (index ? 1 : 0);
        } catch (error) {
          inputBytesComplete = false;
          errors++;
          raw.push(undefined);
          output.push({
            ...metadata(results[index]),
            ...(failure(
              error instanceof ResultDataError ? error.code : 'unsupported-value',
              error instanceof ResultDataError
                ? error.message
                : 'Result could not be read as plain JSON.',
            ) as object),
          } as Json);
          continue;
        }
        raw.push(record);
        const value = record.bytes <= this.limits.inlineBytes ? record.value : this.retain(record);
        if (!record.ref && record.bytes > this.limits.inlineBytes) errors++;
        output.push(value);
      }
    let text = JSON.stringify(output);
    if (byteLength(text) > this.limits.turnBytes) {
      // Budget the entire response, not just individual tool results.
      output = output.map((value, index) => {
        const record = raw[index];
        if (!record || record.ref || record.bytes > this.limits.inlineBytes) return value;
        const receipt = this.retain(record);
        if (!record.ref) errors++;
        return receipt;
      });
      text = JSON.stringify(output);
    }
    if (byteLength(text) > this.limits.turnBytes) {
      // Store the receipt index when possible. A virtual index still exposes all retained data
      // if quota prevents storing that index; the loss of inline/error details is explicit.
      let indexReceipt: Json;
      try {
        indexReceipt = this.retain(snapshot(output, this.limits));
      } catch {
        indexReceipt = failure(
          'index-budget',
          'The receipt index exceeds the result storage limit.',
        );
      }
      if (typeof indexReceipt === 'object' && indexReceipt && 'ref' in indexReceipt) {
        text = JSON.stringify([
          { ...indexReceipt, type: 'result-index', total: results.length, errors },
        ]);
      } else {
        errors++;
        text = JSON.stringify([
          {
            ref: `result:${this.scope}:index`,
            type: 'result-index',
            total: this.records.size,
            complete: false,
            storageError: {
              code: 'turn-budget',
              message:
                'The full receipt batch could not fit or be retained. Inspect this index for every retained result. Some result details are unavailable.',
            },
            errors,
          },
        ]);
      }
    }
    return {
      text,
      inputBytes,
      inputBytesComplete,
      outputBytes: byteLength(text),
      retained: this.records.size - initialRecords,
      errors,
    };
  }

  /**
   * JSON Pointer plus offset/limit pagination. Default pages fit the inline threshold so
   * feeding an inspected page back through encode does not create another opaque handle.
   * Custom inline thresholds below 1024 bytes retain a 1024-byte direct-inspection floor;
   * those configurations should pass an explicit budget when choosing a larger UI page.
   */
  inspect(
    selector: string,
    maxBytes = Math.min(this.limits.turnBytes, Math.max(1024, this.limits.inlineBytes)),
  ): unknown {
    try {
      if (!Number.isSafeInteger(maxBytes) || maxBytes < 1024)
        throw new ResultDataError('inspection-budget', 'Inspection requires at least 1024 bytes.');
      const budget = Math.min(maxBytes, this.limits.turnBytes);
      if (selector.length > 2048)
        throw new ResultDataError('selector', 'Result selector exceeds 2048 characters.');
      const matchRef = /^(\S+)([\s\S]*)$/.exec(selector.trim());
      const ref = matchRef?.[1] ?? '';
      if (!new RegExp(`^result:${this.scope}:(?:[1-9][0-9]*|index)$`).test(ref))
        throw new ResultDataError(
          'reference',
          'Result reference is invalid, expired, or belongs to another task.',
        );
      const options = new Map<string, string>();
      let remaining = matchRef?.[2] ?? '';
      while (remaining.trim()) {
        const match = /^\s+(path|offset|limit)=("(?:[^"\\]|\\.)*"|[^\s"]*)(?=\s|$)/.exec(remaining);
        if (!match || options.has(match[1]!))
          throw new ResultDataError(
            'selector',
            'Use path=/json/pointer offset=N limit=N once each.',
          );
        options.set(
          match[1]!,
          match[2]!.startsWith('"') ? (JSON.parse(match[2]!) as string) : match[2]!,
        );
        remaining = remaining.slice(match[0].length);
      }
      const path = options.get('path') ?? '';
      const integer = (name: string, fallback: number): number => {
        const value = options.get(name);
        if (value === undefined) return fallback;
        if (!/^(0|[1-9][0-9]*)$/.test(value) || !Number.isSafeInteger(Number(value)))
          throw new ResultDataError('selector', `${name} must be a non-negative safe integer.`);
        return Number(value);
      };
      const offset = integer('offset', 0);
      const limit = integer('limit', 20);
      if (limit < 1 || limit > 200)
        throw new ResultDataError('selector', 'limit must be between 1 and 200.');
      let value: Json;
      if (ref.endsWith(':index'))
        value = [...this.records.values()].map((record) => ({
          ref: record.ref,
          type: typeOf(record.value),
          ...shapeOf(record.value, (value) => this.stringCursor(value).length),
          bytes: record.bytes,
          ...metadata(record.value),
        }));
      else {
        const record = this.records.get(ref);
        if (!record)
          throw new ResultDataError('reference', 'Result reference was not found in this task.');
        value = record.value;
      }
      if (path && !path.startsWith('/'))
        throw new ResultDataError('selector', 'path must be an RFC 6901 JSON Pointer.');
      for (const encoded of path ? path.slice(1).split('/') : []) {
        if (/~(?:[^01]|$)/.test(encoded))
          throw new ResultDataError('selector', 'Invalid JSON Pointer escape.');
        const key = encoded.replace(/~1/g, '/').replace(/~0/g, '~');
        if (forbidden.has(key))
          throw new ResultDataError('selector', 'Prototype properties cannot be selected.');
        if (typeof value !== 'object' || !value || !Object.hasOwn(value, key))
          throw new ResultDataError('path', 'The requested result path does not exist.');
        if (Array.isArray(value) && !/^(0|[1-9][0-9]*)$/.test(key))
          throw new ResultDataError('path', 'Array paths require an element index.');
        value = (value as { [key: string]: Json })[key]!;
      }
      const page = this.page(ref, path, value, offset, limit, budget);
      // Long legal pointers and large individual fields must also respect the envelope budget.
      return snapshot(page, { ...this.limits, itemBytes: budget }).value;
    } catch (error) {
      return {
        error: error instanceof ResultDataError ? error.message : 'Result could not be inspected.',
        code: error instanceof ResultDataError ? error.code : 'inspect',
        complete: false,
      };
    }
  }

  private stringCursor(value: string): StringCursor {
    const cached = this.stringCursors.get(value);
    if (cached) return cached;
    let length = 0;
    for (let position = 0; position < value.length; length++) {
      position += value.codePointAt(position)! > 0xffff ? 2 : 1;
    }
    const cursor = { length, offset: 0, position: 0 };
    // Keys reference immutable strings already retained by the task. Only derived cursor
    // metadata is evicted; retained result payloads are never evicted.
    if (this.stringCursors.size >= this.limits.maxItems) {
      const oldest = this.stringCursors.keys().next().value as string | undefined;
      if (oldest !== undefined) this.stringCursors.delete(oldest);
    }
    this.stringCursors.set(value, cursor);
    return cursor;
  }

  private stringWindow(value: string, offset: number, limit: number): string[] {
    const cursor = this.stringCursor(value);
    let position = offset >= cursor.offset ? cursor.position : 0;
    let current = offset >= cursor.offset ? cursor.offset : 0;
    while (current < offset) {
      position += value.codePointAt(position)! > 0xffff ? 2 : 1;
      current++;
    }
    const window: string[] = [];
    while (window.length < limit && position < value.length) {
      const width = value.codePointAt(position)! > 0xffff ? 2 : 1;
      window.push(value.slice(position, position + width));
      position += width;
      current++;
    }
    cursor.offset = current;
    cursor.position = position;
    return window;
  }

  private page(
    ref: string,
    path: string,
    value: Json,
    offset: number,
    limit: number,
    budget: number,
  ): Json {
    const isArray = Array.isArray(value);
    const isObject = typeof value === 'object' && value !== null && !isArray;
    const keys = isObject ? Object.keys(value) : [];
    // String offsets count Unicode code points, so pagination never splits a surrogate pair.
    const total = isArray
      ? value.length
      : isObject
        ? keys.length
        : typeof value === 'string'
          ? this.stringCursor(value).length
          : 1;
    if (offset > total)
      throw new ResultDataError('offset', 'offset exceeds the selected value length.');
    let count = Math.min(limit, total - offset);
    const characters = typeof value === 'string' ? this.stringWindow(value, offset, count) : [];
    const item = (index: number): Json =>
      isArray ? value[index]! : isObject ? keys[index]! : value;
    const make = (): Json => {
      const preview: Json =
        typeof value === 'string'
          ? characters.slice(0, count).join('')
          : isArray || isObject
            ? Array.from({ length: count }, (_, index) => item(offset + index))
            : value;
      const complete = offset === 0 && count === total && !isObject;
      return {
        ref,
        path,
        type: typeOf(value),
        total,
        offset,
        count,
        complete,
        ...(isObject
          ? {
              projection: 'keys',
              help: 'Select a value with path=/key; escape / as ~1 and ~ as ~0.',
            }
          : {}),
        preview,
        ...(offset + count < total
          ? {
              next: `${ref}${path ? ` path=${JSON.stringify(path)}` : ''} offset=${offset + count} limit=${limit}`,
            }
          : {}),
      };
    };
    let response = make();
    while (count > 0) {
      try {
        return snapshot(response, { ...this.limits, itemBytes: budget }).value;
      } catch {
        count = Math.floor(count / 2);
        response = make();
      }
    }
    if (count === 0 && offset < total)
      return {
        ref,
        path,
        type: typeOf(value),
        total,
        offset,
        count: 0,
        complete: false,
        error:
          'One element exceeds the inspection byte budget. Select it with a deeper JSON Pointer.',
        nextPath: `${path}/${isObject ? keys[offset]!.replace(/~/g, '~0').replace(/\//g, '~1') : offset}`,
      };
    // All values are detached before escaping the store, including primitive/object pages.
    return JSON.parse(JSON.stringify(response)) as Json;
  }
}
