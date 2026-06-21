# Content processing — preparing host content for grounding

How `@ge/content` turns raw host/estate content into grounding-ready context. The goal: the
agent operates on **structured, self-situating, metadata-rich** content instead of a raw text dump,
and every piece keeps a **write-back anchor** so the loop (read → ground → actuate) closes. Built on
the consensus of the leading OSS document/RAG libraries, adapted to our client-direct,
`query.parts[]` reality.

## Inspiration (and what we took)

| Library / technique | What we adopted |
|---|---|
| **MarkItDown** (Microsoft), **Docling** (IBM) | Normalize everything to **Markdown** — LLM-native, token-efficient, structure-preserving (headings, GFM tables, lists). The rich format→Markdown conversion belongs in the bridges; `@ge/content` consumes Markdown. |
| **LangChain** `RecursiveCharacterTextSplitter` | **Recursive, separator-aware splitting** of oversized blocks: paragraph → sentence → word, token-aware, with overlap. |
| **LlamaIndex** Hierarchical / parent-child | **Section grouping**: chunks are grouped under their heading breadcrumb and broken at section boundaries; the breadcrumb travels as metadata (small-chunk retrieval, parent-section context). |
| **Anthropic Contextual Retrieval** | **Contextualize** each chunk with a compact `source › section` header so it grounds well out of context (we keep it to one terse line — the engine already does citations). |
| Token-budget guidance | A dependency-free **token estimate** drives chunk sizing and the context-tray budget (we can't ship tiktoken into a webview). |

Sources: [MarkItDown](https://github.com/microsoft/markitdown) ·
[LangChain chunking](https://www.firecrawl.dev/blog/best-chunking-strategies-rag) ·
[Anthropic Contextual Retrieval](https://www.anthropic.com/news/contextual-retrieval).

## Not RAG — and native-first

To be explicit: there is **no embedding model, no vector store, no retrieval** in this package.
Gemini Enterprise does the grounding/retrieval (server-side, over the connected data stores). This
layer only **normalizes + chunks + labels** what the add-in attaches as `query.parts[]`. It's the
opposite of heavyweight.

**Best use of Office = use the native object model, not a string round-trip.** The Office host already
exposes structure — Word `paragraphs`/`tables`/`contentControls` (with ids + styles), Excel `range`
values/addresses, PowerPoint `slides`/shapes. So bridges build `Block[]` **directly** from the object
model via the `native` builders, skipping the Markdown regex reparse and preserving a **native host
locator** for write-back. The Markdown string path is only the *fallback* for sources that genuinely
arrive as a blob (an Outlook HTML body, plain text). (`Document.getFileAsync` OOXML is intentionally
avoided: it's slow for pptx and can't be attached to `streamAssist` anyway — no blob parts.)

## The pipeline

```
 native (preferred):  Office object model ──native builders──▶ Block[] (kind + host locator)
 string (fallback):   text/HTML ──normalize──▶ Markdown ──parse──▶ Block[] (kind + char offsets)
                      ──chunk──▶ Chunk[] (section-grouped, token-budgeted, anchored)
                      ──contextualize + map──▶ ResolvedContext[] ─▶ SessionContext ─▶ query.parts[]
```

- **`NativeContent`** (native) — `{ sourceId, title?, surface?, indexedDocumentName?, blocks: Block[] }`,
  where each `Block` carries a host `locator` (`cc:42`, `range:Sheet1!A1:D9`, `slide:4`). Built with the
  `native.{heading,paragraph,table,listBlock,quote,code,slide}` helpers. → `processNative` / `toContextNative`.
- **`RawContent`** (fallback) — `{ sourceId, text, format: markdown|plain|html, … }`. → `processContent` / `toContext`.
- **`Block`** — a structural unit (`heading|paragraph|list|table|code|quote`); offsets `start`/`end`
  (string path) **or** a native `locator` (object-model path); `data` holds the raw `StructuredData`
  for native tables.
- **`Chunk`** — token-budgeted text + `ChunkMeta`: `sectionPath` breadcrumb, `kinds`, optional char
  range, `tokensEstimate`, and an **`Anchor`** (`matchText` + `contextHint` = section path + `locator`
  = the native handle, else `chars:start-end`). Tables/native structured blocks are never split.
- **`ResolvedContext`** — attach-ready: a contextualized `text/markdown` part per chunk, **or** a single
  `indexed-document` reference when `preferReference` + `indexedDocumentName` are set.

## The interplay (why this composes)

This is the missing middle between the **Access Model** (read = attach, write = actuate) and the
**context mechanisms** (`query.parts[]`):

1. **Read → context.** A bridge extracts host content (Plane A) or estate content (Plane B) as
   Markdown → `processContent`/`toContext` → `ResolvedContext[]`. `SessionContext.add()` attaches
   them; `streamAssist` sends them as a multi-part query (each chunk as data, the user's question
   last). *Verified end-to-end in `gemini-client/pipeline.test.ts`.*
2. **Ground → cite.** The engine returns grounding segments + references; the add-in maps them to
   citation chips and inline highlights (see `context-mechanisms.md`).
3. **Actuate → write back.** Each chunk's `Anchor` is the same shared `@ge/contracts` `Anchor` used
   by `Finding` and actuation targets — so "rewrite this clause" resolves to a `body.search(matchText)`
   range and lands as a tracked change at the exact span the chunk came from.
4. **Budget.** `tokensEstimate` on every `ContextRef` powers the context-tray budget indicator and
   lets the UI prefer references for large items.

So content processing is not a side utility — it's the layer that makes attached context *legible to
the model* and *addressable for write-back*, closing the read→ground→actuate loop.

## Defaults & tuning
`maxTokens` 400 (soft), `overlapTokens` 40, `sectionBreakLevel` 2. Tune per surface: tight budgets +
references for large SharePoint docs; larger chunks for short emails. Prefer references over inline
whenever the source is indexed.
