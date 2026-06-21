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

## The pipeline

```
RawContent ──normalize──▶ Markdown ──parse──▶ Block[] (kind + char offsets)
          ──chunk──▶ Chunk[] (section-grouped, token-budgeted, anchored)
          ──contextualize + map──▶ ResolvedContext[]  ─▶ SessionContext ─▶ query.parts[]
```

- **`RawContent`** — what a bridge hands in: `{ sourceId, text, format: markdown|plain|html, title?,
  surface?, indexedDocumentName? }`.
- **`Block`** — a structural unit (`heading|paragraph|list|table|code|quote`) with `start`/`end`
  **char offsets** into the source — the basis for anchors.
- **`Chunk`** — token-budgeted text + `ChunkMeta`: `sectionPath` breadcrumb, `kinds`, char range,
  `tokensEstimate`, and an **`Anchor`** (`matchText` + `contextHint` = section path + `locator` =
  `chars:start-end`). Tables are never split; oversized blocks are recursively split with overlap.
- **`ResolvedContext`** — attach-ready: a contextualized `text/markdown` part per chunk (carrying the
  anchor + token estimate on its `ref`), **or** a single `indexed-document` reference when
  `preferReference` + `indexedDocumentName` are set (the reference-over-inline policy).

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
