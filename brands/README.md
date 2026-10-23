# Brands

A brand is everything users see that names or styles the add-in: the product name, the icons, the
font and the colours. Each brand is one folder here, and the build picks one with `GE_BRAND`:

```bash
GE_BRAND=acme bun run build     # uses brands/acme/
bun run build                   # brands/current if present, else brands/default/ (the upstream look)
```

To make a brand the default for a repository, put its id in `brands/current` (one line, e.g.
`acme`). Upstream has no `brands/current`, so it always builds the default look, and the file never
conflicts with upstream updates. `GE_BRAND` still overrides it for a single command.

The web build (`packages/web-shell/brand-assets.ts`) and the manifest tools (`tools/release`) both
read the brand through `tools/brand/brand.mjs`, so the task pane and the Office ribbon always agree.

## What a brand folder holds

| File | Required | What it does |
| --- | --- | --- |
| `brand.json` | Yes | `name` (header, errors, ribbon), `assistantName` ("Ask …", "Let … choose"), `fullName` (page titles), `fontStylesheet` (a Google Fonts URL); optional `tagline` (line under the header name), `promptName` (the name the model calls itself in answers; unset keeps the engine's default, "Gemini Enterprise") and `promptRules` (writing rules sent with that name on every turn; a rule is text, or `{ "text": …, "surfaces": ["outlook"], "turns": ["command"] }` to send it only from those Office apps and only on those turns: `chat` (answers), `planner`, `command` (the assistant writes into the document or email), e.g. how emails open) |
| `icons/` | Yes | `icon-16.png`, `icon-32.png`, `icon-80.png` (ribbon, exact sizes), `icon-64.png` (the XML manifests' high-resolution icon; can be larger, e.g. 256×256, so it stays sharp on Retina screens), `icon-color.png` (192×192), `icon-outline.png` (32×32, white on transparent) |
| `theme.css` | No | CSS loaded after the task-pane styles; override the `:root` colour tokens and any rules you need |
| `assets/` | No | Extra files served at `/brand/<file>`, for example a logo that `theme.css` points at |

The build fails if a required field or icon is missing, so a half-branded add-in can't ship.

## Adding a brand

1. Copy `brands/default/` to `brands/<id>/` (lowercase letters, digits and dashes).
2. Edit `brand.json` and replace the six icons.
3. Add `theme.css` for colours and fonts, and put any images it uses in `assets/`.
4. Build with `GE_BRAND=<id>` and check the task pane, the ribbon and the page titles.

Keep a customer's brand folder in that customer's own repository. Upstream only carries
`brands/default/`, so taking upstream updates never touches another brand's files.

## What is not brand text

Code identifiers stay the same for every brand: `openGemini`, `GeminiPane`, the `Gemini.*` manifest
resource ids, and `view: 'Gemini'`. Office binds ribbon buttons to them. Messages about the Gemini
Enterprise service itself (for example skill resource errors in `config.ts`) also stay, because they
name the backend the add-in talks to.
