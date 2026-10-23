import { existsSync, readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import type { Plugin } from 'vite';
import {
  BRAND_ICONS,
  brandAssets as brandAssetFiles,
  brandText,
  contentHash,
  fillBrandTokens,
  loadBrand,
} from '../../tools/brand/brand.mjs';

/**
 * Apply the build's brand (GE_BRAND → brands/<id>/, see tools/brand/brand.mjs): the names the task
 * pane reads (`__GE_BRAND__`, see src/brand.ts), the %GE_BRAND_*% placeholders in every HTML page,
 * the six add-in icons at the site root, and the optional theme.css and /brand/* assets.
 */
export function brandAssets(): Plugin[] {
  const brand = loadBrand();
  const themeHref = brand.themePath ? `/brand/theme.css?v=${contentHash(brand.themePath)}` : null;
  const files: Record<string, string> = {
    ...Object.fromEntries(BRAND_ICONS.map((icon) => [icon, join(brand.iconsDir, icon)])),
    ...(brand.themePath ? { 'brand/theme.css': brand.themePath } : {}),
    ...brandAssetFiles(brand),
  };
  let outDir = '';
  return [
    {
      name: 'brand-assets',
      config: () => ({ define: { __GE_BRAND__: JSON.stringify(brandText(brand)) } }),
      configResolved(config) {
        outDir = resolve(config.root, config.build.outDir);
      },
      transformIndexHtml: { order: 'pre', handler: (html) => fillBrandTokens(html, brand) },
      configureServer(server) {
        const publicDir = resolve(server.config.root, 'public');
        server.middlewares.use((req, res, next) => {
          const path = (req.url ?? '').split('?')[0]!.slice(1);
          const file = files[path];
          if (file) {
            if (path.endsWith('.css')) res.setHeader('Content-Type', 'text/css');
            res.end(readFileSync(file));
            return;
          }
          // Static pages under public/ skip transformIndexHtml; fill their placeholders here.
          const page = join(publicDir, path);
          if (
            path.endsWith('.html') &&
            existsSync(page) &&
            !existsSync(join(server.config.root, path))
          ) {
            res.setHeader('Content-Type', 'text/html');
            res.end(fillBrandTokens(readFileSync(page, 'utf8'), brand));
            return;
          }
          next();
        });
      },
      generateBundle() {
        for (const [fileName, file] of Object.entries(files))
          this.emitFile({ type: 'asset', fileName, source: readFileSync(file) });
      },
      closeBundle() {
        // public/ is copied verbatim, so fill the placeholders in its pages once the build is written.
        if (!existsSync(outDir)) return;
        for (const page of readdirSync(outDir).filter((f) => f.endsWith('.html'))) {
          const path = join(outDir, page);
          const html = readFileSync(path, 'utf8');
          const filled = fillBrandTokens(html, brand);
          if (filled !== html) writeFileSync(path, filled);
        }
      },
    },
    {
      // After Vite's own HTML processing, so the theme link isn't treated as a bundled asset. It opens
      // <body> so it follows every head stylesheet, including the <style> tags Vite adds in dev.
      name: 'brand-theme',
      transformIndexHtml: {
        order: 'post',
        handler: () =>
          themeHref
            ? [
                {
                  tag: 'link',
                  attrs: { rel: 'stylesheet', href: themeHref },
                  injectTo: 'body-prepend',
                },
              ]
            : [],
      },
    },
  ];
}
