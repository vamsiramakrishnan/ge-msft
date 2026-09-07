import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { Plugin } from 'vite';
export const COMPUTE_CSP =
  "default-src 'none'; script-src 'self' 'wasm-unsafe-eval'; connect-src blob:; worker-src 'none'";
/** Bundle self-hosted pinned engine assets and serve the worker under its own restricted CSP. */
export function computeAssets(): Plugin {
  const files: Record<string, string> = {
    '/compute/duckdb-worker.js': 'duckdb-browser-mvp.worker.js',
    '/compute/duckdb.wasm': 'duckdb-mvp.wasm',
  };
  const path = (file: string): string =>
    resolve(__dirname, '../../node_modules/@duckdb/duckdb-wasm/dist', file);
  return {
    name: 'compute-assets',
    configureServer(server) {
      server.middlewares.use((req, res, next) => {
        const file = files[(req.url ?? '').split('?')[0]!];
        if (!file) return next();
        res.setHeader(
          'Content-Type',
          file.endsWith('.wasm') ? 'application/wasm' : 'text/javascript',
        );
        if (file.endsWith('.js')) res.setHeader('Content-Security-Policy', COMPUTE_CSP);
        res.end(readFileSync(path(file)));
      });
    },
    generateBundle() {
      for (const [name, file] of Object.entries(files))
        this.emitFile({ type: 'asset', fileName: name.slice(1), source: readFileSync(path(file)) });
    },
  };
}
