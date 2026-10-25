import { brandAssets } from './brand-assets.js';
import { computeAssets } from './compute-assets.js';
import { execSync } from 'node:child_process';
import { resolve } from 'node:path';
import { defineConfig, loadEnv } from 'vite';
import react from '@vitejs/plugin-react';
import basicSsl from '@vitejs/plugin-basic-ssl';

const hostFromUrl = (value: string | undefined): string[] => {
  if (!value) return [];
  try {
    return [new URL(value).host];
  } catch {
    return [value.replace(/^https?:\/\//, '').replace(/\/.*$/, '')].filter(Boolean);
  }
};

/**
 * Which build a pasted diagnostic came from: the commit when git is available (a Docker build
 * may have no .git; GE_BUILD_COMMIT overrides) and the build time, which is always known.
 */
function buildId(env: Record<string, string | undefined>): string {
  let commit = env.GE_BUILD_COMMIT;
  if (!commit) {
    try {
      commit = execSync('git rev-parse --short HEAD', { cwd: __dirname, stdio: 'pipe' })
        .toString()
        .trim();
    } catch {
      commit = undefined;
    }
  }
  const builtAt = new Date().toISOString().replace(/\.\d{3}Z$/, 'Z');
  return [commit, builtAt].filter(Boolean).join(' ');
}

/**
 * Build the runnable add-in shell. Its entry HTML pages are:
 *   • index.html     — the hosting status page and Teams launch forwarder.
 *   • taskpane.html  — the React task pane (the document-bound assist panel).
 *   • commands.html  — the headless function-command runtime (ribbon actions + OnMessageSend).
 *   • auth-redirect.html — the minimal MSAL redirect bridge for popup/iframe token flows.
 *
 * Office requires the dev server to be HTTPS (the add-in is iframed by the host over TLS), so
 * `@vitejs/plugin-basic-ssl` mints a local certificate for `bun run dev`.
 */
export default defineConfig(({ mode }) => {
  const env = { ...loadEnv(mode, __dirname, ''), ...process.env };
  const port = Number(env.GE_DEV_PORT ?? process.env.GE_DEV_PORT ?? '13000');
  const useHttps = env.GOOGLE_CLOUD_WORKSTATIONS !== 'true';
  const allowedHosts = Array.from(
    new Set([
      ...hostFromUrl(env.GE_DEV_WEB_ORIGIN),
      ...hostFromUrl(env.GE_DEV_WEB_DOMAIN),
      'equally-surfing-watches-compilation.trycloudflare.com',
    ]),
  );

  return {
    root: __dirname,
    define: { __BUILD_ID__: JSON.stringify(buildId(env)) },
    plugins: [computeAssets(), ...brandAssets(), react(), ...(useHttps ? [basicSsl()] : [])],
    server: {
      https: useHttps ? true : undefined,
      port,
      strictPort: true,
      allowedHosts,
      // The Office host loads the pane cross-origin; allow it.
      cors: true,
    },
    build: {
      outDir: 'dist-web',
      emptyOutDir: true,
      rollupOptions: {
        input: {
          index: resolve(__dirname, 'index.html'),
          taskpane: resolve(__dirname, 'taskpane.html'),
          commands: resolve(__dirname, 'commands.html'),
          functions: resolve(__dirname, 'functions.html'),
          authRedirect: resolve(__dirname, 'auth-redirect.html'),
        },
      },
    },
  };
});
