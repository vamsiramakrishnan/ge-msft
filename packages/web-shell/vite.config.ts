import { resolve } from 'node:path';
import { defineConfig, loadEnv } from 'vite';
import react from '@vitejs/plugin-react';
import basicSsl from '@vitejs/plugin-basic-ssl';

/**
 * Build the runnable add-in shell. Two entry HTML pages:
 *   • taskpane.html  — the React task pane (the document-bound assist panel).
 *   • commands.html  — the headless function-command runtime (ribbon actions + OnMessageSend).
 *   • auth-redirect.html — the minimal MSAL redirect bridge for popup/iframe token flows.
 *
 * Office requires the dev server to be HTTPS (the add-in is iframed by the host over TLS), so
 * `@vitejs/plugin-basic-ssl` mints a local certificate for `npm run dev`.
 */
export default defineConfig(({ mode }) => {
  const env = { ...loadEnv(mode, __dirname, ''), ...process.env };
  const port = Number(env.GE_DEV_PORT ?? process.env.GE_DEV_PORT ?? '13000');
  const useHttps = env.GOOGLE_CLOUD_WORKSTATIONS !== 'true';

  return {
    root: __dirname,
    plugins: [react(), ...(useHttps ? [basicSsl()] : [])],
    server: {
      https: useHttps ? true : undefined,
      port,
      strictPort: true,
      allowedHosts: true,
      // The Office host loads the pane cross-origin; allow it.
      cors: true,
    },
    build: {
      outDir: 'dist-web',
      emptyOutDir: true,
      rollupOptions: {
        input: {
          taskpane: resolve(__dirname, 'taskpane.html'),
          commands: resolve(__dirname, 'commands.html'),
          authRedirect: resolve(__dirname, 'auth-redirect.html'),
        },
      },
    },
  };
});
