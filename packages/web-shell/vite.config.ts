import { resolve } from 'node:path';
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import basicSsl from '@vitejs/plugin-basic-ssl';

/**
 * Build the runnable add-in shell. Two entry HTML pages:
 *   • taskpane.html  — the React task pane (the document-bound assist panel).
 *   • commands.html  — the headless function-command runtime (ribbon actions + OnMessageSend).
 *
 * Office requires the dev server to be HTTPS (the add-in is iframed by the host over TLS), so
 * `@vitejs/plugin-basic-ssl` mints a local certificate for `npm run dev`.
 */
export default defineConfig({
  root: __dirname,
  plugins: [react(), basicSsl()],
  server: {
    https: true,
    port: 3000,
    strictPort: true,
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
      },
    },
  },
});
