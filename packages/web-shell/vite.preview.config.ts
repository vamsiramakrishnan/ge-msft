import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

/**
 * The preview-harness dev server (`npm run preview -w packages/web-shell`). Unlike the add-in dev
 * server (vite.config.ts), this is plain HTTP and opens `preview.html` directly — the harness has
 * no Office host and no TLS requirement, so it renders the panel in any browser tab. View layer
 * only; it never ships in the add-in bundle.
 */
export default defineConfig({
  root: __dirname,
  plugins: [react()],
  server: {
    port: 3100,
    strictPort: true,
    open: '/preview.html',
  },
});
