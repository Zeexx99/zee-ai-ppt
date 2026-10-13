import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import basicSsl from '@vitejs/plugin-basic-ssl';

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react(), basicSsl()],
  base: './', // CRITICAL: Use relative paths so assets work on GitHub Pages subpaths
  server: {
    port: 3300,
    https: true,
    headers: {
      'Access-Control-Allow-Origin': '*'
    }
  },
  build: {
    outDir: 'dist',
    sourcemap: false, // CRITICAL: Remove source maps from production builds
    rollupOptions: {
      input: {
        main: './index.html'
      }
    }
  },
  // @ts-expect-error vitest options
  test: {
    globals: true,
    environment: 'happy-dom'
  }
});
