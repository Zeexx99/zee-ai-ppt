import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { getHttpsServerOptions } from 'office-addin-dev-certs';

// https://vitejs.dev/config/
export default defineConfig(async () => {
  let httpsConfig: unknown = true;
  try {
    httpsConfig = await getHttpsServerOptions();
  } catch (e) {
    console.warn('Could not load office-addin-dev-certs, falling back to default https:', e);
  }

  return {
    plugins: [react()],
    base: './', // CRITICAL: Use relative paths so assets work on GitHub Pages subpaths
    server: {
      port: 3300,
      https: httpsConfig as boolean | object,
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
  };
});
