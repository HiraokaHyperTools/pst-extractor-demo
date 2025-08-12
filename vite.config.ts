import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { nodePolyfills } from "vite-plugin-node-polyfills";
import process from 'process';

// https://vite.dev/config/
export default defineConfig({
  plugins: [react(), nodePolyfills()],
  build: {
    outDir: 'docs',
  },
  base: '/pst-extractor-demo/',
  define: {
    APP_VERSION: JSON.stringify(process.env.npm_package_version),
  }
});
