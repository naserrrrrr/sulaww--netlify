import { defineConfig } from 'vite';
import { viteSingleFile } from 'vite-plugin-singlefile';
import path from 'node:path';

// `npm run build`        -> dist/ (regular multi-file build)
// `npm run build:single` -> dist-single/index.html, one self-contained file
export default defineConfig(({ mode }) => {
  const single = mode === 'single';
  return {
    base: './',
    publicDir: false,
    resolve: { alias: { '@brand': path.resolve(__dirname, '../public/assets/brand') } },
    server: { fs: { allow: [path.resolve(__dirname, '..')] } },
    build: {
      outDir: single ? 'dist-single' : 'dist',
      assetsInlineLimit: single ? 1e9 : 4096,
      cssCodeSplit: false,
      target: 'es2020',
      chunkSizeWarningLimit: 4000,
    },
    plugins: single ? [viteSingleFile({ removeViteModuleLoader: true })] : [],
  };
});
