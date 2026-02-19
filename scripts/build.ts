/**
 * Build Apps Script bundle from src/ ES modules.
 *
 * Uses Vite (Rollup) + rollup-plugin-gas to produce a single Code.js
 * with top-level function declarations suitable for Google Apps Script.
 *
 * Usage: npm run build   (or: jiti scripts/build.ts)
 */

/* eslint-disable no-console -- CLI script output. */

import {
  dirname,
  resolve
} from 'node:path';
import { fileURLToPath } from 'node:url';
import gas from 'rollup-plugin-gas';
import { build } from 'vite';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const SRC = resolve(ROOT, 'src');
const DIST = resolve(ROOT, 'dist');

async function main(): Promise<void> {
  await build({
    build: {
      emptyOutDir: false,
      lib: {
        entry: resolve(SRC, 'main.ts'),
        fileName: 'Code',
        formats: ['es']
      },
      minify: false,
      outDir: DIST,
      rollupOptions: {
        output: {
          entryFileNames: 'Code.js'
        },
        plugins: [gas()]
      }
    },
    logLevel: 'info',
    root: ROOT
  });

  console.log('Build complete: dist/Code.js');
}

await main();

/* eslint-enable no-console -- End CLI script output. */
