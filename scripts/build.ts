/**
 * Build Apps Script bundle from src/ ES modules.
 *
 * Uses Vite (Rollup) to produce a single Code.js with top-level function
 * declarations suitable for Google Apps Script.
 *
 * Usage: npm run build   (or: jiti scripts/build.ts)
 */

/* eslint-disable no-console -- CLI script output. */

import type { Plugin } from 'vite';

import {
  dirname,
  resolve
} from 'node:path';
import { fileURLToPath } from 'node:url';
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
        }
      },
      target: 'es2019'
    },
    logLevel: 'info',
    plugins: [stripExports()],
    root: ROOT
  });

  console.log('Build complete: dist/Code.js');
}

/**
 * Vite plugin that strips ES module export statements from the bundle,
 * making the output compatible with Google Apps Script (global scope).
 */
function stripExports(): Plugin {
  return {
    enforce: 'post',
    name: 'strip-exports',
    renderChunk(code: string): string {
      return code.replace(/^export\s*\{[^}]*\};\s*$/gm, '');
    }
  };
}

await main();

/* eslint-enable no-console -- End CLI script output. */
