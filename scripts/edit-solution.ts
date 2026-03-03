/**
 * Start the Mathdoku solver dev server with a pre-loaded solution YAML.
 *
 * Usage: npm run edit-solution -- path/to/solution.yaml
 */

/* eslint-disable no-console -- CLI script output. */

import {
  readFileSync,
  writeFileSync
} from 'node:fs';
import {
  basename,
  dirname,
  resolve
} from 'node:path';
import { fileURLToPath } from 'node:url';
import { createServer } from 'vite';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const APP_SRC = resolve(ROOT, 'src', 'app');
const YAML_ARG_INDEX = 2;

const solutionPath = process.argv[YAML_ARG_INDEX];
if (!solutionPath) {
  console.error('Usage: npm run edit-solution -- <solution.yaml>');
  process.exit(1);
}

const resolvedPath = resolve(solutionPath);
const solutionContent = readFileSync(resolvedPath, 'utf-8');
const solutionName = basename(resolvedPath).replace(/\.solution\.ya?ml$/i, '').replace(/\.ya?ml$/i, '');

console.log(`Loading solution: ${resolvedPath}`);

const server = await createServer({
  plugins: [
    {
      configureServer(srv): void {
        srv.middlewares.use('/api/solution', (req, res) => {
          if (req.method === 'POST') {
            const chunks: Buffer[] = [];
            req.on('data', (chunk: Buffer) => {
              chunks.push(chunk);
            });
            req.on('end', () => {
              const body = Buffer.concat(chunks).toString('utf-8');
              writeFileSync(resolvedPath, body, 'utf-8');
              res.setHeader('Content-Type', 'application/json');
              res.end(JSON.stringify({ ok: true }));
            });
            return;
          }
          res.setHeader('Content-Type', 'application/json');
          res.end(JSON.stringify({ content: solutionContent, name: solutionName }));
        });
      },
      name: 'solution-yaml-middleware'
    }
  ],
  root: APP_SRC,
  server: {
    open: true
  }
});

await server.listen();
server.printUrls();

/* eslint-enable no-console -- End CLI script output. */
