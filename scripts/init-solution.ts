/**
 * Initialize a solution YAML from a puzzle YAML.
 *
 * Runs init strategies + automated strategies in Node.js, writes <name>.solution.yaml.
 *
 * Usage: npm run init-solution -- path/to/puzzle.yaml
 */

/* eslint-disable no-console -- CLI script output. */

import yaml from 'js-yaml';
import {
  readFileSync,
  writeFileSync
} from 'node:fs';
import {
  basename,
  dirname,
  resolve
} from 'node:path';

import {
  initPuzzleSlides,
  parsePuzzleJson
} from '../src/Puzzle.ts';
import {
  buildPuzzleJson,
  type YamlSpec
} from '../src/puzzleYamlParser.ts';
import {
  buildSolutionYaml,
  SLIDE_PAIR_SIZE
} from '../src/SolutionYaml.ts';
import {
  createInitialStrategies,
  createStrategies
} from '../src/strategies/createDefaultStrategies.ts';
import { SvgRenderer } from '../src/SvgRenderer.ts';

const YAML_ARG_INDEX = 2;

const yamlPath = process.argv[YAML_ARG_INDEX];
if (!yamlPath) {
  console.error('Usage: npm run init-solution -- <puzzle.yaml>');
  process.exit(1);
}

const resolvedPath = resolve(yamlPath);
const yamlContent = readFileSync(resolvedPath, 'utf-8');
const puzzleName = basename(resolvedPath).replace(/\.ya?ml$/i, '');

console.log(`Loading puzzle: ${resolvedPath}`);

const spec = yaml.load(yamlContent) as YamlSpec;
const puzzleJson = parsePuzzleJson(buildPuzzleJson(spec, puzzleName));

const renderer = new SvgRenderer();
renderer.initGrid(
  puzzleJson.puzzleSize,
  puzzleJson.cages,
  puzzleJson.hasOperators ?? true,
  puzzleJson.title ?? '',
  puzzleJson.meta ?? ''
);
renderer.pushInitialSlide();

let initError: null | string = null;
try {
  initPuzzleSlides({
    cages: puzzleJson.cages,
    hasOperators: puzzleJson.hasOperators !== false,
    initialStrategies: createInitialStrategies(),
    meta: puzzleJson.meta ?? '',
    puzzleSize: puzzleJson.puzzleSize,
    renderer,
    strategies: createStrategies(puzzleJson.puzzleSize),
    title: puzzleJson.title ?? ''
  });
} catch (e: unknown) {
  initError = e instanceof Error ? e.message : String(e);
  console.error(`Strategy error: ${initError}`);

  // Append error to the last pending slide's notes so it appears in the YAML step.
  // Pending slides are at odd indices (1, 3, 5, ...).
  const lastPendingIndex = renderer.slides.length % SLIDE_PAIR_SIZE === 0
    ? renderer.slides.length - 1
    : renderer.slides.length - SLIDE_PAIR_SIZE;
  if (lastPendingIndex >= 1) {
    const slide = renderer.slides[lastPendingIndex];
    if (slide) {
      renderer.slides[lastPendingIndex] = {
        ...slide,
        notes: `${slide.notes}\nERROR: ${initError}`
      };
    }
  }
}

const manualNotes = renderer.slides.map((slide) => slide.notes);

const solutionYaml = buildSolutionYaml({
  hasOperators: puzzleJson.hasOperators !== false,
  manualNotes,
  puzzleJson,
  slides: renderer.slides
});

const outputPath = resolve(dirname(resolvedPath), `${puzzleName}.solution.yaml`);
writeFileSync(outputPath, solutionYaml, 'utf-8');

if (initError) {
  console.error(`Partial solution written: ${outputPath}`);
  console.error(`${String(renderer.slides.length)} slides generated (with error)`);
  process.exitCode = 1;
} else {
  console.log(`Solution written: ${outputPath}`);
  console.log(`${String(renderer.slides.length)} slides generated`);
}

/* eslint-enable no-console -- End CLI script output. */
