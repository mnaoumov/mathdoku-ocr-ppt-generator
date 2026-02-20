import type { CageRaw } from './Puzzle.ts';

import { cellRefA1 } from './parsers.ts';
import { ensureNonNullable } from './typeGuards.ts';

export interface GridBoundaries {
  readonly horizontalBounds: readonly (readonly boolean[])[];
  readonly verticalBounds: readonly (readonly boolean[])[];
}

const BINARY_OP_SIZE = 2;

export function computeGridBoundaries(cages: readonly CageRaw[], puzzleSize: number): GridBoundaries {
  const cellToCage: Record<string, number> = {};
  for (let cageIndex = 0; cageIndex < cages.length; cageIndex++) {
    const cage = ensureNonNullable(cages[cageIndex]);
    for (const cell of cage.cells) {
      cellToCage[cell] = cageIndex;
    }
  }

  const verticalBounds: boolean[][] = [];
  const horizontalBounds: boolean[][] = [];
  for (let r = 0; r < puzzleSize; r++) {
    const row: boolean[] = [];
    verticalBounds[r] = row;
    for (let c = 1; c < puzzleSize; c++) {
      row[c - 1] = cellToCage[cellRefA1(r, c - 1)] !== cellToCage[cellRefA1(r, c)];
    }
  }
  for (let r = 1; r < puzzleSize; r++) {
    const row: boolean[] = [];
    horizontalBounds[r - 1] = row;
    for (let c = 0; c < puzzleSize; c++) {
      row[c] = cellToCage[cellRefA1(r - 1, c)] !== cellToCage[cellRefA1(r, c)];
    }
  }

  return { horizontalBounds, verticalBounds };
}

export function evaluateTuple(tuple: readonly number[], operator: string): null | number {
  if (tuple.length === 0) {
    return null;
  }
  if (tuple.length === 1) {
    return ensureNonNullable(tuple[0]);
  }

  const a = ensureNonNullable(tuple[0]);
  const b = ensureNonNullable(tuple[1]);

  switch (operator) {
    case '-': {
      if (tuple.length !== BINARY_OP_SIZE) {
        return null;
      }
      return Math.abs(a - b);
    }
    case '*':
    case 'x': {
      let product = 1;
      for (const d of tuple) {
        product *= d;
      }
      return product;
    }
    case '/': {
      if (tuple.length !== BINARY_OP_SIZE) {
        return null;
      }
      const maxD = Math.max(a, b);
      const minD = Math.min(a, b);
      if (minD === 0) {
        return null;
      }
      return maxD % minD === 0 ? maxD / minD : null;
    }
    case '+': {
      let sum = 0;
      for (const d of tuple) {
        sum += d;
      }
      return sum;
    }
    default:
      return null;
  }
}

export function generateSubsets<T>(items: readonly T[], subsetSize: number): T[][] {
  if (subsetSize === 0) {
    return [[]];
  }
  if (subsetSize > items.length) {
    return [];
  }

  const results: T[][] = [];
  const indices: number[] = Array.from({ length: subsetSize }, (_, i) => i);

  for (;;) {
    results.push(indices.map((idx) => ensureNonNullable(items[idx])));

    let pos = subsetSize - 1;
    while (pos >= 0 && ensureNonNullable(indices[pos]) === pos + items.length - subsetSize) {
      pos--;
    }
    if (pos < 0) {
      break;
    }
    indices[pos] = ensureNonNullable(indices[pos]) + 1;
    for (let j = pos + 1; j < subsetSize; j++) {
      indices[j] = ensureNonNullable(indices[j - 1]) + 1;
    }
  }

  return results;
}
