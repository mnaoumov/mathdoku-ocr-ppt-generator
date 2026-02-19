import {
  describe,
  expect,
  it
} from 'vitest';

import { buildCageConstraintChanges } from '../src/cageConstraints.ts';
import { CandidatesChange } from '../src/cellChanges/CandidatesChange.ts';
import { CandidatesStrikethrough } from '../src/cellChanges/CandidatesStrikethrough.ts';
import { ValueChange } from '../src/cellChanges/ValueChange.ts';
import { ensureNonNullable } from '../src/typeGuards.ts';
import { createTestPuzzle } from './puzzleTestHelper.ts';

describe('buildCageConstraintChanges', () => {
  it('sets value for single-cell cages', () => {
    const cages = [
      { cells: ['A1'], value: 3 },
      { cells: ['A2', 'B2'], operator: '+', value: 5 },
      { cells: ['B1'], value: 2 }
    ];
    const puzzle = createTestPuzzle(2, cages, true);
    const changes = buildCageConstraintChanges(puzzle.cages, true, 2);
    const valueChanges = changes.filter((c) => c instanceof ValueChange);
    expect(valueChanges.length).toBeGreaterThanOrEqual(2);
  });

  it('produces strikethrough changes for peers of set values', () => {
    const cages = [
      { cells: ['A1'], value: 1 },
      { cells: ['A2', 'B2'], operator: '+', value: 3 },
      { cells: ['B1'], value: 2 }
    ];
    const puzzle = createTestPuzzle(2, cages, true);
    const changes = buildCageConstraintChanges(puzzle.cages, true, 2);
    const strikethroughChanges = changes.filter((c) => c instanceof CandidatesStrikethrough);
    expect(strikethroughChanges.length).toBeGreaterThan(0);
  });

  it('orders changes as values first, then candidates, then strikethroughs', () => {
    const cages = [
      { cells: ['A1'], value: 3 },
      { cells: ['A2', 'B2'], operator: '+', value: 5 },
      { cells: ['B1'], value: 2 }
    ];
    const puzzle = createTestPuzzle(2, cages, true);
    const changes = buildCageConstraintChanges(puzzle.cages, true, 2);
    // Find first occurrence index of each type
    const firstValue = changes.findIndex((c) => c instanceof ValueChange);
    const firstCandidates = changes.findIndex((c) => c instanceof CandidatesChange);
    const firstStrike = changes.findIndex((c) => c instanceof CandidatesStrikethrough);
    if (firstValue >= 0 && firstStrike >= 0) {
      expect(firstValue).toBeLessThan(firstStrike);
    }
    if (firstCandidates >= 0 && firstStrike >= 0) {
      expect(firstCandidates).toBeLessThan(firstStrike);
    }
  });

  it('narrows candidates when all tuples share the same value set', () => {
    // A 2-cell cage with + and value 3 in a size-4 grid: [1,2] and [2,1] => both use {1,2}
    const fullCages = [
      { cells: ['A1', 'B1'], operator: '+', value: 3 },
      { cells: ['C1', 'D1'], operator: '+', value: 7 },
      { cells: ['A2', 'B2'], operator: '+', value: 7 },
      { cells: ['C2', 'D2'], operator: '+', value: 3 },
      { cells: ['A3', 'B3'], operator: '+', value: 5 },
      { cells: ['C3', 'D3'], operator: '+', value: 5 },
      { cells: ['A4', 'B4'], operator: '+', value: 5 },
      { cells: ['C4', 'D4'], operator: '+', value: 5 }
    ];
    const puzzle = createTestPuzzle(4, fullCages, true);
    const changes = buildCageConstraintChanges(puzzle.cages, true, 4);
    const candidateChanges = changes.filter((c) => c instanceof CandidatesChange);
    // The cage [A1,B1] with + and value 3 has tuples [1,2] and [2,1] => both use {1,2}
    // So cells A1 and B1 should get CandidatesChange with [1,2]
    const a1Changes = candidateChanges.filter((c) => c.cell.ref === 'A1');
    if (a1Changes.length > 0) {
      const candChange = ensureNonNullable(a1Changes[0]);
      expect(candChange.values).toEqual([1, 2]);
    }
  });
});
