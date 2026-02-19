import {
  describe,
  expect,
  it
} from 'vitest';

import { ValueChange } from '../../src/cellChanges/ValueChange.ts';
import { HiddenSingleStrategy } from '../../src/strategies/HiddenSingleStrategy.ts';
import { ensureNonNullable } from '../../src/typeGuards.ts';
import { createTestPuzzle } from '../puzzleTestHelper.ts';

describe('HiddenSingleStrategy', () => {
  const strategy = new HiddenSingleStrategy();

  it('returns null when no hidden singles exist', () => {
    const cages = [
      { cells: ['A1', 'B1'], operator: '+', value: 3 },
      { cells: ['A2', 'B2'], operator: '+', value: 3 }
    ];
    const puzzle = createTestPuzzle(2, cages, true);
    // Both cells in each row have both candidates
    for (const cell of puzzle.cells) {
      cell.setCandidates([1, 2]);
    }
    expect(strategy.tryApply(puzzle)).toBeNull();
  });

  it('finds hidden single when digit appears in only one cell in a house', () => {
    const cages = [
      { cells: ['A1', 'B1'], operator: '+', value: 3 },
      { cells: ['A2', 'B2'], operator: '+', value: 3 }
    ];
    const puzzle = createTestPuzzle(2, cages, true);
    // In row 1: A1 has [1,2], B1 has [2] only
    // Digit 1 only in A1 in row 1 → hidden single
    puzzle.getCell('A1').setCandidates([1, 2]);
    puzzle.getCell('B1').setCandidates([2]);
    puzzle.getCell('A2').setCandidates([1, 2]);
    puzzle.getCell('B2').setCandidates([1, 2]);

    const result = strategy.tryApply(puzzle);
    expect(result).not.toBeNull();
    const valueChanges = ensureNonNullable(result).filter((c) => c instanceof ValueChange);
    // A1 should be set to 1 (hidden single in row 1)
    expect(valueChanges.some((c) => c.cell.ref === 'A1' && c.value === 1)).toBe(true);
  });

  it('skips solved cells', () => {
    const cages = [
      { cells: ['A1', 'B1'], operator: '+', value: 3 },
      { cells: ['A2', 'B2'], operator: '+', value: 3 }
    ];
    const puzzle = createTestPuzzle(2, cages, true);
    puzzle.getCell('A1').setValue(1);
    puzzle.getCell('B1').setCandidates([2]);
    puzzle.getCell('A2').setCandidates([1, 2]);
    puzzle.getCell('B2').setCandidates([1, 2]);

    const result = strategy.tryApply(puzzle);
    // B1 has single candidate [2] but that's SingleCandidate, not HiddenSingle
    // HiddenSingle looks for a digit in only one cell in a house
    // Since A1 is solved, in row 1 only B1 is unsolved with candidate 2 → hidden single for 2
    expect(result).not.toBeNull();
  });
});
