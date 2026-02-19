import {
  describe,
  expect,
  it
} from 'vitest';

import { ValueChange } from '../../src/cellChanges/ValueChange.ts';
import { SingleCandidateStrategy } from '../../src/strategies/SingleCandidateStrategy.ts';
import { ensureNonNullable } from '../../src/typeGuards.ts';
import { createTestPuzzle } from '../puzzleTestHelper.ts';

describe('SingleCandidateStrategy', () => {
  const strategy = new SingleCandidateStrategy();

  it('returns null when no cells have exactly one candidate', () => {
    const cages = [
      { cells: ['A1', 'B1'], operator: '+', value: 3 },
      { cells: ['A2', 'B2'], operator: '+', value: 3 }
    ];
    const puzzle = createTestPuzzle(2, cages, true);
    // Give all cells 2 candidates
    for (const cell of puzzle.cells) {
      cell.setCandidates([1, 2]);
    }
    expect(strategy.tryApply(puzzle)).toBeNull();
  });

  it('finds cells with single candidate and sets their value', () => {
    const cages = [
      { cells: ['A1', 'B1'], operator: '+', value: 3 },
      { cells: ['A2', 'B2'], operator: '+', value: 3 }
    ];
    const puzzle = createTestPuzzle(2, cages, true);
    puzzle.getCell('A1').setCandidates([1]);
    puzzle.getCell('B1').setCandidates([1, 2]);
    puzzle.getCell('A2').setCandidates([1, 2]);
    puzzle.getCell('B2').setCandidates([1, 2]);

    const result = strategy.tryApply(puzzle);
    expect(result).not.toBeNull();
    expect(ensureNonNullable(result).some((c) => c instanceof ValueChange && c.cell.ref === 'A1' && c.value === 1)).toBe(true);
  });

  it('skips already solved cells', () => {
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
    expect(result).not.toBeNull();
    // Only B1 should be found (A1 is already solved)
    const valueChanges = ensureNonNullable(result).filter((c) => c instanceof ValueChange);
    expect(valueChanges).toHaveLength(1);
    expect(ensureNonNullable(valueChanges[0]).cell.ref).toBe('B1');
  });
});
