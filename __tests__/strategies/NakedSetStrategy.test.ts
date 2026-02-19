import {
  describe,
  expect,
  it
} from 'vitest';

import { CandidatesStrikethrough } from '../../src/cellChanges/CandidatesStrikethrough.ts';
import { NakedSetStrategy } from '../../src/strategies/NakedSetStrategy.ts';
import { ensureNonNullable } from '../../src/typeGuards.ts';
import { createTestPuzzle } from '../puzzleTestHelper.ts';

describe('NakedSetStrategy', () => {
  it('returns null when no naked pairs exist', () => {
    const cages = [
      { cells: ['A1', 'A2', 'A3'], operator: '+', value: 6 },
      { cells: ['B1', 'B2', 'B3'], operator: '+', value: 6 },
      { cells: ['C1', 'C2', 'C3'], operator: '+', value: 6 }
    ];
    const puzzle = createTestPuzzle(3, cages, true);
    // All cells have all 3 candidates — no naked pair possible
    for (const cell of puzzle.cells) {
      cell.setCandidates([1, 2, 3]);
    }
    const strategy = new NakedSetStrategy(2);
    expect(strategy.tryApply(puzzle)).toBeNull();
  });

  it('finds naked pair and eliminates from other cells in the house', () => {
    const cages = [
      { cells: ['A1', 'A2', 'A3'], operator: '+', value: 6 },
      { cells: ['B1', 'B2', 'B3'], operator: '+', value: 6 },
      { cells: ['C1', 'C2', 'C3'], operator: '+', value: 6 }
    ];
    const puzzle = createTestPuzzle(3, cages, true);
    // Row 1: A1={1,2}, B1={1,2}, C1={1,2,3}
    // Naked pair {1,2} in A1,B1 → eliminate 1,2 from C1
    puzzle.getCell('A1').setCandidates([1, 2]);
    puzzle.getCell('B1').setCandidates([1, 2]);
    puzzle.getCell('C1').setCandidates([1, 2, 3]);
    // Other cells
    puzzle.getCell('A2').setCandidates([1, 2, 3]);
    puzzle.getCell('B2').setCandidates([1, 2, 3]);
    puzzle.getCell('C2').setCandidates([1, 2, 3]);
    puzzle.getCell('A3').setCandidates([1, 2, 3]);
    puzzle.getCell('B3').setCandidates([1, 2, 3]);
    puzzle.getCell('C3').setCandidates([1, 2, 3]);

    const strategy = new NakedSetStrategy(2);
    const result = strategy.tryApply(puzzle);
    expect(result).not.toBeNull();

    // C1 should have strikethrough of [1, 2]
    const c1Changes = ensureNonNullable(result).filter(
      (c) => c instanceof CandidatesStrikethrough && c.cell.ref === 'C1'
    ) as CandidatesStrikethrough[];
    expect(c1Changes.length).toBeGreaterThan(0);
    const eliminatedValues = c1Changes.flatMap((c) => [...c.values]);
    expect(eliminatedValues).toContain(1);
    expect(eliminatedValues).toContain(2);
  });

  it('skips solved cells', () => {
    const cages = [
      { cells: ['A1', 'A2', 'A3'], operator: '+', value: 6 },
      { cells: ['B1', 'B2', 'B3'], operator: '+', value: 6 },
      { cells: ['C1', 'C2', 'C3'], operator: '+', value: 6 }
    ];
    const puzzle = createTestPuzzle(3, cages, true);
    puzzle.getCell('A1').setValue(1);
    puzzle.getCell('B1').setCandidates([2, 3]);
    puzzle.getCell('C1').setCandidates([2, 3]);
    puzzle.getCell('A2').setCandidates([1, 2, 3]);
    puzzle.getCell('B2').setCandidates([1, 2, 3]);
    puzzle.getCell('C2').setCandidates([1, 2, 3]);
    puzzle.getCell('A3').setCandidates([1, 2, 3]);
    puzzle.getCell('B3').setCandidates([1, 2, 3]);
    puzzle.getCell('C3').setCandidates([1, 2, 3]);

    const strategy = new NakedSetStrategy(2);
    const result = strategy.tryApply(puzzle);
    // B1 and C1 form a naked pair {2,3}, but there's no other unsolved cell in row 1 to eliminate from
    // (A1 is solved), so no eliminations in row 1
    // But in columns B and C, the naked pair might cause eliminations
    if (result) {
      // All eliminations should be for unsolved cells only
      for (const change of result) {
        expect(change.cell.isSolved).toBe(false);
      }
    }
  });
});
