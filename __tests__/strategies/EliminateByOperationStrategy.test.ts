import {
  describe,
  expect,
  it
} from 'vitest';

import { CandidatesStrikethrough } from '../../src/cellChanges/CandidatesStrikethrough.ts';
import { EliminateByOperationStrategy } from '../../src/strategies/EliminateByOperationStrategy.ts';
import { ensureNonNullable } from '../../src/typeGuards.ts';
import {
  createTestPuzzle,
  fillRemainingCells
} from '../puzzleTestHelper.ts';

describe('EliminateByOperationStrategy', () => {
  const strategy = new EliminateByOperationStrategy();

  it('eliminates candidates that do not divide the cage value for multiplication', () => {
    const cages = fillRemainingCells([
      { cells: ['A1', 'B1', 'A2'], operator: 'x', value: 20 }
    ], 6);
    const puzzle = createTestPuzzle({ cages, hasOperators: true, puzzleSize: 6 });
    for (const cell of puzzle.cells) {
      cell.setCandidates([1, 2, 3, 4, 5, 6]);
    }

    const result = strategy.tryApply(puzzle);
    expect(result).not.toBeNull();
    const { changes, note } = ensureNonNullable(result);

    const a1Eliminations = changes.filter(
      (c) => c instanceof CandidatesStrikethrough && c.cell.ref === 'A1'
    );
    expect(a1Eliminations.length).toBeGreaterThan(0);
    expect(note).toContain('Cage operation.');
    expect(note).toContain('@A1:');
  });

  it('deduplicates eliminated values across cells in the same cage', () => {
    // 3-cell multiplication cage: values 3 and 6 don't divide 20
    // All 3 cells have {1,2,3,4,5,6} — should produce ONE "{36} don't divide 20" not repeated per cell
    const cages = fillRemainingCells([
      { cells: ['A1', 'B1', 'A2'], operator: 'x', value: 20 }
    ], 6);
    const puzzle = createTestPuzzle({ cages, hasOperators: true, puzzleSize: 6 });
    for (const cell of puzzle.cells) {
      cell.setCandidates([1, 2, 3, 4, 5, 6]);
    }

    const result = strategy.tryApply(puzzle);
    expect(result).not.toBeNull();
    const { note } = ensureNonNullable(result);

    // Should NOT repeat "3 doesn't divide 20" multiple times
    const matches = note.match(/doesn't divide|don't divide/g);
    expect(ensureNonNullable(matches).length).toBe(1);
  });

  it('uses set notation for multiple eliminated values', () => {
    const cages = fillRemainingCells([
      { cells: ['A1', 'B1', 'A2'], operator: 'x', value: 20 }
    ], 6);
    const puzzle = createTestPuzzle({ cages, hasOperators: true, puzzleSize: 6 });
    for (const cell of puzzle.cells) {
      cell.setCandidates([1, 2, 3, 4, 5, 6]);
    }

    const result = strategy.tryApply(puzzle);
    expect(result).not.toBeNull();
    const { note } = ensureNonNullable(result);

    // Concatenated digits in set notation: {36} don't divide 20
    expect(note).toContain('{36} don\'t divide 20');
  });

  it('eliminates candidates using latin square bounds for addition', () => {
    // Cage 8+ with 3 cells: A1, B1 (both in row 1), A2 (in row 2)
    // All cells have candidates {1..6}
    // For A2 checking V=6: remainder = 8-6 = 2, but A1 and B1 are in the same row,
    // So they need distinct values, min sum = 1+2 = 3 > 2 → 6 eliminated from A2
    const cages = fillRemainingCells([
      { cells: ['A1', 'B1', 'A2'], operator: '+', value: 8 }
    ], 6);
    const puzzle = createTestPuzzle({ cages, hasOperators: true, puzzleSize: 6 });
    for (const cell of puzzle.cells) {
      cell.setCandidates([1, 2, 3, 4, 5, 6]);
    }

    const result = strategy.tryApply(puzzle);
    expect(result).not.toBeNull();
    const { changes } = ensureNonNullable(result);

    // A2 should have 6 eliminated (latin square constraint on row 1 cells)
    const a2Strikethroughs = changes.filter(
      (c) => c instanceof CandidatesStrikethrough && c.cell.ref === 'A2'
    ) as CandidatesStrikethrough[];
    const a2EliminatedValues = a2Strikethroughs.flatMap((c) => [...c.values]);
    expect(a2EliminatedValues).toContain(6);
  });

  it('returns null when no candidates can be eliminated', () => {
    const cages = fillRemainingCells([
      { cells: ['A1', 'B1'], operator: '+', value: 3 }
    ], 2);
    const puzzle = createTestPuzzle({ cages, hasOperators: true, puzzleSize: 2 });
    puzzle.getCell('A1').setCandidates([1, 2]);
    puzzle.getCell('B1').setCandidates([1, 2]);

    expect(strategy.tryApply(puzzle)).toBeNull();
  });
});
