import {
  describe,
  expect,
  it
} from 'vitest';

import { CandidatesChange } from '../src/cellChanges/CandidatesChange.ts';
import { CandidatesStrikethrough } from '../src/cellChanges/CandidatesStrikethrough.ts';
import { CellClearance } from '../src/cellChanges/CellClearance.ts';
import { ValueChange } from '../src/cellChanges/ValueChange.ts';
import { ensureNonNullable } from '../src/typeGuards.ts';
import { createTestPuzzle } from './puzzleTestHelper.ts';

const SIZE_4_CAGES = [
  { cells: ['A1', 'A2'], operator: '+', value: 3 },
  { cells: ['B1', 'B2'], operator: '-', value: 1 },
  { cells: ['A3', 'B3'], operator: '+', value: 7 },
  { cells: ['A4', 'B4'], operator: 'x', value: 12 },
  { cells: ['C1', 'C2'], operator: '+', value: 5 },
  { cells: ['D1', 'D2'], operator: '-', value: 1 },
  { cells: ['C3', 'D3'], operator: '+', value: 5 },
  { cells: ['C4', 'D4'], operator: 'x', value: 2 }
];

describe('Puzzle', () => {
  describe('constructor', () => {
    it('creates cells for all grid positions', () => {
      const puzzle = createTestPuzzle(4, SIZE_4_CAGES, true);
      expect(puzzle.cells).toHaveLength(16);
    });

    it('creates correct number of rows and columns', () => {
      const puzzle = createTestPuzzle(4, SIZE_4_CAGES, true);
      expect(puzzle.rows).toHaveLength(4);
      expect(puzzle.columns).toHaveLength(4);
    });

    it('creates houses combining rows and columns', () => {
      const puzzle = createTestPuzzle(4, SIZE_4_CAGES, true);
      expect(puzzle.houses).toHaveLength(8);
    });

    it('throws if cell is not in any cage', () => {
      const incompleteCages = [
        { cells: ['A1'], value: 1 }
      ];
      expect(() => createTestPuzzle(2, incompleteCages, true)).toThrow('not found in any cage');
    });
  });

  describe('getCell', () => {
    it('retrieves cell by ref string', () => {
      const puzzle = createTestPuzzle(4, SIZE_4_CAGES, true);
      const cell = puzzle.getCell('B3');
      expect(cell.ref).toBe('B3');
    });

    it('retrieves cell by row and column ids', () => {
      const puzzle = createTestPuzzle(4, SIZE_4_CAGES, true);
      const cell = puzzle.getCell(2, 3);
      expect(cell.ref).toBe('C2');
    });
  });

  describe('buildInitChanges', () => {
    it('returns at least one batch of candidate changes', () => {
      const puzzle = createTestPuzzle(4, SIZE_4_CAGES, true);
      const batches = puzzle.buildInitChanges();
      expect(batches.length).toBeGreaterThanOrEqual(1);
      // First batch: all cells get all candidates
      const firstBatch = ensureNonNullable(batches[0]);
      expect(firstBatch).toHaveLength(16);
      for (const change of firstBatch) {
        expect(change).toBeInstanceOf(CandidatesChange);
      }
    });
  });

  describe('enter', () => {
    it('creates value change for =N', () => {
      const puzzle = createTestPuzzle(4, SIZE_4_CAGES, true);
      // First init so cells have candidates
      const batches = puzzle.buildInitChanges();
      for (const batch of batches) {
        puzzle.applyChanges(batch);
        puzzle.commit();
      }
      puzzle.enter('A1:=3');
      // Commit so changes are applied
      puzzle.commit();
      expect(puzzle.getCell('A1').value).toBe(3);
    });

    it('creates candidates change for digits', () => {
      const puzzle = createTestPuzzle(4, SIZE_4_CAGES, true);
      puzzle.enter('A1:123');
      puzzle.commit();
      expect(puzzle.getCell('A1').getCandidates()).toEqual([1, 2, 3]);
    });

    it('creates strikethrough change for -digits', () => {
      const puzzle = createTestPuzzle(4, SIZE_4_CAGES, true);
      // Set up candidates first
      puzzle.enter('A1:1234');
      puzzle.commit();
      puzzle.enter('A1:-34');
      puzzle.commit();
      expect(puzzle.getCell('A1').getCandidates()).toEqual([1, 2]);
    });

    it('creates clearance change for x', () => {
      const puzzle = createTestPuzzle(4, SIZE_4_CAGES, true);
      puzzle.enter('A1:=3');
      puzzle.commit();
      puzzle.enter('A1:x');
      puzzle.commit();
      expect(puzzle.getCell('A1').value).toBeNull();
      expect(puzzle.getCell('A1').getCandidates()).toEqual([]);
    });

    it('throws for empty input', () => {
      const puzzle = createTestPuzzle(4, SIZE_4_CAGES, true);
      expect(() => {
        puzzle.enter('');
      }).toThrow('No commands specified');
    });

    it('throws for invalid format', () => {
      const puzzle = createTestPuzzle(4, SIZE_4_CAGES, true);
      expect(() => {
        puzzle.enter('invalid');
      }).toThrow('Invalid format');
    });
  });
});

describe('Cell', () => {
  it('tracks candidates correctly', () => {
    const puzzle = createTestPuzzle(4, SIZE_4_CAGES, true);
    const cell = puzzle.getCell('A1');
    cell.setCandidates([1, 2, 3]);
    expect(cell.getCandidates()).toEqual([1, 2, 3]);
    expect(cell.hasCandidate(2)).toBe(true);
    expect(cell.hasCandidate(4)).toBe(false);
    cell.removeCandidate(2);
    expect(cell.getCandidates()).toEqual([1, 3]);
  });

  it('tracks value correctly', () => {
    const puzzle = createTestPuzzle(4, SIZE_4_CAGES, true);
    const cell = puzzle.getCell('A1');
    expect(cell.isSolved).toBe(false);
    expect(cell.value).toBeNull();
    cell.setValue(3);
    expect(cell.isSolved).toBe(true);
    expect(cell.value).toBe(3);
    cell.clearValue();
    expect(cell.isSolved).toBe(false);
  });

  it('has correct peers (row + column, excluding self)', () => {
    const puzzle = createTestPuzzle(4, SIZE_4_CAGES, true);
    const cell = puzzle.getCell('B2');
    // Row 2: A2, C2, D2 (3 peers) + Column B: B1, B3, B4 (3 peers) = 6
    expect(cell.peers).toHaveLength(6);
  });
});

describe('CellChange subclasses', () => {
  it('CandidatesChange sets candidates and clears value', () => {
    const puzzle = createTestPuzzle(4, SIZE_4_CAGES, true);
    const cell = puzzle.getCell('A1');
    cell.setValue(5);
    const change = new CandidatesChange(cell, [1, 2, 3]);
    change.applyToModel();
    expect(cell.getCandidates()).toEqual([1, 2, 3]);
    expect(cell.value).toBeNull();
  });

  it('CandidatesStrikethrough removes specific candidates', () => {
    const puzzle = createTestPuzzle(4, SIZE_4_CAGES, true);
    const cell = puzzle.getCell('A1');
    cell.setCandidates([1, 2, 3, 4]);
    const change = new CandidatesStrikethrough(cell, [2, 4]);
    change.applyToModel();
    expect(cell.getCandidates()).toEqual([1, 3]);
  });

  it('CellClearance clears value and candidates', () => {
    const puzzle = createTestPuzzle(4, SIZE_4_CAGES, true);
    const cell = puzzle.getCell('A1');
    cell.setValue(5);
    cell.setCandidates([1, 2]);
    const change = new CellClearance(cell);
    change.applyToModel();
    expect(cell.value).toBeNull();
    expect(cell.getCandidates()).toEqual([]);
  });

  it('ValueChange sets value and clears candidates', () => {
    const puzzle = createTestPuzzle(4, SIZE_4_CAGES, true);
    const cell = puzzle.getCell('A1');
    cell.setCandidates([1, 2, 3]);
    const change = new ValueChange(cell, 5);
    change.applyToModel();
    expect(cell.value).toBe(5);
    expect(cell.getCandidates()).toEqual([]);
  });
});
