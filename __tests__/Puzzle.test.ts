import {
  describe,
  expect,
  it
} from 'vitest';

import { CandidatesChange } from '../src/cellChanges/CandidatesChange.ts';
import { CandidatesStrikethrough } from '../src/cellChanges/CandidatesStrikethrough.ts';
import { CellClearance } from '../src/cellChanges/CellClearance.ts';
import { ValueChange } from '../src/cellChanges/ValueChange.ts';
import { initPuzzleSlides } from '../src/Puzzle.ts';
import { createDefaultStrategies } from '../src/strategies/createDefaultStrategies.ts';
import { ensureNonNullable } from '../src/typeGuards.ts';
import {
  createTestPuzzle,
  TrackingRenderer
} from './puzzleTestHelper.ts';

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
    it('returns two batches: all candidates then cage constraints', () => {
      const puzzle = createTestPuzzle(4, SIZE_4_CAGES, true);
      const batches = puzzle.buildInitChanges();
      expect(batches).toHaveLength(2);
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

describe('tryApplyAutomatedStrategies', () => {
  const SIZE_2_CAGES = [
    { cells: ['A1', 'B1'], operator: '+', value: 3 },
    { cells: ['A2', 'B2'], operator: '+', value: 3 }
  ];

  it('returns false when no strategies apply', () => {
    const puzzle = createTestPuzzle(2, SIZE_2_CAGES, true, undefined, undefined, createDefaultStrategies(2));
    for (const cell of puzzle.cells) {
      cell.setCandidates([1, 2]);
    }
    expect(puzzle.tryApplyAutomatedStrategies()).toBe(false);
  });

  it('applies single candidate strategy and returns true', () => {
    const puzzle = createTestPuzzle(2, SIZE_2_CAGES, true, undefined, undefined, createDefaultStrategies(2));
    puzzle.getCell('A1').setCandidates([1]);
    puzzle.getCell('B1').setCandidates([1, 2]);
    puzzle.getCell('A2').setCandidates([1, 2]);
    puzzle.getCell('B2').setCandidates([1, 2]);

    expect(puzzle.tryApplyAutomatedStrategies()).toBe(true);
    expect(puzzle.getCell('A1').value).toBe(1);
  });

  it('chains multiple strategy steps until no more apply', () => {
    const puzzle = createTestPuzzle(2, SIZE_2_CAGES, true, undefined, undefined, createDefaultStrategies(2));
    // A1 has single candidate → solved → peers lose that candidate → chain
    puzzle.getCell('A1').setCandidates([1]);
    puzzle.getCell('B1').setCandidates([2]);
    puzzle.getCell('A2').setCandidates([1, 2]);
    puzzle.getCell('B2').setCandidates([1, 2]);

    puzzle.tryApplyAutomatedStrategies();
    // All cells should be solved
    for (const cell of puzzle.cells) {
      expect(cell.isSolved).toBe(true);
    }
  });
});

describe('slide notes tracking', () => {
  const SIZE_2_CAGES = [
    { cells: ['A1', 'B1'], operator: '+', value: 3 },
    { cells: ['A2', 'B2'], operator: '+', value: 3 }
  ];

  it('records note text on both pending and committed slides for enter+commit', () => {
    const renderer = new TrackingRenderer();
    const puzzle = createTestPuzzle(2, SIZE_2_CAGES, true, undefined, undefined, undefined, renderer);
    // Init candidates
    const batches = puzzle.buildInitChanges();
    for (const batch of batches) {
      puzzle.applyChanges(batch);
      puzzle.commit();
    }

    puzzle.enter('A1:=1');
    puzzle.commit();

    // Both the pending slide and committed slide should have the note
    const notedSlides = renderer.notesBySlide.filter((n) => n === 'A1:=1');
    expect(notedSlides).toHaveLength(2);
  });

  it('records note text for each strategy step', () => {
    const renderer = new TrackingRenderer();
    const puzzle = createTestPuzzle(2, SIZE_2_CAGES, true, undefined, undefined, createDefaultStrategies(2), renderer);
    puzzle.getCell('A1').setCandidates([1]);
    puzzle.getCell('B1').setCandidates([2]);
    puzzle.getCell('A2').setCandidates([1, 2]);
    puzzle.getCell('B2').setCandidates([1, 2]);

    puzzle.tryApplyAutomatedStrategies();

    // 2 strategy steps (A1+B1 then A2+B2), each producing 2 slides (pending + committed)
    const notedSlides = renderer.notesBySlide.filter((n) => n === 'Applying automated strategies');
    expect(notedSlides).toHaveLength(4);
  });

  it('records different notes for different operations', () => {
    const renderer = new TrackingRenderer();
    const puzzle = createTestPuzzle(2, SIZE_2_CAGES, true, undefined, undefined, undefined, renderer);

    puzzle.enter('A1:12');
    puzzle.commit();

    puzzle.enter('B1:12');
    puzzle.commit();

    const batch1Notes = renderer.notesBySlide.filter((n) => n === 'A1:12');
    const batch2Notes = renderer.notesBySlide.filter((n) => n === 'B1:12');
    expect(batch1Notes).toHaveLength(2);
    expect(batch2Notes).toHaveLength(2);
  });
});

describe('initPuzzleSlides notes', () => {
  const BATCH_1_NOTE = 'Filling all possible cell candidates';
  const BATCH_2_NOTE = 'Filling single cell values and unique cage multisets';

  it('produces 5 slides with notes on slides 1-4 when cage constraints exist', () => {
    const renderer = new TrackingRenderer();
    initPuzzleSlides(renderer, 4, SIZE_4_CAGES, true, '', '', []);

    expect(renderer.slideCount).toBe(5);
    expect(renderer.notesBySlide[0]).toBe(BATCH_1_NOTE);
    expect(renderer.notesBySlide[1]).toBe(BATCH_1_NOTE);
    expect(renderer.notesBySlide[2]).toBe(BATCH_2_NOTE);
    expect(renderer.notesBySlide[3]).toBe(BATCH_2_NOTE);
    expect(renderer.notesBySlide[4]).toBeUndefined();
  });

  it('produces 3 slides with notes on slides 1-2 when no cage constraints exist', () => {
    const cages = [
      { cells: ['A1', 'B1'] },
      { cells: ['A2', 'B2'] }
    ];
    const renderer = new TrackingRenderer();
    initPuzzleSlides(renderer, 2, cages, false, '', '', []);

    expect(renderer.slideCount).toBe(3);
    expect(renderer.notesBySlide[0]).toBe(BATCH_1_NOTE);
    expect(renderer.notesBySlide[1]).toBe(BATCH_1_NOTE);
    expect(renderer.notesBySlide[2]).toBeUndefined();
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
