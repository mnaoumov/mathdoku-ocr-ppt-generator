import type {
  Cell,
  CellValueSetter,
  House,
  Puzzle
} from '../Puzzle.ts';
import type {
  Strategy,
  StrategyResult
} from './Strategy.ts';

import { buildAutoEliminateChanges } from '../cageConstraints.ts';

interface HiddenSingleFound {
  readonly cell: Cell;
  readonly house: House;
  readonly value: number;
}

export class HiddenSingleStrategy implements Strategy {
  public tryApply(puzzle: Puzzle): null | StrategyResult {
    const results: HiddenSingleFound[] = [];
    const seen = new Set<Cell>();

    for (const house of puzzle.houses) {
      this.scanHouse(puzzle, house, results, seen);
    }

    if (results.length === 0) {
      return null;
    }

    const valueSetters: CellValueSetter[] = results.map((r) => ({ cell: r.cell, value: r.value }));
    const noteEntries = results.map(
      (r) => `${r.cell.ref} (${r.house.type} ${r.house.label})`
    );
    return {
      changes: buildAutoEliminateChanges(valueSetters),
      note: `Hidden single: ${noteEntries.join(', ')}`
    };
  }

  private scanHouse(
    puzzle: Puzzle,
    house: House,
    results: HiddenSingleFound[],
    seen: Set<Cell>
  ): void {
    for (let hiddenCandidate = 1; hiddenCandidate <= puzzle.size; hiddenCandidate++) {
      let foundCell: Cell | null = null;
      let count = 0;
      for (const cell of house.cells) {
        if (cell.isSolved) {
          continue;
        }
        if (cell.hasCandidate(hiddenCandidate)) {
          foundCell = cell;
          count++;
        }
      }
      if (count === 1 && foundCell && !seen.has(foundCell)) {
        results.push({ cell: foundCell, house, value: hiddenCandidate });
        seen.add(foundCell);
      }
    }
  }
}
