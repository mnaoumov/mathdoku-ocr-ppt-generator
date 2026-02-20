import type { Puzzle } from '../Puzzle.ts';
import type {
  Strategy,
  StrategyResult
} from './Strategy.ts';

import { CandidatesStrikethrough } from '../cellChanges/CandidatesStrikethrough.ts';
import { generateSubsets } from '../combinatorics.ts';
import { Cell } from '../Puzzle.ts';

const PAIR_SIZE = 2;
const TRIPLET_SIZE = 3;
const QUAD_SIZE = 4;

const NAKED_SET_NAMES: Record<number, string> = {
  [PAIR_SIZE]: 'Naked pair',
  [QUAD_SIZE]: 'Naked quad',
  [TRIPLET_SIZE]: 'Naked triplet'
};

export class NakedSetStrategy implements Strategy {
  public constructor(private readonly subsetSize: number) {
  }

  public tryApply(puzzle: Puzzle): null | StrategyResult {
    const allEliminations = new Map<Cell, Set<number>>();
    const foundSubsets: Cell[][] = [];

    for (const house of puzzle.houses) {
      const filtered = house.cells.filter((cell) => !cell.isSolved && cell.candidateCount > 0);
      if (filtered.length > this.subsetSize) {
        this.scanHouseForNakedSet(filtered, allEliminations, foundSubsets);
      }
    }

    if (allEliminations.size === 0) {
      return null;
    }

    const changes = [...allEliminations.entries()]
      .map(([cell, values]) => new CandidatesStrikethrough(cell, [...values].sort((a, b) => a - b)))
      .sort((a, b) => Cell.compare(a.cell, b.cell));

    const name = NAKED_SET_NAMES[this.subsetSize] ?? `Naked set (${String(this.subsetSize)})`;
    const subsetDescriptions = foundSubsets.map(
      (subset) => `(${subset.map((c) => c.ref).join(' ')})`
    );
    return {
      changes,
      note: `${name}: ${subsetDescriptions.join(', ')}`
    };
  }

  private scanHouseForNakedSet(
    cells: readonly Cell[],
    allEliminations: Map<Cell, Set<number>>,
    foundSubsets: Cell[][]
  ): void {
    for (const subset of generateSubsets(cells, this.subsetSize)) {
      const union = new Set<number>();
      for (const cell of subset) {
        for (const v of cell.getCandidates()) {
          union.add(v);
        }
      }

      if (union.size !== this.subsetSize) {
        continue;
      }

      const subsetSet = new Set(subset);
      let hasEliminations = false;
      for (const cell of cells) {
        if (subsetSet.has(cell)) {
          continue;
        }
        const toEliminate = cell.getCandidates().filter((v) => union.has(v));
        if (toEliminate.length > 0) {
          hasEliminations = true;
          let existing = allEliminations.get(cell);
          if (!existing) {
            existing = new Set<number>();
            allEliminations.set(cell, existing);
          }
          for (const v of toEliminate) {
            existing.add(v);
          }
        }
      }

      if (hasEliminations) {
        foundSubsets.push([...subset]);
      }
    }
  }
}
