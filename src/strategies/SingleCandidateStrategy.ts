import type {
  CellValueSetter,
  Puzzle
} from '../Puzzle.ts';
import type {
  Strategy,
  StrategyResult
} from './Strategy.ts';

import { buildAutoEliminateChanges } from '../cageConstraints.ts';
import { ensureNonNullable } from '../typeGuards.ts';

export class SingleCandidateStrategy implements Strategy {
  public tryApply(puzzle: Puzzle): null | StrategyResult {
    const results: CellValueSetter[] = [];
    for (const cell of puzzle.cells) {
      if (cell.isSolved) {
        continue;
      }
      const cands = cell.getCandidates();
      if (cands.length === 1) {
        results.push({ cell, value: ensureNonNullable(cands[0]) });
      }
    }
    if (results.length === 0) {
      return null;
    }
    const cellRefs = results.map((r) => r.cell.ref).join(', ');
    return {
      changes: buildAutoEliminateChanges(results),
      note: `Single candidate: ${cellRefs}`
    };
  }
}
