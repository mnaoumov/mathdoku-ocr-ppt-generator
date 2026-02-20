import type { Puzzle } from '../Puzzle.ts';
import type {
  Strategy,
  StrategyResult
} from './Strategy.ts';

import { CandidatesChange } from '../cellChanges/CandidatesChange.ts';

export class FillAllCandidatesStrategy implements Strategy {
  public tryApply(puzzle: Puzzle): null | StrategyResult {
    const allValues = Array.from({ length: puzzle.size }, (_, i) => i + 1);
    const changes = puzzle.cells
      .filter((cell) => !cell.isSolved)
      .map((cell) => new CandidatesChange(cell, allValues));
    return changes.length > 0
      ? { changes, note: 'Filling all candidates' }
      : null;
  }
}
