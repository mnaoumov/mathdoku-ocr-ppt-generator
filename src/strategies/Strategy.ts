import type { CellChange } from '../cellChanges/CellChange.ts';
import type { Puzzle } from '../Puzzle.ts';

export interface Strategy {
  tryApply(puzzle: Puzzle): null | StrategyResult;
}

export interface StrategyResult {
  readonly changes: readonly CellChange[];
  readonly note: string;
}
