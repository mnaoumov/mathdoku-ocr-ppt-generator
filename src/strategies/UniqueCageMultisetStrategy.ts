import type { CellChange } from '../cellChanges/CellChange.ts';
import type {
  CellValueSetter,
  Puzzle
} from '../Puzzle.ts';
import type {
  Strategy,
  StrategyResult
} from './Strategy.ts';

import {
  applyCageConstraint,
  buildAutoEliminateChanges
} from '../cageConstraints.ts';
import { ValueChange } from '../cellChanges/ValueChange.ts';

export class UniqueCageMultisetStrategy implements Strategy {
  public tryApply(puzzle: Puzzle): null | StrategyResult {
    const affectedCageRefs: string[] = [];

    const allValueSetters: CellValueSetter[] = [];
    const allCandidateChanges: CellChange[] = [];

    for (const cage of puzzle.cages) {
      if (cage.cells.length <= 1) {
        continue;
      }

      const valueSetters: CellValueSetter[] = [];
      const candidateChanges: CellChange[] = [];
      applyCageConstraint({ cage, hasOperators: puzzle.hasOperators, puzzleSize: puzzle.puzzleSize }, valueSetters, candidateChanges);

      if (valueSetters.length > 0 || candidateChanges.length > 0) {
        affectedCageRefs.push(`@${cage.topLeft.ref}`);
        allValueSetters.push(...valueSetters);
        allCandidateChanges.push(...candidateChanges);
      }
    }

    if (allValueSetters.length === 0 && allCandidateChanges.length === 0) {
      return null;
    }

    const valuedCells = new Set(allValueSetters.map((s) => s.cell));
    const autoChanges = buildAutoEliminateChanges(allValueSetters);
    const valueChanges = autoChanges.filter((c) => c instanceof ValueChange);
    const strikethroughChanges = autoChanges.filter(
      (c) => !(c instanceof ValueChange) && !valuedCells.has(c.cell)
    );
    const changes = [...valueChanges, ...allCandidateChanges, ...strikethroughChanges];

    return {
      changes,
      note: `Unique cage multiset: ${affectedCageRefs.join(', ')}`
    };
  }
}
