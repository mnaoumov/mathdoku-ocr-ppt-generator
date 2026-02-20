import type { CellValueSetter } from '../Puzzle.ts';
import type { Puzzle } from '../Puzzle.ts';
import type {
  Strategy,
  StrategyResult
} from './Strategy.ts';

import { buildAutoEliminateChanges } from '../cageConstraints.ts';
import { ensureNonNullable } from '../typeGuards.ts';

export class SingleCellCageStrategy implements Strategy {
  public tryApply(puzzle: Puzzle): null | StrategyResult {
    const valueSetters: CellValueSetter[] = [];
    for (const cage of puzzle.cages) {
      if (cage.cells.length !== 1) {
        continue;
      }
      const cellValue = cage.value ?? (cage.label ? parseInt(cage.label, 10) : undefined);
      if (cellValue === undefined || isNaN(cellValue)) {
        continue;
      }
      valueSetters.push({ cell: ensureNonNullable(cage.cells[0]), value: cellValue });
    }

    if (valueSetters.length === 0) {
      return null;
    }

    const changes = buildAutoEliminateChanges(valueSetters);
    const cellRefs = valueSetters.map((s) => s.cell.ref).join(', ');
    return { changes, note: `Single cell: ${cellRefs}` };
  }
}
