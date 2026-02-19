import type { Cell } from '../Puzzle.ts';
import type { PuzzleRenderer } from '../Puzzle.ts';

import { CellChange } from './CellChange.ts';

export class CandidatesChange extends CellChange {
  public constructor(cell: Cell, public readonly values: readonly number[]) {
    super(cell);
  }

  public applyToModel(): void {
    this.cell.setCandidates(this.values);
    this.cell.clearValue();
  }

  public renderPending(renderer: PuzzleRenderer): void {
    renderer.renderPendingCandidates(this);
  }
}
