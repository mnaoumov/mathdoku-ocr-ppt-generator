import type { Cell } from '../Puzzle.ts';
import type { PuzzleRenderer } from '../Puzzle.ts';

import { CellChange } from './CellChange.ts';

export class CellClearance extends CellChange {
  public constructor(cell: Cell) {
    super(cell);
  }

  public applyToModel(): void {
    this.cell.clearValue();
    this.cell.clearCandidates();
  }

  public renderPending(renderer: PuzzleRenderer): void {
    renderer.renderPendingClearance(this);
  }
}
