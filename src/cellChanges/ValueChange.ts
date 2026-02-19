import type { Cell } from '../Puzzle.ts';
import type { PuzzleRenderer } from '../Puzzle.ts';

import { CellChange } from './CellChange.ts';

export class ValueChange extends CellChange {
  public constructor(cell: Cell, public readonly value: number) {
    super(cell);
  }

  public applyToModel(): void {
    this.cell.setValue(this.value);
    this.cell.clearCandidates();
  }

  public renderPending(renderer: PuzzleRenderer): void {
    renderer.renderPendingValue(this);
  }
}
