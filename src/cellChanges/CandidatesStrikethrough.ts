import type { Cell } from '../Puzzle.ts';
import type { PuzzleRenderer } from '../Puzzle.ts';

import { CellChange } from './CellChange.ts';

export class CandidatesStrikethrough extends CellChange {
  public constructor(cell: Cell, public readonly values: readonly number[]) {
    super(cell);
  }

  public applyToModel(): void {
    for (const v of this.values) {
      this.cell.removeCandidate(v);
    }
  }

  public renderPending(renderer: PuzzleRenderer): void {
    renderer.renderPendingStrikethrough(this);
  }
}
