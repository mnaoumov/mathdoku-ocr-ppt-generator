import type { Cell } from '../Puzzle.ts';
import type { PuzzleRenderer } from '../Puzzle.ts';

export abstract class CellChange {
  protected constructor(public readonly cell: Cell) {
  }

  public abstract applyToModel(): void;
  public abstract renderPending(renderer: PuzzleRenderer): void;
}
