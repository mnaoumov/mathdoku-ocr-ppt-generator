import type {
  CageRaw,
  PuzzleRenderer
} from '../src/Puzzle.ts';
import type { Strategy } from '../src/strategies/Strategy.ts';

import { Puzzle } from '../src/Puzzle.ts';

export class TrackingRenderer implements PuzzleRenderer {
  public isLastSlide = true;
  public readonly notesBySlide: string[] = [];
  public get slideCount(): number {
    return this.currentSlide + 1;
  }

  private currentSlide = 0;

  private noteText = '';

  public beginPendingRender(): void {
    this.recordNote();
  }

  public ensureLastSlide(): boolean {
    return this.isLastSlide;
  }

  public renderCommittedChanges(): void {
    this.recordNote();
  }

  public renderPendingCandidates(): void {
    // No-op
  }

  public renderPendingClearance(): void {
    // No-op
  }

  public renderPendingStrikethrough(): void {
    // No-op
  }

  public renderPendingValue(): void {
    // No-op
  }

  public setNoteText(text: string): void {
    this.noteText = text;
  }

  private recordNote(): void {
    if (this.noteText) {
      this.notesBySlide[this.currentSlide] = this.noteText;
    }
    this.currentSlide++;
  }
}

export function createTestPuzzle(
  size: number,
  cages: readonly CageRaw[],
  hasOperators: boolean,
  initialValues?: Map<string, number>,
  initialCandidates?: Map<string, Set<number>>,
  strategies?: readonly Strategy[],
  renderer: PuzzleRenderer = new TrackingRenderer()
): Puzzle {
  return new Puzzle(
    renderer,
    size,
    cages,
    hasOperators,
    'Test Puzzle',
    'test',
    strategies ?? [],
    initialValues,
    initialCandidates
  );
}
