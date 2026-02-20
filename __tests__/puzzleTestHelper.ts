import type {
  CageRaw,
  PuzzleRenderer
} from '../src/Puzzle.ts';
import type { Strategy } from '../src/strategies/Strategy.ts';

import { Puzzle } from '../src/Puzzle.ts';

export interface CreateTestPuzzleOptions {
  readonly cages: readonly CageRaw[];
  readonly hasOperators: boolean;
  readonly initialCandidates?: Map<string, Set<number>>;
  readonly initialValues?: Map<string, number>;
  readonly renderer?: PuzzleRenderer;
  readonly size: number;
  readonly strategies?: readonly Strategy[];
}

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

export function createTestPuzzle(options: CreateTestPuzzleOptions): Puzzle {
  return new Puzzle({
    cages: options.cages,
    hasOperators: options.hasOperators,
    meta: 'test',
    renderer: options.renderer ?? new TrackingRenderer(),
    size: options.size,
    strategies: options.strategies ?? [],
    title: 'Test Puzzle',
    ...options.initialCandidates !== undefined && { initialCandidates: options.initialCandidates },
    ...options.initialValues !== undefined && { initialValues: options.initialValues }
  });
}
