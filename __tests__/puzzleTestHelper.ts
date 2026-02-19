import type {
  CageRaw,
  PuzzleRenderer
} from '../src/Puzzle.ts';

import { Puzzle } from '../src/Puzzle.ts';

class NoOpRenderer implements PuzzleRenderer {
  public beginPendingRender(): void {
    // No-op
  }

  public renderCommittedChanges(): void {
    // No-op
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
}

export function createTestPuzzle(
  size: number,
  cages: readonly CageRaw[],
  hasOperators: boolean,
  initialValues?: Map<string, number>,
  initialCandidates?: Map<string, Set<number>>
): Puzzle {
  return new Puzzle(
    new NoOpRenderer(),
    size,
    cages,
    hasOperators,
    'Test Puzzle',
    'test',
    [],
    initialValues,
    initialCandidates
  );
}
