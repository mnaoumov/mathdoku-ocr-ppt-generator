/**
 * SingleCandidateStrategy.ts -- cells with exactly 1 candidate -> set value.
 */

class SingleCandidateStrategy implements Strategy {
  public tryApply(puzzle: Puzzle): null | SolveChange[] {
    const results: ValueSet[] = [];
    for (let r = 0; r < puzzle.size; r++) {
      for (let c = 0; c < puzzle.size; c++) {
        const ref = cellRefA1(r, c);
        if (puzzle.getCellValue(ref)) {
          continue;
        }
        const cands = puzzle.getCellCandidates(ref);
        if (cands.length === 1) {
          results.push({ cellRef: ref, digit: ensureNonNullable(cands[0]) });
        }
      }
    }
    if (results.length === 0) {
      return null;
    }
    return buildAutoEliminateChanges(results, puzzle.size);
  }
}
