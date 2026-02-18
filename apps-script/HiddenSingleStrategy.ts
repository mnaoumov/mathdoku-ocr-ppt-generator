/**
 * HiddenSingleStrategy.ts -- digit appears in only 1 cell in a house -> set value.
 */

class HiddenSingleStrategy implements Strategy {
  public tryApply(puzzle: Puzzle): null | SolveChange[] {
    const results: ValueSet[] = [];
    const seen = new Set<string>();

    for (const house of puzzle.houses) {
      this.scanHouse(puzzle, house, results, seen);
    }

    if (results.length === 0) {
      return null;
    }
    return buildAutoEliminateChanges(results, puzzle.size);
  }

  private scanHouse(
    puzzle: Puzzle,
    house: readonly string[],
    results: ValueSet[],
    seen: Set<string>
  ): void {
    for (let d = 1; d <= puzzle.size; d++) {
      const digit = String(d);
      let foundRef: null | string = null;
      let count = 0;
      for (const ref of house) {
        if (puzzle.getCellValue(ref)) {
          continue;
        }
        if (puzzle.getCellCandidates(ref).includes(digit)) {
          foundRef = ref;
          count++;
        }
      }
      if (count === 1 && foundRef && !seen.has(foundRef)) {
        results.push({ cellRef: foundRef, digit });
        seen.add(foundRef);
      }
    }
  }
}
