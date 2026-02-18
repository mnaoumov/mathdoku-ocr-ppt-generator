/**
 * HiddenSingleStrategy.ts -- value appears in only 1 cell in a house -> set value.
 */

class HiddenSingleStrategy implements Strategy {
  public tryApply(puzzle: Puzzle): CellChange[] | null {
    const results: CellValueSetter[] = [];
    const seen = new Set<Cell>();

    for (const house of puzzle.houses) {
      this.scanHouse(puzzle, house, results, seen);
    }

    if (results.length === 0) {
      return null;
    }
    return buildAutoEliminateChanges(results);
  }

  private scanHouse(
    puzzle: Puzzle,
    house: House,
    results: CellValueSetter[],
    seen: Set<Cell>
  ): void {
    for (let hiddenCandidate = 1; hiddenCandidate <= puzzle.size; hiddenCandidate++) {
      let foundCell: Cell | null = null;
      let count = 0;
      for (const cell of house.cells) {
        if (cell.isSolved) {
          continue;
        }
        if (cell.hasCandidate(hiddenCandidate)) {
          foundCell = cell;
          count++;
        }
      }
      if (count === 1 && foundCell && !seen.has(foundCell)) {
        results.push({ cell: foundCell, value: hiddenCandidate });
        seen.add(foundCell);
      }
    }
  }
}
