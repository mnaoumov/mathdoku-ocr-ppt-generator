/**
 * NakedSetStrategy.ts -- naked set of size k: k cells whose combined candidates = k values.
 */

class NakedSetStrategy implements Strategy {
  public constructor(private readonly subsetSize: number) {
  }

  public tryApply(puzzle: Puzzle): CellChange[] | null {
    const allEliminations = new Map<Cell, Set<number>>();
    for (const house of puzzle.houses) {
      const filtered = house.cells.filter((cell) => !cell.isSolved && cell.candidateCount > 0);
      if (filtered.length > this.subsetSize) {
        this.scanHouseForNakedSet(filtered, allEliminations);
      }
    }

    const result: CellChange[] = [];
    for (const [cell, values] of allEliminations) {
      result.push(new CandidatesStrikethrough(cell, [...values].sort((a, b) => a - b)));
    }
    result.sort((a, b) => a.cell.ref.localeCompare(b.cell.ref));
    return result.length > 0 ? result : null;
  }

  private scanHouseForNakedSet(
    cells: readonly Cell[],
    allEliminations: Map<Cell, Set<number>>
  ): void {
    for (const subset of generateSubsets(cells, this.subsetSize)) {
      const union = new Set<number>();
      for (const cell of subset) {
        for (const v of cell.getCandidates()) {
          union.add(v);
        }
      }

      if (union.size !== this.subsetSize) {
        continue;
      }

      const subsetSet = new Set(subset);
      for (const cell of cells) {
        if (subsetSet.has(cell)) {
          continue;
        }
        const toEliminate = cell.getCandidates().filter((v) => union.has(v));
        if (toEliminate.length > 0) {
          let existing = allEliminations.get(cell);
          if (!existing) {
            existing = new Set<number>();
            allEliminations.set(cell, existing);
          }
          for (const v of toEliminate) {
            existing.add(v);
          }
        }
      }
    }
  }
}
