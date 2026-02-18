/**
 * NakedSetStrategy.ts -- naked set of size k: k cells whose combined candidates = k digits.
 */

class NakedSetStrategy implements Strategy {
  private readonly k: number;

  public constructor(k: number) {
    this.k = k;
  }

  public tryApply(puzzle: Puzzle): null | SolveChange[] {
    const candidateMap: Record<string, string[]> = {};
    for (let r = 0; r < puzzle.size; r++) {
      for (let c = 0; c < puzzle.size; c++) {
        const ref = cellRefA1(r, c);
        if (puzzle.getCellValue(ref)) {
          continue;
        }
        const cands = puzzle.getCellCandidates(ref);
        if (cands.length > 0) {
          candidateMap[ref] = cands;
        }
      }
    }

    const allEliminations = new Map<string, Set<string>>();
    for (const house of puzzle.houses) {
      const filtered = house.filter((ref) => Boolean(candidateMap[ref]));
      if (filtered.length > this.k) {
        scanHouseForNakedSet(filtered, this.k, candidateMap, allEliminations);
      }
    }

    const result: SolveChange[] = [];
    for (const [ref, digits] of allEliminations) {
      result.push({ cellRef: ref, digits: [...digits].sort().join(''), type: 'strike' });
    }
    result.sort((a, b) => a.cellRef.localeCompare(b.cellRef));
    return result.length > 0 ? result : null;
  }
}

function scanHouseForNakedSet(
  house: readonly string[],
  k: number,
  candidateMap: Record<string, string[]>,
  allEliminations: Map<string, Set<string>>
): void {
  for (const subset of generateSubsets(house, k)) {
    const union = new Set<string>();
    for (const ref of subset) {
      for (const d of ensureNonNullable(candidateMap[ref])) {
        union.add(d);
      }
    }

    if (union.size !== k) {
      continue;
    }

    const subsetSet = new Set(subset);
    for (const ref of house) {
      if (subsetSet.has(ref)) {
        continue;
      }
      const cands = ensureNonNullable(candidateMap[ref]);
      const toEliminate = cands.filter((d) => union.has(d));
      if (toEliminate.length > 0) {
        let existing = allEliminations.get(ref);
        if (!existing) {
          existing = new Set<string>();
          allEliminations.set(ref, existing);
        }
        for (const d of toEliminate) {
          existing.add(d);
        }
      }
    }
  }
}
