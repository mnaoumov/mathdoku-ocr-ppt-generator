import type {
  House,
  Puzzle
} from '../Puzzle.ts';
import type {
  ChangeGroup,
  Strategy,
  StrategyResult
} from './Strategy.ts';

import { CandidatesStrikethrough } from '../cellChanges/CandidatesStrikethrough.ts';
import { Cell } from '../Puzzle.ts';
import { ensureNonNullable } from '../typeGuards.ts';
import {
  collectValidTuples,
  getOperatorsForCage
} from './cageTupleAnalysis.ts';

const MINIMUM_CAGE_CELLS = 2;
const MINIMUM_UNSOLVED_CELLS = 2;

export class RequiredCageCandidateStrategy implements Strategy {
  public readonly name = 'Required cage candidate';

  public tryApply(puzzle: Puzzle): null | StrategyResult {
    const allGroups: ChangeGroup[] = [];
    const allNoteEntries: string[] = [];

    for (const cage of puzzle.cages) {
      if (cage.cells.length < MINIMUM_CAGE_CELLS) {
        continue;
      }

      const unsolvedCells = cage.cells.filter((c) => !c.isSolved);
      if (unsolvedCells.length < MINIMUM_UNSOLVED_CELLS) {
        continue;
      }

      const operators = getOperatorsForCage(cage, puzzle.puzzleSize);
      if (operators.length === 0) {
        continue;
      }

      const solvedValues = cage.cells.filter((c) => c.isSolved).map((c) => ensureNonNullable(c.value));
      const validTuples = collectValidTuples(unsolvedCells, cage.value, operators, solvedValues);
      if (validTuples.length === 0) {
        continue;
      }

      const requiredValues = this.findRequiredValues(validTuples, puzzle.puzzleSize);
      const cageSet = new Set(cage.cells);
      const cageRef = `@${cage.topLeft.ref}`;

      for (const value of requiredValues) {
        const candidateCells = unsolvedCells.filter((c) => c.hasCandidate(value));
        if (candidateCells.length === 0) {
          continue;
        }

        const firstCand = ensureNonNullable(candidateCells[0]);

        if (candidateCells.every((c) => c.row === firstCand.row)) {
          this.eliminateFromHouse(firstCand.row, value, cageSet, cageRef, allGroups, allNoteEntries);
        }

        if (candidateCells.every((c) => c.column === firstCand.column)) {
          this.eliminateFromHouse(firstCand.column, value, cageSet, cageRef, allGroups, allNoteEntries);
        }
      }
    }

    if (allGroups.length === 0) {
      return null;
    }

    return {
      changeGroups: allGroups,
      details: allNoteEntries.join('; ')
    };
  }

  private eliminateFromHouse(
    house: House,
    value: number,
    cageSet: ReadonlySet<Cell>,
    cageRef: string,
    allGroups: ChangeGroup[],
    allNoteEntries: string[]
  ): void {
    const changes: CandidatesStrikethrough[] = [];
    for (const cell of house.cells) {
      if (!cageSet.has(cell) && !cell.isSolved && cell.hasCandidate(value)) {
        changes.push(new CandidatesStrikethrough(cell, [value]));
      }
    }
    if (changes.length === 0) {
      return;
    }

    changes.sort((a, b) => Cell.compare(a.cell, b.cell));
    const houseLabel = `${house.type} ${house.label}`;
    const reason = `${cageRef} requires ${String(value)} in ${houseLabel}`;
    allGroups.push({ changes, reason });
    allNoteEntries.push(`${cageRef} -${String(value)} ${houseLabel}`);
  }

  private findRequiredValues(validTuples: readonly number[][], puzzleSize: number): number[] {
    const required: number[] = [];
    for (let value = 1; value <= puzzleSize; value++) {
      if (validTuples.every((tuple) => tuple.includes(value))) {
        required.push(value);
      }
    }
    return required;
  }
}
