import type {
  Cage,
  Puzzle
} from '../Puzzle.ts';
import type {
  Strategy,
  StrategyResult
} from './Strategy.ts';

import { CandidatesStrikethrough } from '../cellChanges/CandidatesStrikethrough.ts';

const BINARY_CELL_COUNT = 2;

interface EliminationEntry {
  readonly cageRef: string;
  readonly reasons: string[];
}

export class EliminateByOperationStrategy implements Strategy {
  public tryApply(puzzle: Puzzle): null | StrategyResult {
    const changes: CandidatesStrikethrough[] = [];
    const entries: EliminationEntry[] = [];

    for (const cage of puzzle.cages) {
      if (cage.cells.length <= 1) {
        continue;
      }
      const cageValue = cage.value ?? (cage.label ? parseInt(cage.label, 10) : undefined);
      if (cageValue === undefined || isNaN(cageValue)) {
        continue;
      }

      const entry: EliminationEntry = { cageRef: `@${cage.topLeft.ref}`, reasons: [] };

      const validCandidates = this.computeValidCandidates(
        cage,
        cageValue,
        puzzle.hasOperators,
        puzzle.puzzleSize
      );

      for (const cell of cage.cells) {
        if (cell.isSolved) {
          continue;
        }
        const toEliminate = cell.getCandidates().filter((v) => !validCandidates.has(v));
        if (toEliminate.length > 0) {
          changes.push(new CandidatesStrikethrough(cell, toEliminate));
          for (const v of toEliminate) {
            entry.reasons.push(this.buildReason(v, cage, cageValue, puzzle.puzzleSize));
          }
        }
      }

      if (entry.reasons.length > 0) {
        entries.push(entry);
      }
    }

    if (changes.length === 0) {
      return null;
    }

    const note = `Cage operation: ${entries.map((e) => `${e.cageRef} ${e.reasons.join(', ')}`).join(', ')}`;
    return { changes, note };
  }

  private buildReason(value: number, cage: Cage, cageValue: number, puzzleSize: number): string {
    if (cage.operator === 'x' || cage.operator === '*') {
      return `${String(value)} doesn't divide ${String(cageValue)}`;
    }
    if (cage.operator === '+') {
      const otherCellCount = cage.cells.length - 1;
      const maxOtherSum = otherCellCount * puzzleSize;
      if (cageValue - value > maxOtherSum) {
        return `${String(value)} too small`;
      }
      return `${String(value)} too big`;
    }
    if (cage.operator === '-') {
      return `${String(value)} impossible`;
    }
    if (cage.operator === '/') {
      return `${String(value)} impossible`;
    }
    // Unknown operator — generic reason
    return `${String(value)} impossible`;
  }

  private checkAddition(value: number, cageValue: number, cellCount: number, puzzleSize: number): boolean {
    const otherCellCount = cellCount - 1;
    const minOtherSum = otherCellCount;
    const maxOtherSum = otherCellCount * puzzleSize;
    const remainder = cageValue - value;
    return remainder >= minOtherSum && remainder <= maxOtherSum;
  }

  private checkDivision(value: number, cageValue: number, puzzleSize: number): boolean {
    // C is the smaller value: c * cageValue <= puzzleSize
    if (value * cageValue <= puzzleSize) {
      return true;
    }
    // C is the larger value: c % cageValue === 0 and c / cageValue >= 1
    return value % cageValue === 0 && value / cageValue >= 1;
  }

  private checkMultiplication(value: number, cageValue: number, puzzleSize: number, cellCount: number): boolean {
    if (cageValue % value !== 0) {
      return false;
    }
    if (cellCount === BINARY_CELL_COUNT) {
      return cageValue / value <= puzzleSize;
    }
    return true;
  }

  private checkSubtraction(value: number, cageValue: number, puzzleSize: number): boolean {
    return value + cageValue <= puzzleSize || value - cageValue >= 1;
  }

  private computeValidCandidates(
    cage: Cage,
    cageValue: number,
    hasOperators: boolean,
    puzzleSize: number
  ): Set<number> {
    const valid = new Set<number>();
    const cellCount = cage.cells.length;

    if (hasOperators && cage.operator) {
      for (let c = 1; c <= puzzleSize; c++) {
        if (this.isValidForOperator(c, cage.operator, cageValue, puzzleSize, cellCount)) {
          valid.add(c);
        }
      }
    } else {
      // Unknown operator: union across all applicable operators
      const operators = cellCount === BINARY_CELL_COUNT
        ? ['+', '-', 'x', '/']
        : ['+', 'x'];
      for (const op of operators) {
        for (let c = 1; c <= puzzleSize; c++) {
          if (this.isValidForOperator(c, op, cageValue, puzzleSize, cellCount)) {
            valid.add(c);
          }
        }
      }
    }

    return valid;
  }

  private isValidForOperator(
    value: number,
    operator: string,
    cageValue: number,
    puzzleSize: number,
    cellCount: number
  ): boolean {
    switch (operator) {
      case '-':
        return cellCount === BINARY_CELL_COUNT && this.checkSubtraction(value, cageValue, puzzleSize);
      case '*':
      case 'x':
        return this.checkMultiplication(value, cageValue, puzzleSize, cellCount);
      case '/':
        return cellCount === BINARY_CELL_COUNT && this.checkDivision(value, cageValue, puzzleSize);
      case '+':
        return this.checkAddition(value, cageValue, cellCount, puzzleSize);
      default:
        return true;
    }
  }
}
