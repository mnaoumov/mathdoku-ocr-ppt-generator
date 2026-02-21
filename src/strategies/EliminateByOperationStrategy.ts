import type {
  Cage,
  Cell,
  HouseType,
  Puzzle
} from '../Puzzle.ts';
import type {
  Strategy,
  StrategyResult
} from './Strategy.ts';

import { CandidatesStrikethrough } from '../cellChanges/CandidatesStrikethrough.ts';
import { ensureNonNullable } from '../typeGuards.ts';

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

      const eliminatedValues = new Set<number>();

      for (const cell of cage.cells) {
        if (cell.isSolved) {
          continue;
        }
        const validForCell = this.computeValidCandidatesForCell(
          cell,
          cage,
          cageValue,
          puzzle.hasOperators,
          puzzle.puzzleSize
        );
        const toEliminate = cell.getCandidates().filter((v) => !validForCell.has(v));
        if (toEliminate.length > 0) {
          changes.push(new CandidatesStrikethrough(cell, toEliminate));
          for (const v of toEliminate) {
            eliminatedValues.add(v);
          }
        }
      }

      if (eliminatedValues.size > 0) {
        const reasons = this.buildGroupedReasons(eliminatedValues, cage, cageValue, puzzle.puzzleSize);
        entries.push({ cageRef: `@${cage.topLeft.ref}`, reasons });
      }
    }

    if (changes.length === 0) {
      return null;
    }

    const note = `Cage operation. ${entries.map((e) => `${e.cageRef}: ${e.reasons.join(', ')}`).join(', ')}`;
    return { changes, note };
  }

  private buildGroupedReasons(
    eliminatedValues: Set<number>,
    cage: Cage,
    cageValue: number,
    puzzleSize: number
  ): string[] {
    const sorted = [...eliminatedValues].sort((a, b) => a - b);

    if (cage.operator === 'x' || cage.operator === '*') {
      const noDivide = sorted.filter((v) => cageValue % v !== 0);
      const infeasible = sorted.filter((v) => cageValue % v === 0);
      const reasons: string[] = [];
      if (noDivide.length > 0) {
        reasons.push(formatValueGroup(noDivide, 'divide', cageValue));
      }
      if (infeasible.length > 0) {
        reasons.push(formatValueGroup(infeasible, 'impossible'));
      }
      return reasons;
    }

    if (cage.operator === '+') {
      const solvedSum = cage.cells
        .filter((c) => c.isSolved)
        .reduce((s, c) => s + ensureNonNullable(c.value), 0);
      const otherUnsolvedCount = cage.cells.filter((c) => !c.isSolved).length - 1;
      const naiveMaxOtherSum = otherUnsolvedCount * puzzleSize;

      const tooBig: number[] = [];
      const tooSmall: number[] = [];
      for (const v of sorted) {
        const remainder = cageValue - v - solvedSum;
        if (remainder > naiveMaxOtherSum) {
          tooSmall.push(v);
        } else {
          tooBig.push(v);
        }
      }

      const reasons: string[] = [];
      if (tooSmall.length > 0) {
        reasons.push(formatValueGroup(tooSmall, 'tooSmall'));
      }
      if (tooBig.length > 0) {
        reasons.push(formatValueGroup(tooBig, 'tooBig'));
      }
      return reasons;
    }

    return [formatValueGroup(sorted, 'impossible')];
  }

  private checkAdditionForCell(
    value: number,
    cell: Cell,
    otherCells: readonly Cell[],
    cageValue: number,
    solvedSum: number,
    puzzleSize: number
  ): boolean {
    const remainder = cageValue - value - solvedSum;
    if (otherCells.length === 0) {
      return remainder === 0;
    }
    const minOtherSum = computeLatinSquareBound(cell, otherCells, value, puzzleSize, 'sum', 'min');
    const maxOtherSum = computeLatinSquareBound(cell, otherCells, value, puzzleSize, 'sum', 'max');
    return remainder >= minOtherSum && remainder <= maxOtherSum;
  }

  private checkDivision(value: number, cageValue: number, puzzleSize: number): boolean {
    if (value * cageValue <= puzzleSize) {
      return true;
    }
    return value % cageValue === 0 && value / cageValue >= 1;
  }

  private checkMultiplicationForCell(
    value: number,
    cell: Cell,
    otherCells: readonly Cell[],
    cageValue: number,
    solvedProduct: number,
    puzzleSize: number,
    cellCount: number
  ): boolean {
    const totalProduct = value * solvedProduct;
    if (cageValue % totalProduct !== 0) {
      return false;
    }
    const quotient = cageValue / totalProduct;
    if (otherCells.length === 0) {
      return quotient === 1;
    }
    if (cellCount === BINARY_CELL_COUNT) {
      return quotient <= puzzleSize;
    }
    const minOtherProduct = computeLatinSquareBound(cell, otherCells, value, puzzleSize, 'product', 'min');
    const maxOtherProduct = computeLatinSquareBound(cell, otherCells, value, puzzleSize, 'product', 'max');
    return quotient >= minOtherProduct && quotient <= maxOtherProduct;
  }

  private checkSubtraction(value: number, cageValue: number, puzzleSize: number): boolean {
    return value + cageValue <= puzzleSize || value - cageValue >= 1;
  }

  private computeValidCandidatesForCell(
    cell: Cell,
    cage: Cage,
    cageValue: number,
    hasOperators: boolean,
    puzzleSize: number
  ): Set<number> {
    const valid = new Set<number>();
    const cellCount = cage.cells.length;
    const otherCells = cage.cells.filter((c) => c !== cell && !c.isSolved);
    const solvedValues = cage.cells.filter((c) => c.isSolved).map((c) => ensureNonNullable(c.value));
    const solvedSum = solvedValues.reduce((s, v) => s + v, 0);
    const solvedProduct = solvedValues.reduce((p, v) => p * v, 1);

    if (hasOperators && cage.operator) {
      for (let v = 1; v <= puzzleSize; v++) {
        if (this.isValidForCellOperator(v, cell, otherCells, cage.operator, cageValue, solvedSum, solvedProduct, puzzleSize, cellCount)) {
          valid.add(v);
        }
      }
    } else {
      const operators = cellCount === BINARY_CELL_COUNT
        ? ['+', '-', 'x', '/']
        : ['+', 'x'];
      for (const op of operators) {
        for (let v = 1; v <= puzzleSize; v++) {
          if (this.isValidForCellOperator(v, cell, otherCells, op, cageValue, solvedSum, solvedProduct, puzzleSize, cellCount)) {
            valid.add(v);
          }
        }
      }
    }

    return valid;
  }

  private isValidForCellOperator(
    value: number,
    cell: Cell,
    otherCells: readonly Cell[],
    operator: string,
    cageValue: number,
    solvedSum: number,
    solvedProduct: number,
    puzzleSize: number,
    cellCount: number
  ): boolean {
    switch (operator) {
      case '-':
        return cellCount === BINARY_CELL_COUNT && this.checkSubtraction(value, cageValue, puzzleSize);
      case '*':
      case 'x':
        return this.checkMultiplicationForCell(value, cell, otherCells, cageValue, solvedProduct, puzzleSize, cellCount);
      case '/':
        return cellCount === BINARY_CELL_COUNT && this.checkDivision(value, cageValue, puzzleSize);
      case '+':
        return this.checkAdditionForCell(value, cell, otherCells, cageValue, solvedSum, puzzleSize);
      default:
        return true;
    }
  }
}

function computeHouseBound(
  targetCell: Cell,
  otherCells: readonly Cell[],
  value: number,
  puzzleSize: number,
  aggregateType: 'product' | 'sum',
  boundType: 'max' | 'min',
  houseType: HouseType
): number {
  const groups = new Map<number, number>();
  for (const cell of otherCells) {
    const houseId = houseType === 'row' ? cell.row.id : cell.column.id;
    groups.set(houseId, (groups.get(houseId) ?? 0) + 1);
  }

  const targetHouseId = houseType === 'row' ? targetCell.row.id : targetCell.column.id;

  let result = aggregateType === 'sum' ? 0 : 1;

  for (const [houseId, count] of groups) {
    const excludeValue = houseId === targetHouseId ? value : undefined;
    const fn = boundType === 'min' ? minDistinctAggregate : maxDistinctAggregate;
    const groupBound = fn(count, puzzleSize, aggregateType, excludeValue);
    result = aggregateType === 'sum' ? result + groupBound : result * groupBound;
  }

  return result;
}

function computeLatinSquareBound(
  targetCell: Cell,
  otherCells: readonly Cell[],
  value: number,
  puzzleSize: number,
  aggregateType: 'product' | 'sum',
  boundType: 'max' | 'min'
): number {
  const rowBound = computeHouseBound(targetCell, otherCells, value, puzzleSize, aggregateType, boundType, 'row');
  const colBound = computeHouseBound(targetCell, otherCells, value, puzzleSize, aggregateType, boundType, 'column');
  return boundType === 'min' ? Math.max(rowBound, colBound) : Math.min(rowBound, colBound);
}

function formatValueGroup(
  values: readonly number[],
  reasonType: 'divide' | 'impossible' | 'tooBig' | 'tooSmall',
  cageValue?: number
): string {
  const valueStr = values.length === 1
    ? String(ensureNonNullable(values[0]))
    : `{${values.join('')}}`;

  switch (reasonType) {
    case 'divide':
      return values.length === 1
        ? `${valueStr} doesn't divide ${String(cageValue)}`
        : `${valueStr} don't divide ${String(cageValue)}`;
    case 'impossible':
      return `${valueStr} impossible`;
    case 'tooBig':
      return `${valueStr} too big`;
    case 'tooSmall':
      return `${valueStr} too small`;
    default:
      return `${valueStr} impossible`;
  }
}

function maxDistinctAggregate(
  count: number,
  puzzleSize: number,
  aggregateType: 'product' | 'sum',
  excludedValue?: number
): number {
  let result = aggregateType === 'sum' ? 0 : 1;
  let picked = 0;
  for (let v = puzzleSize; v >= 1 && picked < count; v--) {
    if (v !== excludedValue) {
      result = aggregateType === 'sum' ? result + v : result * v;
      picked++;
    }
  }
  return result;
}

function minDistinctAggregate(
  count: number,
  puzzleSize: number,
  aggregateType: 'product' | 'sum',
  excludedValue?: number
): number {
  let result = aggregateType === 'sum' ? 0 : 1;
  let picked = 0;
  for (let v = 1; v <= puzzleSize && picked < count; v++) {
    if (v !== excludedValue) {
      result = aggregateType === 'sum' ? result + v : result * v;
      picked++;
    }
  }
  return result;
}
