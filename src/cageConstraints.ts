import type { CellChange } from './cellChanges/CellChange.ts';
import type { CellValueSetter } from './Puzzle.ts';

import { CandidatesChange } from './cellChanges/CandidatesChange.ts';
import { CandidatesStrikethrough } from './cellChanges/CandidatesStrikethrough.ts';
import { ValueChange } from './cellChanges/ValueChange.ts';
import { evaluateTuple } from './combinatorics.ts';
import {
  Cage,
  Cell
} from './Puzzle.ts';
import { ensureNonNullable } from './typeGuards.ts';

export function applyCageConstraint(
  cage: Cage,
  hasOperators: boolean,
  gridSize: number,
  valueSetters: CellValueSetter[],
  candidateChanges: CellChange[]
): void {
  const cageValue = cage.value ?? (cage.label ? parseInt(cage.label, 10) : undefined);
  if (cageValue === undefined || isNaN(cageValue)) {
    return;
  }

  const tuples = collectCageTuples(cageValue, cage, hasOperators, gridSize);
  if (tuples.length === 0) {
    return;
  }

  if (tuples.length === 1) {
    const tuple = ensureNonNullable(tuples[0]);
    for (let i = 0; i < cage.cells.length; i++) {
      valueSetters.push({ cell: ensureNonNullable(cage.cells[i]), value: ensureNonNullable(tuple[i]) });
    }
    return;
  }

  const distinctSets = new Set<string>(tuples.map(
    (t) => [...new Set(t)].sort((a, b) => a - b).join(',')
  ));
  if (distinctSets.size !== 1) {
    return;
  }

  const narrowedValues = ensureNonNullable([...distinctSets][0]).split(',').map(Number);
  if (narrowedValues.length >= gridSize) {
    return;
  }

  for (const cell of cage.cells) {
    candidateChanges.push(new CandidatesChange(cell, narrowedValues));
  }
}

export function buildAutoEliminateChanges(
  cellValueSetters: readonly CellValueSetter[]
): CellChange[] {
  const changes: CellChange[] = [];
  for (const setter of cellValueSetters) {
    changes.push(new ValueChange(setter.cell, setter.value));
    for (const peer of setter.cell.peers) {
      changes.push(new CandidatesStrikethrough(peer, [setter.value]));
    }
  }
  return changes;
}

export function collectCageTuples(
  cageValue: number,
  cage: Cage,
  hasOperators: boolean,
  gridSize: number
): number[][] {
  if (hasOperators && cage.operator) {
    return computeValidCageTuples(cageValue, cage.operator, cage.cells, gridSize);
  }
  const tupleSet = new Set<string>();
  const tuples: number[][] = [];
  for (const op of ['+', '-', 'x', '/']) {
    for (const t of computeValidCageTuples(cageValue, op, cage.cells, gridSize)) {
      const key = t.join(',');
      if (!tupleSet.has(key)) {
        tupleSet.add(key);
        tuples.push(t);
      }
    }
  }
  return tuples;
}

export function computeValidCageTuples(
  value: number,
  operator: string,
  cells: readonly Cell[],
  gridSize: number
): number[][] {
  const tuples: number[][] = [];
  const numCells = cells.length;

  function search(tuple: number[], depth: number): void {
    if (depth === numCells) {
      if (evaluateTuple(tuple, operator) === value) {
        tuples.push([...tuple]);
      }
      return;
    }
    const cell = ensureNonNullable(cells[depth]);
    for (let v = 1; v <= gridSize; v++) {
      let valid = true;
      for (let i = 0; i < depth; i++) {
        if (ensureNonNullable(tuple[i]) === v) {
          const prevCell = ensureNonNullable(cells[i]);
          if (prevCell.row === cell.row || prevCell.column === cell.column) {
            valid = false;
            break;
          }
        }
      }
      if (!valid) {
        continue;
      }
      tuple.push(v);
      search(tuple, depth + 1);
      tuple.pop();
    }
  }

  search([], 0);
  return tuples;
}
