import type { CellChange } from './cellChanges/CellChange.ts';
import type { CellOperation } from './parsers.ts';
import type { Strategy } from './strategies/Strategy.ts';

import { buildCageConstraintChanges } from './cageConstraints.ts';
import { CandidatesChange } from './cellChanges/CandidatesChange.ts';
import { CandidatesStrikethrough } from './cellChanges/CandidatesStrikethrough.ts';
import { CellClearance } from './cellChanges/CellClearance.ts';
import { ValueChange } from './cellChanges/ValueChange.ts';
import {
  cellRefA1,
  parseOperation
} from './parsers.ts';
import { ensureNonNullable } from './typeGuards.ts';

export interface CageRaw {
  readonly cells: readonly string[];
  readonly label?: string;
  readonly operator?: string;
  readonly value?: number;
}

export interface CellValueSetter {
  readonly cell: Cell;
  readonly value: number;
}

export type HouseType = 'column' | 'row';

export interface PuzzleJson {
  readonly cages: readonly CageRaw[];
  readonly hasOperators?: boolean;
  readonly meta?: string;
  readonly size: number;
  readonly title?: string;
}

export interface PuzzleRenderer {
  beginPendingRender(gridSize: number): void;
  renderCommittedChanges(gridSize: number): void;
  renderPendingCandidates(change: CandidatesChange): void;
  renderPendingClearance(change: CellClearance): void;
  renderPendingStrikethrough(change: CandidatesStrikethrough): void;
  renderPendingValue(change: ValueChange): void;
}

export interface PuzzleState {
  readonly cages: readonly CageRaw[];
  readonly hasOperators: boolean;
  readonly size: number;
}

interface EnterCommand {
  readonly cells: readonly Cell[];
  readonly operation: CellOperation;
}

const CHAR_CODE_A = 65;

export class Cage {
  public get topLeft(): Cell {
    return ensureNonNullable(this.cells[0]);
  }

  public constructor(
    public readonly id: number,
    public readonly cells: readonly Cell[],
    public readonly label: string | undefined,
    public readonly operator: string | undefined,
    public readonly value: number | undefined
  ) {
  }

  public contains(cell: Cell): boolean {
    return this.cells.includes(cell);
  }

  public toString(): string {
    const labelPart = this.label ?? '';
    const cellRefs = this.cells.map(String).join(',');
    return `Cage(${labelPart} ${cellRefs})`;
  }
}

export class Cell {
  public readonly ref: string;
  public get cage(): Cage {
    return this.puzzle.getCage(this.cageIndex + 1);
  }

  public get candidateCount(): number {
    return this._candidates.size;
  }

  public get column(): House {
    return this.puzzle.getColumn(this.colIndex + 1);
  }

  public get isSolved(): boolean {
    return this._value !== null;
  }

  public get peers(): readonly Cell[] {
    if (!this._peers) {
      const result: Cell[] = [];
      for (const cell of this.row.cells) {
        if (cell !== this) {
          result.push(cell);
        }
      }
      for (const cell of this.column.cells) {
        if (cell !== this) {
          result.push(cell);
        }
      }
      this._peers = result;
    }
    return this._peers;
  }

  public get peerValues(): readonly number[] {
    const result: number[] = [];
    for (const peer of this.peers) {
      if (peer.value !== null) {
        result.push(peer.value);
      }
    }
    return result;
  }

  public get row(): House {
    return this.puzzle.getRow(this.rowIndex + 1);
  }

  public get value(): null | number {
    return this._value;
  }

  private readonly _candidates = new Set<number>();
  private _peers: null | readonly Cell[] = null;

  private _value: null | number = null;

  public constructor(
    private readonly puzzle: Puzzle,
    private readonly rowIndex: number,
    private readonly colIndex: number,
    private readonly cageIndex: number
  ) {
    this.ref = cellRefA1(rowIndex, colIndex);
  }

  public addCandidate(value: number): void {
    this._candidates.add(value);
  }

  public clearCandidates(): void {
    this._candidates.clear();
  }

  public clearValue(): void {
    this._value = null;
  }

  public getCandidates(): number[] {
    return [...this._candidates].sort((a, b) => a - b);
  }

  public hasCandidate(value: number): boolean {
    return this._candidates.has(value);
  }

  public removeCandidate(value: number): void {
    this._candidates.delete(value);
  }

  public setCandidates(values: Iterable<number>): void {
    this._candidates.clear();
    for (const v of values) {
      this._candidates.add(v);
    }
  }

  public setValue(value: number): void {
    this._value = value;
  }

  public toString(): string {
    return this.ref;
  }
}

export class House {
  public readonly label: string;

  public constructor(public readonly type: HouseType, public readonly id: number, public readonly cells: readonly Cell[]) {
    this.label = type === 'row' ? String(id) : String.fromCharCode(CHAR_CODE_A + id - 1);
  }

  public getCell(id: number): Cell {
    return ensureNonNullable(this.cells[id - 1]);
  }

  public toString(): string {
    return `${this.type === 'row' ? 'Row' : 'Column'} ${this.label}`;
  }
}

export class Puzzle {
  public readonly cages: readonly Cage[];
  public readonly cells: readonly Cell[];
  public readonly columns: readonly House[];
  public readonly houses: readonly House[];
  public readonly rows: readonly House[];
  private pendingChanges: readonly CellChange[] = [];

  public constructor(
    private readonly renderer: PuzzleRenderer,
    public readonly size: number,
    cagesRaw: readonly CageRaw[],
    public readonly hasOperators: boolean,
    public readonly title: string,
    public readonly meta: string,
    private readonly strategies: readonly Strategy[],
    initialValues?: Map<string, number>,
    initialCandidates?: Map<string, Set<number>>
  ) {
    const cellToCageIndex: Record<string, number> = {};
    for (let i = 0; i < cagesRaw.length; i++) {
      const cage = ensureNonNullable(cagesRaw[i]);
      for (const cellRef of cage.cells) {
        cellToCageIndex[cellRef] = i;
      }
    }

    const grid: Cell[][] = [];
    for (let r = 0; r < size; r++) {
      const row: Cell[] = [];
      for (let c = 0; c < size; c++) {
        const ref = cellRefA1(r, c);
        const cageIndex = cellToCageIndex[ref];
        if (cageIndex === undefined) {
          throw new Error(`Cell ${ref} not found in any cage`);
        }
        row.push(new Cell(this, r, c, cageIndex));
      }
      grid.push(row);
    }

    const rows: House[] = [];
    const columns: House[] = [];
    for (let i = 0; i < size; i++) {
      rows.push(new House('row', i + 1, ensureNonNullable(grid[i])));
      columns.push(new House('column', i + 1, grid.map((gridRow) => ensureNonNullable(gridRow[i]))));
    }
    this.rows = rows;
    this.columns = columns;
    this.houses = [...rows, ...columns];

    const cages: Cage[] = [];
    for (let i = 0; i < cagesRaw.length; i++) {
      const raw = ensureNonNullable(cagesRaw[i]);
      const cageCells = raw.cells.map((ref) => {
        const parsed = parseCellRef(ref);
        return ensureNonNullable(ensureNonNullable(grid[parsed.rowId - 1])[parsed.columnId - 1]);
      });
      cageCells.sort((a, b) => a.row.id - b.row.id || a.column.id - b.column.id);
      cages.push(new Cage(i + 1, cageCells, raw.label, raw.operator, raw.value));
    }
    this.cages = cages;

    this.cells = rows.flatMap((row) => [...row.cells]);

    if (initialValues) {
      for (const [ref, cellValue] of initialValues) {
        this.getCell(ref).setValue(cellValue);
      }
    }
    if (initialCandidates) {
      for (const [ref, cands] of initialCandidates) {
        this.getCell(ref).setCandidates(cands);
      }
    }
  }

  public applyChanges(changes: readonly CellChange[]): void {
    this.pendingChanges = changes;
    this.renderer.beginPendingRender(this.size);
    for (const change of changes) {
      change.renderPending(this.renderer);
    }
  }

  public applyEasyStrategies(): number {
    let totalSteps = 0;
    let canApply = true;
    while (canApply) {
      canApply = false;
      for (const strategy of this.strategies) {
        const result = strategy.tryApply(this);
        if (result) {
          this.applyChanges(result);
          this.commit();
          totalSteps++;
          canApply = true;
          break;
        }
      }
    }
    return totalSteps;
  }

  public buildInitChanges(): CellChange[][] {
    const batches: CellChange[][] = [];
    const allValues = Array.from({ length: this.size }, (_, i) => i + 1);
    const allCandidateChanges: CellChange[] = [];
    for (const cell of this.cells) {
      allCandidateChanges.push(new CandidatesChange(cell, allValues));
    }
    batches.push(allCandidateChanges);
    const cageChanges = buildCageConstraintChanges(this.cages, this.hasOperators, this.size);
    if (cageChanges.length > 0) {
      batches.push(cageChanges);
    }
    return batches;
  }

  public commit(): void {
    for (const change of this.pendingChanges) {
      change.applyToModel();
    }
    this.renderer.renderCommittedChanges(this.size);
    this.pendingChanges = [];
  }

  public enter(input: string): void {
    const changes = this.buildEnterChanges(input);
    this.applyChanges(changes);
  }

  public getCage(id: number): Cage {
    return ensureNonNullable(this.cages[id - 1]);
  }

  public getCell(ref: string): Cell;
  public getCell(rowId: number, columnId: number): Cell;
  public getCell(refOrRowId: number | string, columnId?: number): Cell {
    if (typeof refOrRowId === 'string') {
      const parsed = parseCellRef(refOrRowId);
      return this.getRow(parsed.rowId).getCell(parsed.columnId);
    }
    return this.getRow(refOrRowId).getCell(ensureNonNullable(columnId));
  }

  public getColumn(id: number): House {
    return ensureNonNullable(this.columns[id - 1]);
  }

  public getRow(id: number): House {
    return ensureNonNullable(this.rows[id - 1]);
  }

  private buildEnterChanges(input: string): CellChange[] {
    const commands = this.parseInput(input);
    const changes: CellChange[] = [];
    for (const cmd of commands) {
      for (const cell of cmd.cells) {
        switch (cmd.operation.type) {
          case 'candidates':
            changes.push(new CandidatesChange(cell, cmd.operation.values));
            break;
          case 'clear':
            changes.push(new CellClearance(cell));
            break;
          case 'strikethrough':
            changes.push(new CandidatesStrikethrough(cell, cmd.operation.values));
            break;
          case 'value':
            changes.push(new ValueChange(cell, cmd.operation.value));
            break;
          default: {
            const exhaustive: never = cmd.operation;
            throw new Error(`Unknown operation type: ${String(exhaustive)}`);
          }
        }
      }
      if (cmd.operation.type === 'value') {
        const cell = ensureNonNullable(cmd.cells[0]);
        for (const peer of cell.row.cells) {
          if (peer !== cell) {
            changes.push(new CandidatesStrikethrough(peer, [cmd.operation.value]));
          }
        }
        for (const peer of cell.column.cells) {
          if (peer !== cell) {
            changes.push(new CandidatesStrikethrough(peer, [cmd.operation.value]));
          }
        }
      } else if (cmd.operation.type === 'candidates') {
        for (const cell of cmd.cells) {
          const conflicting = cell.peerValues;
          if (conflicting.length > 0) {
            changes.push(new CandidatesStrikethrough(cell, [...conflicting]));
          }
        }
      }
    }
    return changes;
  }

  private parseCellPart(cellPart: string): Cell[] {
    let inner = cellPart;
    if (inner.startsWith('(') && inner.endsWith(')')) {
      inner = inner.substring(1, inner.length - 1);
    }
    const cells: Cell[] = [];
    for (const token of inner.trim().split(/\s+/)) {
      if (token.startsWith('@')) {
        const anchor = this.getCell(token.substring(1));
        for (const cell of anchor.cage.cells) {
          cells.push(cell);
        }
      } else {
        cells.push(this.getCell(token));
      }
    }
    return cells;
  }

  private parseInput(input: string): EnterCommand[] {
    const trimmed = input.trim();
    if (!trimmed) {
      throw new Error('No commands specified');
    }
    const pattern = /(?:\([^)]+\)|@?[A-Za-z]\d+):[^\s]+/g;
    const matches = trimmed.match(pattern);
    if (!matches || matches.length === 0) {
      throw new Error('Invalid format. Expected: A1:=1, A1:234, (B2 C3):-567, @D4:234');
    }
    const commands: EnterCommand[] = [];
    for (const match of matches) {
      const colonIdx = match.startsWith('(')
        ? match.indexOf(':', match.indexOf(')'))
        : match.indexOf(':');
      const cellPart = match.substring(0, colonIdx);
      const opPart = match.substring(colonIdx + 1);
      const cells = this.parseCellPart(cellPart);
      const operation = parseOperation(opPart, cells.length);
      commands.push({ cells, operation });
    }
    return commands;
  }
}

// Re-export parseCellRef for use outside
import { parseCellRef } from './parsers.ts';
