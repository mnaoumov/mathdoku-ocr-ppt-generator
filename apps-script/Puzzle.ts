/**
 * Puzzle.ts -- Business logic layer (zero GoogleAppsScript references).
 *
 * Contains: Puzzle class (model + strategy loop + Enter parsing),
 * PuzzleRenderer / Strategy interfaces, shared types, pure utilities.
 */

interface Cage {
  readonly cells: readonly string[];
  readonly label?: string;
  readonly op?: string;
  readonly value?: number;
}

interface CandidatesOp {
  readonly digits: string;
  readonly type: 'candidates';
}

type CellOperation = CandidatesOp | ClearOp | StrikeOp | ValueOp;

interface CellRef {
  readonly c: number;
  readonly r: number;
}

interface ClearOp {
  readonly type: 'clear';
}

interface EnterCommand {
  readonly cellRefs: readonly string[];
  readonly operation: CellOperation;
}

interface GridBoundaries {
  readonly horizontalBounds: readonly (readonly boolean[])[];
  readonly verticalBounds: readonly (readonly boolean[])[];
}

interface PuzzleJson {
  readonly cages: readonly Cage[];
  readonly meta?: string;
  readonly operations?: boolean;
  readonly size: number;
  readonly title?: string;
}

interface PuzzleRenderer {
  renderCommittedChanges(gridSize: number): void;
  renderPendingChanges(changes: readonly SolveChange[], gridSize: number): void;
}

interface PuzzleState {
  readonly cages: readonly Cage[];
  readonly operations: boolean;
  readonly size: number;
}

type SolveChange =
  | { readonly cellRef: string; readonly digits: string; readonly type: 'candidates' }
  | { readonly cellRef: string; readonly digits: string; readonly type: 'strike' }
  | { readonly cellRef: string; readonly digits: string; readonly type: 'value' }
  | { readonly cellRef: string; readonly type: 'clear' };

interface Strategy {
  tryApply(puzzle: Puzzle): null | readonly SolveChange[];
}

interface StrikeOp {
  readonly digits: string;
  readonly type: 'strike';
}

interface ValueOp {
  readonly digit: string;
  readonly type: 'value';
}

interface ValueSet {
  readonly cellRef: string;
  readonly digit: string;
}

class Puzzle {
  public readonly cages: readonly Cage[];
  public readonly houses: readonly (readonly string[])[];
  public readonly meta: string;
  public readonly operations: boolean;
  public readonly size: number;

  public readonly title: string;
  private readonly candidates: Map<string, Set<string>>;
  private readonly renderer: PuzzleRenderer;
  private readonly strategies: readonly Strategy[];
  private readonly values: Map<string, string>;

  public constructor(
    renderer: PuzzleRenderer,
    size: number,
    cages: readonly Cage[],
    operations: boolean,
    title: string,
    meta: string,
    initialValues?: Map<string, string>,
    initialCandidates?: Map<string, Set<string>>
  ) {
    this.renderer = renderer;
    this.size = size;
    this.cages = cages;
    this.operations = operations;
    this.title = title;
    this.meta = meta;
    this.values = initialValues ?? new Map<string, string>();
    this.candidates = initialCandidates ?? new Map<string, Set<string>>();
    const houses: string[][] = [];
    for (let i = 0; i < size; i++) {
      const row: string[] = [];
      const col: string[] = [];
      for (let j = 0; j < size; j++) {
        row.push(cellRefA1(i, j));
        col.push(cellRefA1(j, i));
      }
      houses.push(row);
      houses.push(col);
    }
    this.houses = houses;
    this.strategies = [
      new SingleCandidateStrategy(),
      new HiddenSingleStrategy(),
      ...Array.from({ length: size - MIN_NAKED_SET_SIZE }, (_, i) => new NakedSetStrategy(i + MIN_NAKED_SET_SIZE))
    ];
  }

  public applyChanges(changes: readonly SolveChange[]): void {
    this.renderer.renderPendingChanges(changes, this.size);
    this.updateModel(changes);
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

  public buildInitChanges(): SolveChange[][] {
    const batches: SolveChange[][] = [];
    const allDigits = Array.from({ length: this.size }, (_, i) => String(i + 1)).join('');
    const allCandidateChanges: SolveChange[] = [];
    for (let r = 0; r < this.size; r++) {
      for (let c = 0; c < this.size; c++) {
        allCandidateChanges.push({ cellRef: cellRefA1(r, c), digits: allDigits, type: 'candidates' });
      }
    }
    batches.push(allCandidateChanges);
    const cageChanges = buildCageConstraintChanges(this.cages, this.operations, this.size);
    if (cageChanges.length > 0) {
      batches.push(cageChanges);
    }
    return batches;
  }

  public commit(): void {
    this.renderer.renderCommittedChanges(this.size);
  }

  public enter(input: string): void {
    const changes = this.buildEnterChanges(input);
    this.applyChanges(changes);
  }

  public getCageCells(r: number, c: number): readonly string[] {
    const ref = cellRefA1(r, c);
    for (const cage of this.cages) {
      if (cage.cells.includes(ref)) {
        return cage.cells;
      }
    }
    throw new Error(`Cell ${ref} not found in any cage`);
  }

  public getCellCandidates(cellRef: string): string[] {
    const cands = this.candidates.get(cellRef);
    if (!cands) {
      return [];
    }
    return [...cands].sort();
  }

  public getCellValue(cellRef: string): null | string {
    return this.values.get(cellRef) ?? null;
  }

  private buildEnterChanges(input: string): SolveChange[] {
    const commands = this.parseInput(input);
    const changes: SolveChange[] = [];
    for (const cmd of commands) {
      for (const cellRef of cmd.cellRefs) {
        switch (cmd.operation.type) {
          case 'candidates':
            changes.push({ cellRef, digits: cmd.operation.digits, type: 'candidates' });
            break;
          case 'clear':
            changes.push({ cellRef, type: 'clear' });
            break;
          case 'strike':
            changes.push({ cellRef, digits: cmd.operation.digits, type: 'strike' });
            break;
          case 'value':
            changes.push({ cellRef, digits: cmd.operation.digit, type: 'value' });
            break;
          default: {
            const exhaustive: never = cmd.operation;
            throw new Error(`Unknown operation type: ${String(exhaustive)}`);
          }
        }
      }
      if (cmd.operation.type === 'value') {
        const ref = parseCellRef(ensureNonNullable(cmd.cellRefs[0]));
        const digit = cmd.operation.digit;
        for (let i = 0; i < this.size; i++) {
          if (i !== ref.c) {
            changes.push({ cellRef: cellRefA1(ref.r, i), digits: digit, type: 'strike' });
          }
          if (i !== ref.r) {
            changes.push({ cellRef: cellRefA1(i, ref.c), digits: digit, type: 'strike' });
          }
        }
      } else if (cmd.operation.type === 'candidates') {
        for (const cellRef of cmd.cellRefs) {
          const ref = parseCellRef(cellRef);
          const conflicting = this.getRowColValues(ref.r, ref.c);
          if (conflicting.length > 0) {
            changes.push({ cellRef, digits: conflicting.join(''), type: 'strike' });
          }
        }
      }
    }
    return changes;
  }

  private getRowColValues(r: number, c: number): string[] {
    const result: string[] = [];
    for (let i = 0; i < this.size; i++) {
      if (i !== c) {
        const v = this.values.get(cellRefA1(r, i));
        if (v) {
          result.push(v);
        }
      }
      if (i !== r) {
        const v = this.values.get(cellRefA1(i, c));
        if (v) {
          result.push(v);
        }
      }
    }
    return result;
  }

  private parseCellPart(cellPart: string): string[] {
    let inner = cellPart;
    if (inner.startsWith('(') && inner.endsWith(')')) {
      inner = inner.substring(1, inner.length - 1);
    }
    const cellRefs: string[] = [];
    for (const token of inner.trim().split(/\s+/)) {
      if (token.startsWith('@')) {
        const anchor = parseCellRef(token.substring(1));
        for (const cell of this.getCageCells(anchor.r, anchor.c)) {
          cellRefs.push(cell);
        }
      } else {
        const cell = parseCellRef(token);
        cellRefs.push(cellRefA1(cell.r, cell.c));
      }
    }
    return cellRefs;
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
      const cellRefs = this.parseCellPart(cellPart);
      const operation = parseOperation(opPart, cellRefs.length);
      commands.push({ cellRefs, operation });
    }
    return commands;
  }

  private updateModel(changes: readonly SolveChange[]): void {
    for (const change of changes) {
      switch (change.type) {
        case 'candidates': {
          const digits = new Set(change.digits.split(''));
          this.candidates.set(change.cellRef, digits);
          this.values.delete(change.cellRef);
          break;
        }
        case 'clear':
          this.values.delete(change.cellRef);
          this.candidates.delete(change.cellRef);
          break;
        case 'strike': {
          const existing = this.candidates.get(change.cellRef);
          if (existing) {
            for (const d of change.digits) {
              existing.delete(d);
            }
          }
          break;
        }
        case 'value':
          this.values.set(change.cellRef, change.digits);
          this.candidates.delete(change.cellRef);
          break;
        default: {
          const exhaustive: never = change;
          throw new Error(`Unknown change type: ${String(exhaustive)}`);
        }
      }
    }
  }
}

const BINARY_OP_SIZE = 2;
const CHAR_CODE_A = 65;
const MIN_NAKED_SET_SIZE = 2;

function buildAutoEliminateChanges(
  valueSets: readonly ValueSet[],
  gridSize: number
): SolveChange[] {
  const changes: SolveChange[] = [];
  for (const vs of valueSets) {
    changes.push({ cellRef: vs.cellRef, digits: vs.digit, type: 'value' });
    const ref = parseCellRef(vs.cellRef);
    for (let i = 0; i < gridSize; i++) {
      if (i !== ref.c) {
        changes.push({ cellRef: cellRefA1(ref.r, i), digits: vs.digit, type: 'strike' });
      }
      if (i !== ref.r) {
        changes.push({ cellRef: cellRefA1(i, ref.c), digits: vs.digit, type: 'strike' });
      }
    }
  }
  return changes;
}

function buildCageConstraintChanges(
  cages: readonly Cage[],
  operations: boolean,
  gridSize: number
): SolveChange[] {
  const changes: SolveChange[] = [];
  for (const cage of cages) {
    if (cage.cells.length === 1) {
      const value = cage.value ?? cage.label;
      if (value !== undefined) {
        changes.push({
          cellRef: ensureNonNullable(cage.cells[0]),
          digits: String(value),
          type: 'value'
        });
      }
      continue;
    }

    const cageValue = cage.value ?? (cage.label ? parseInt(cage.label, 10) : undefined);
    if (cageValue === undefined || isNaN(cageValue)) {
      continue;
    }

    let possible: Set<string>;
    if (operations && cage.op) {
      possible = computeCagePossibleDigits(cageValue, cage.op, cage.cells.length, gridSize);
    } else if (operations) {
      continue;
    } else {
      possible = new Set<string>();
      for (const op of ['+', '-', 'x', '/']) {
        for (const d of computeCagePossibleDigits(cageValue, op, cage.cells.length, gridSize)) {
          possible.add(d);
        }
      }
    }

    if (possible.size >= gridSize || possible.size === 0) {
      continue;
    }

    const narrowedDigits = [...possible].sort().join('');
    for (const cellRef of cage.cells) {
      changes.push({ cellRef, digits: narrowedDigits, type: 'candidates' });
    }
  }
  return changes;
}

function cellRefA1(r: number, c: number): string {
  return String.fromCharCode(CHAR_CODE_A + c) + String(r + 1);
}

function computeCagePossibleDigits(
  value: number,
  op: string,
  numCells: number,
  gridSize: number
): Set<string> {
  const possible = new Set<string>();

  function search(tuple: number[], depth: number): void {
    if (depth === numCells) {
      if (evaluateTuple(tuple, op) === value) {
        for (const d of tuple) {
          possible.add(String(d));
        }
      }
      return;
    }
    for (let d = 1; d <= gridSize; d++) {
      tuple.push(d);
      search(tuple, depth + 1);
      tuple.pop();
    }
  }

  search([], 0);
  return possible;
}

function computeGridBoundaries(cages: readonly Cage[], gridDimension: number): GridBoundaries {
  const cellToCage: Record<string, number> = {};
  for (let cageIndex = 0; cageIndex < cages.length; cageIndex++) {
    const cage = ensureNonNullable(cages[cageIndex]);
    for (const cell of cage.cells) {
      cellToCage[cell] = cageIndex;
    }
  }

  const verticalBounds: boolean[][] = [];
  const horizontalBounds: boolean[][] = [];
  for (let r = 0; r < gridDimension; r++) {
    const row: boolean[] = [];
    verticalBounds[r] = row;
    for (let c = 1; c < gridDimension; c++) {
      row[c - 1] = cellToCage[cellRefA1(r, c - 1)] !== cellToCage[cellRefA1(r, c)];
    }
  }
  for (let r = 1; r < gridDimension; r++) {
    const row: boolean[] = [];
    horizontalBounds[r - 1] = row;
    for (let c = 0; c < gridDimension; c++) {
      row[c] = cellToCage[cellRefA1(r - 1, c)] !== cellToCage[cellRefA1(r, c)];
    }
  }

  return { horizontalBounds, verticalBounds };
}

function evaluateTuple(tuple: readonly number[], op: string): null | number {
  if (tuple.length === 0) {
    return null;
  }
  if (tuple.length === 1) {
    return ensureNonNullable(tuple[0]);
  }

  const a = ensureNonNullable(tuple[0]);
  const b = ensureNonNullable(tuple[1]);

  switch (op) {
    case '-': {
      if (tuple.length !== BINARY_OP_SIZE) {
        return null;
      }
      return Math.abs(a - b);
    }
    case '*':
    case 'x': {
      let product = 1;
      for (const d of tuple) {
        product *= d;
      }
      return product;
    }
    case '/': {
      if (tuple.length !== BINARY_OP_SIZE) {
        return null;
      }
      const maxD = Math.max(a, b);
      const minD = Math.min(a, b);
      if (minD === 0) {
        return null;
      }
      return maxD % minD === 0 ? maxD / minD : null;
    }
    case '+': {
      let sum = 0;
      for (const d of tuple) {
        sum += d;
      }
      return sum;
    }
    default:
      return null;
  }
}

function generateSubsets<T>(items: readonly T[], k: number): T[][] {
  if (k === 0) {
    return [[]];
  }
  if (k > items.length) {
    return [];
  }

  const results: T[][] = [];
  const indices: number[] = Array.from({ length: k }, (_, i) => i);

  for (;;) {
    results.push(indices.map((idx) => ensureNonNullable(items[idx])));

    let pos = k - 1;
    while (pos >= 0 && ensureNonNullable(indices[pos]) === pos + items.length - k) {
      pos--;
    }
    if (pos < 0) {
      break;
    }
    indices[pos] = ensureNonNullable(indices[pos]) + 1;
    for (let j = pos + 1; j < k; j++) {
      indices[j] = ensureNonNullable(indices[j - 1]) + 1;
    }
  }

  return results;
}

function parseCellRef(token: string): CellRef {
  const m = /^(?<col>[A-Z])(?<row>[1-9]\d*)$/.exec(token.trim().toUpperCase());
  if (!m) {
    throw new Error(`Bad cell ref: ${token}`);
  }
  const groups = ensureNonNullable(m.groups);
  return {
    c: ensureNonNullable(groups['col']).charCodeAt(0) - CHAR_CODE_A,
    r: parseInt(ensureNonNullable(groups['row']), 10) - 1
  };
}

function parseOperation(text: string, cellCount: number): CellOperation {
  if (text.toLowerCase() === 'x') {
    return { type: 'clear' };
  }

  if (text.startsWith('=')) {
    const digit = text.substring(1);
    if (!/^[1-9]$/.test(digit)) {
      throw new Error('=N expects a single digit 1-9');
    }
    if (cellCount > 1) {
      throw new Error('=N can only be used with a single cell');
    }
    return { digit, type: 'value' };
  }

  if (text.startsWith('-')) {
    const digits = text.substring(1);
    if (!/^[1-9]+$/.test(digits)) {
      throw new Error('-digits expects digits 1-9');
    }
    return { digits, type: 'strike' };
  }

  if (!/^[1-9]+$/.test(text)) {
    throw new Error(`Expected digits 1-9: ${text}`);
  }
  return { digits: text, type: 'candidates' };
}
