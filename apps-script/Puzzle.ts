/**
 * Puzzle.ts -- Business logic layer (zero GoogleAppsScript references).
 *
 * Contains: Cell, House, Cage domain classes, CellChange hierarchy,
 * Puzzle class (model + strategy loop + Enter parsing),
 * PuzzleRenderer / Strategy interfaces, shared types, pure utilities.
 */

interface CageRaw {
  readonly cells: readonly string[];
  readonly label?: string;
  readonly operator?: string;
  readonly value?: number;
}

interface CandidatesOperation {
  readonly type: 'candidates';
  readonly values: readonly number[];
}

interface CellCandidatesSetter {
  readonly cell: Cell;
  readonly values: readonly number[];
}

type CellOperation = CandidatesOperation | ClearanceOperation | StrikethroughOperation | ValueOperation;

interface CellRef {
  readonly columnId: number;
  readonly rowId: number;
}

interface CellValueSetter {
  readonly cell: Cell;
  readonly value: number;
}

interface ClearanceOperation {
  readonly type: 'clear';
}

interface EnterCommand {
  readonly cells: readonly Cell[];
  readonly operation: CellOperation;
}

interface GridBoundaries {
  readonly horizontalBounds: readonly (readonly boolean[])[];
  readonly verticalBounds: readonly (readonly boolean[])[];
}

type HouseType = 'column' | 'row';

interface PuzzleJson {
  readonly cages: readonly CageRaw[];
  readonly hasOperators?: boolean;
  readonly meta?: string;
  readonly size: number;
  readonly title?: string;
}

interface PuzzleRenderer {
  beginPendingRender(gridSize: number): void;
  renderCommittedChanges(gridSize: number): void;
  renderPendingCandidates(change: CandidatesChange): void;
  renderPendingClearance(change: CellClearance): void;
  renderPendingStrikethrough(change: CandidatesStrikethrough): void;
  renderPendingValue(change: ValueChange): void;
}

interface PuzzleState {
  readonly cages: readonly CageRaw[];
  readonly hasOperators: boolean;
  readonly size: number;
}

interface Strategy {
  tryApply(puzzle: Puzzle): null | readonly CellChange[];
}

interface StrikethroughOperation {
  readonly type: 'strikethrough';
  readonly values: readonly number[];
}

interface ValueOperation {
  readonly type: 'value';
  readonly value: number;
}

class Cage {
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

abstract class CellChange {
  protected constructor(public readonly cell: Cell) {
  }

  public abstract applyToModel(): void;
  public abstract renderPending(renderer: PuzzleRenderer): void;
}

class CandidatesChange extends CellChange {
  public constructor(cell: Cell, public readonly values: readonly number[]) {
    super(cell);
  }

  public applyToModel(): void {
    this.cell.setCandidates(this.values);
    this.cell.clearValue();
  }

  public renderPending(renderer: PuzzleRenderer): void {
    renderer.renderPendingCandidates(this);
  }
}

class CandidatesStrikethrough extends CellChange {
  public constructor(cell: Cell, public readonly values: readonly number[]) {
    super(cell);
  }

  public applyToModel(): void {
    for (const v of this.values) {
      this.cell.removeCandidate(v);
    }
  }

  public renderPending(renderer: PuzzleRenderer): void {
    renderer.renderPendingStrikethrough(this);
  }
}

class Cell {
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
    this.rowIndex = rowIndex;
    this.colIndex = colIndex;
    this.cageIndex = cageIndex;
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

class CellClearance extends CellChange {
  public constructor(cell: Cell) {
    super(cell);
  }

  public applyToModel(): void {
    this.cell.clearValue();
    this.cell.clearCandidates();
  }

  public renderPending(renderer: PuzzleRenderer): void {
    renderer.renderPendingClearance(this);
  }
}

class House {
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

class Puzzle {
  public readonly cages: readonly Cage[];
  public readonly cells: readonly Cell[];
  public readonly columns: readonly House[];
  public readonly houses: readonly House[];
  public readonly rows: readonly House[];
  private pendingChanges: readonly CellChange[] = [];
  private readonly strategies: readonly Strategy[];

  public constructor(
    private readonly renderer: PuzzleRenderer,
    public readonly size: number,
    cagesRaw: readonly CageRaw[],
    public readonly hasOperators: boolean,
    public readonly title: string,
    public readonly meta: string,
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

    this.strategies = [
      new SingleCandidateStrategy(),
      new HiddenSingleStrategy(),
      ...Array.from({ length: size - MIN_NAKED_SET_SIZE }, (_, i) => new NakedSetStrategy(i + MIN_NAKED_SET_SIZE))
    ];
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

class ValueChange extends CellChange {
  public constructor(cell: Cell, public readonly value: number) {
    super(cell);
  }

  public applyToModel(): void {
    this.cell.setValue(this.value);
    this.cell.clearCandidates();
  }

  public renderPending(renderer: PuzzleRenderer): void {
    renderer.renderPendingValue(this);
  }
}

const BINARY_OP_SIZE = 2;
const CHAR_CODE_A = 65;
const MIN_NAKED_SET_SIZE = 2;

function buildAutoEliminateChanges(
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

function buildCageConstraintChanges(
  cages: readonly Cage[],
  hasOperators: boolean,
  gridSize: number
): CellChange[] {
  const valueSetters: CellValueSetter[] = [];
  const candidateChanges: CellChange[] = [];
  const narrowedCells = new Set<Cell>();
  for (const cage of cages) {
    if (cage.cells.length === 1) {
      const cellValue = cage.value ?? (cage.label ? parseInt(cage.label, 10) : undefined);
      if (cellValue !== undefined && !isNaN(cellValue)) {
        valueSetters.push({ cell: ensureNonNullable(cage.cells[0]), value: cellValue });
      }
      continue;
    }

    const cageValue = cage.value ?? (cage.label ? parseInt(cage.label, 10) : undefined);
    if (cageValue === undefined || isNaN(cageValue)) {
      continue;
    }

    let distinctSets: Set<string>;
    if (hasOperators && cage.operator) {
      distinctSets = computeCageValueSets(cageValue, cage.operator, cage.cells.length, gridSize);
    } else {
      distinctSets = new Set<string>();
      for (const op of ['+', '-', 'x', '/']) {
        for (const s of computeCageValueSets(cageValue, op, cage.cells.length, gridSize)) {
          distinctSets.add(s);
        }
      }
    }

    if (distinctSets.size !== 1) {
      continue;
    }

    const narrowedValues = ensureNonNullable([...distinctSets][0]).split(',').map(Number);
    if (narrowedValues.length >= gridSize) {
      continue;
    }

    for (const cell of cage.cells) {
      candidateChanges.push(new CandidatesChange(cell, narrowedValues));
      narrowedCells.add(cell);
    }
  }

  const valuedCells = new Set(valueSetters.map((s) => s.cell));
  const autoChanges = buildAutoEliminateChanges(valueSetters);
  const filteredChanges = autoChanges.filter(
    (change) => !(change instanceof CandidatesStrikethrough
      && (narrowedCells.has(change.cell) || valuedCells.has(change.cell)))
  );
  return [...filteredChanges, ...candidateChanges];
}

function cellRefA1(r: number, c: number): string {
  return String.fromCharCode(CHAR_CODE_A + c) + String(r + 1);
}

function computeCageValueSets(
  value: number,
  operator: string,
  numCells: number,
  gridSize: number
): Set<string> {
  const distinctSets = new Set<string>();

  function search(tuple: number[], depth: number): void {
    if (depth === numCells) {
      if (evaluateTuple(tuple, operator) === value) {
        const valueSet = [...new Set(tuple)].sort((a, b) => a - b).join(',');
        distinctSets.add(valueSet);
      }
      return;
    }
    for (let v = 1; v <= gridSize; v++) {
      tuple.push(v);
      search(tuple, depth + 1);
      tuple.pop();
    }
  }

  search([], 0);
  return distinctSets;
}

function computeGridBoundaries(cages: readonly CageRaw[], gridDimension: number): GridBoundaries {
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

function evaluateTuple(tuple: readonly number[], operator: string): null | number {
  if (tuple.length === 0) {
    return null;
  }
  if (tuple.length === 1) {
    return ensureNonNullable(tuple[0]);
  }

  const a = ensureNonNullable(tuple[0]);
  const b = ensureNonNullable(tuple[1]);

  switch (operator) {
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
    columnId: ensureNonNullable(groups['col']).charCodeAt(0) - CHAR_CODE_A + 1,
    rowId: parseInt(ensureNonNullable(groups['row']), 10)
  };
}

function parseOperation(text: string, cellCount: number): CellOperation {
  if (text.toLowerCase() === 'x') {
    return { type: 'clear' };
  }

  if (text.startsWith('=')) {
    const valueStr = text.substring(1);
    if (!/^[1-9]$/.test(valueStr)) {
      throw new Error('=N expects a single digit 1-9');
    }
    if (cellCount > 1) {
      throw new Error('=N can only be used with a single cell');
    }
    return { type: 'value', value: parseInt(valueStr, 10) };
  }

  if (text.startsWith('-')) {
    const valuesStr = text.substring(1);
    if (!/^[1-9]+$/.test(valuesStr)) {
      throw new Error('-digits expects digits 1-9');
    }
    const values = Array.from(valuesStr, (ch) => parseInt(ch, 10));
    return { type: 'strikethrough', values };
  }

  if (!/^[1-9]+$/.test(text)) {
    throw new Error(`Expected digits 1-9: ${text}`);
  }
  const values = Array.from(text, (ch) => parseInt(ch, 10));
  return { type: 'candidates', values };
}
