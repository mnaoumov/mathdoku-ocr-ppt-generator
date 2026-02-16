/**
 * Enter.ts -- Enter command: apply value/candidate changes to cells.
 *
 * Unified format:  cell:op  (or multiple commands separated by spaces)
 *   A1:=1          -- set final value
 *   A1:234         -- set candidates
 *   A1:-567        -- strikethrough candidates
 *   A1:x           -- clear cell (value + candidates)
 *   (B2 C3):234   -- multiple cells
 *   @D4:-567       -- expand cage
 */

// ── Operation types (discriminated union) ──────────────────────────────────

interface CandidatesOp {
  digits: string;
  type: 'candidates';
}

type CellOperation = CandidatesOp | ClearOp | StrikeOp | ValueOp;

interface ClearOp {
  type: 'clear';
}

interface EnterCommand {
  cellRefs: string[];
  operation: CellOperation;
}

interface StrikeOp {
  digits: string;
  type: 'strike';
}

interface ValueOp {
  digit: string;
  type: 'value';
}

// ── Functions ─────────────────────────────────────────────────────────────

function applyGreenCandidates(
  slide: GoogleAppsScript.Slides.Slide,
  cellRef: string,
  digits: string,
  gridSize: number
): void {
  const shape = getShapeByTitle(slide, `CANDIDATES_${cellRef}`);
  if (!shape) {
    throw new Error(`Candidates shape not found: CANDIDATES_${cellRef}`);
  }

  const formatted = formatCandidates(digits, gridSize);
  const textRange = shape.getText();
  textRange.setText(formatted);
  textRange.getTextStyle()
    .setFontFamily(CANDIDATES_FONT)
    .setForegroundColor(GREEN);

  // Clear final value when setting candidates
  clearShapeText(slide, `VALUE_${cellRef}`);
}

function applyGreenStrikethrough(
  slide: GoogleAppsScript.Slides.Slide,
  cellRef: string,
  digits: string
): void {
  const shape = getShapeByTitle(slide, `CANDIDATES_${cellRef}`);
  if (!shape) {
    throw new Error(`Candidates shape not found: CANDIDATES_${cellRef}`);
  }

  const textRange = shape.getText();
  const fullText = textRange.asString();

  for (let i = 0; i < fullText.length; i++) {
    const ch = fullText.charAt(i);
    if (digits.includes(ch)) {
      textRange.getRange(i, i + 1).getTextStyle()
        .setForegroundColor(GREEN)
        .setStrikethrough(true);
    }
  }
}

function applyGreenValue(
  slide: GoogleAppsScript.Slides.Slide,
  cellRef: string,
  digit: string
): void {
  const shape = getShapeByTitle(slide, `VALUE_${cellRef}`);
  if (!shape) {
    throw new Error(`Value shape not found: VALUE_${cellRef}`);
  }

  const textRange = shape.getText();
  textRange.setText(digit);
  textRange.getTextStyle()
    .setFontFamily('Segoe UI')
    .setBold(true)
    .setForegroundColor(GREEN);

  // Clear candidates when setting final value
  clearShapeText(slide, `CANDIDATES_${cellRef}`);
}

function applyOps(
  slide: GoogleAppsScript.Slides.Slide,
  cellRef: string,
  ops: CellOperation,
  gridSize: number
): void {
  switch (ops.type) {
    case 'candidates':
      applyGreenCandidates(slide, cellRef, ops.digits, gridSize);
      break;
    case 'clear':
      clearShapeText(slide, `VALUE_${cellRef}`);
      clearShapeText(slide, `CANDIDATES_${cellRef}`);
      break;
    case 'strike':
      applyGreenStrikethrough(slide, cellRef, ops.digits);
      break;
    case 'value':
      applyGreenValue(slide, cellRef, ops.digit);
      break;
    default: {
      const exhaustive: never = ops;
      throw new Error(`Unknown operation type: ${String(exhaustive)}`);
    }
  }
}

function autoEliminate(
  slide: GoogleAppsScript.Slides.Slide,
  cellRefs: string[],
  operation: CellOperation,
  gridSize: number
): void {
  if (operation.type === 'value') {
    // Setting a value: strikethrough that digit in all other cells in the same row/column
    const ref = parseCellRef(ensureNonNullable(cellRefs[0]));
    const digit = operation.digit;
    for (let i = 0; i < gridSize; i++) {
      if (i !== ref.c) {
        applyGreenStrikethrough(slide, cellRefA1(ref.r, i), digit);
      }
      if (i !== ref.r) {
        applyGreenStrikethrough(slide, cellRefA1(i, ref.c), digit);
      }
    }
  } else if (operation.type === 'candidates') {
    // Adding candidates: strikethrough any digits that conflict with
    // existing values in the same row or column
    for (const cellRef of cellRefs) {
      const ref = parseCellRef(cellRef);
      const conflicting = getRowColValues(slide, ref.r, ref.c, gridSize);
      if (conflicting.length > 0) {
        applyGreenStrikethrough(slide, cellRef, conflicting.join(''));
      }
    }
  }
}

function clearShapeText(slide: GoogleAppsScript.Slides.Slide, title: string): void {
  const shape = getShapeByTitle(slide, title);
  if (shape) {
    shape.getText().setText(' ');
  }
}

/**
 * Enter command: parse and apply one or more cell:operation commands.
 *
 * Format:  A1:=1  A1:234  A1:-567  (B2 C3):234  @D4:-567
 * Multiple commands in one line:  A1:=1 (B2 C3):234 @D4:-567
 */
function enter(input: string): void {
  const state = getPuzzleState();
  const gridSize = state.size;
  const slide = getCurrentSlide();

  const commands = parseInput(input);
  for (const cmd of commands) {
    for (const cellRef of cmd.cellRefs) {
      applyOps(slide, cellRef, cmd.operation, gridSize);
    }
    autoEliminate(slide, cmd.cellRefs, cmd.operation, gridSize);
  }
}

/** Read the current value digit from a VALUE shape, or null if empty. */
function getCellValue(
  slide: GoogleAppsScript.Slides.Slide,
  cellRef: string
): null | string {
  const shape = getShapeByTitle(slide, `VALUE_${cellRef}`);
  if (!shape) {
    return null;
  }
  const text = shape.getText().asString().replace(/\n$/, '').trim();
  if (/^[1-9]$/.test(text)) {
    return text;
  }
  return null;
}

/** Read confirmed values from all other cells in the same row and column. */
function getRowColValues(
  slide: GoogleAppsScript.Slides.Slide,
  r: number,
  c: number,
  gridSize: number
): string[] {
  const values: string[] = [];
  for (let i = 0; i < gridSize; i++) {
    if (i !== c) {
      const v = getCellValue(slide, cellRefA1(r, i));
      if (v) {
        values.push(v);
      }
    }
    if (i !== r) {
      const v = getCellValue(slide, cellRefA1(i, c));
      if (v) {
        values.push(v);
      }
    }
  }
  return values;
}

/** Expand a cell-part token into resolved A1-style cell references. */
function parseCellPart(cellPart: string): string[] {
  let inner = cellPart;
  if (inner.startsWith('(') && inner.endsWith(')')) {
    inner = inner.substring(1, inner.length - 1);
  }

  const cellRefs: string[] = [];
  for (const token of inner.trim().split(/\s+/)) {
    if (token.startsWith('@')) {
      const anchor = parseCellRef(token.substring(1));
      for (const cell of getCageCells(anchor.r, anchor.c)) {
        cellRefs.push(cell);
      }
    } else {
      const cell = parseCellRef(token);
      cellRefs.push(cellRefA1(cell.r, cell.c));
    }
  }
  return cellRefs;
}

/**
 * Parse unified input into commands.
 *
 * Tokenizes by matching `cellPart:opPart` groups where cellPart is
 * `(...)`, `@REF`, or `REF`, and spaces inside parentheses are preserved.
 */
function parseInput(input: string): EnterCommand[] {
  const trimmed = input.trim();
  if (!trimmed) {
    throw new Error('No commands specified');
  }

  // Match: (cells):op | @cell:op | cell:op
  const pattern = /(?:\([^)]+\)|@?[A-Za-z]\d+):[^\s]+/g;
  const matches = trimmed.match(pattern);
  if (!matches || matches.length === 0) {
    throw new Error('Invalid format. Expected: A1:=1, A1:234, (B2 C3):-567, @D4:234');
  }

  const commands: EnterCommand[] = [];
  for (const match of matches) {
    // Find the colon that separates cell part from operation
    const colonIdx = match.startsWith('(')
      ? match.indexOf(':', match.indexOf(')'))
      : match.indexOf(':');
    const cellPart = match.substring(0, colonIdx);
    const opPart = match.substring(colonIdx + 1);
    const cellRefs = parseCellPart(cellPart);
    const operation = parseOperation(opPart, cellRefs.length);
    commands.push({ cellRefs, operation });
  }
  return commands;
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

function showEnterDialog(): void {
  if (PropertiesService.getDocumentProperties().getProperty('mathdokuInitialized') !== 'true') {
    SlidesApp.getUi().alert('Please run Mathdoku > Init first.');
    return;
  }
  const html = HtmlService.createHtmlOutputFromFile('EnterDialog')
    .setWidth(400)
    .setHeight(140);
  SlidesApp.getUi().showModelessDialog(html, 'Edit Cell');
}
