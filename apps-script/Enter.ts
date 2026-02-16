/**
 * Enter.ts -- Enter command: apply value/candidate changes to cells.
 *
 * Operations:
 *   =N     -- set final value (single cell only)
 *   456    -- set candidates (digits, no prefix)
 *   -789   -- cross out candidates (strikethrough)
 */

// ── Operation types (discriminated union) ──────────────────────────────────

interface CandidatesOp {
  digits: string;
  type: 'candidates';
}

type CellOperation = CandidatesOp | StrikeOp | ValueOp;

interface StrikeOp {
  digits: string;
  type: 'strike';
}

interface ValueOp {
  digit: string;
  type: 'value';
}

// ── Dialog ─────────────────────────────────────────────────────────────────

function clearShapeText(slide: GoogleAppsScript.Slides.Slide, title: string): void {
  const shape = getShapeByTitle(slide, title);
  if (shape) {
    shape.getText().setText(' ');
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

// ── Main entry ─────────────────────────────────────────────────────────────

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

// ── Parse operations ───────────────────────────────────────────────────────

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

// ── Apply operations ───────────────────────────────────────────────────────

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

/** Read the current value digit from a VALUE shape, or null if empty. */
function getCellValue(
  slide: GoogleAppsScript.Slides.Slide,
  cellRef: string
): string | null {
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

function parseOperations(text: string, cellCount: number): CellOperation {
  const trimmed = text.trim();
  if (!trimmed) {
    throw new Error('No operations specified');
  }

  if (trimmed.startsWith('=')) {
    const digit = trimmed.substring(1).trim();
    if (!/^[1-9]$/.test(digit)) {
      throw new Error('=N expects a single digit 1-9');
    }
    if (cellCount > 1) {
      throw new Error('=N can only be used with a single cell');
    }
    return { digit, type: 'value' };
  }

  if (trimmed.startsWith('-')) {
    const digits = trimmed.substring(1).trim();
    if (!/^[1-9]+$/.test(digits)) {
      throw new Error('-digits expects digits 1-9');
    }
    return { digits, type: 'strike' };
  }

  // Set candidates (plain digits or +digits)
  let raw = trimmed;
  if (raw.startsWith('+')) {
    raw = raw.substring(1);
  }
  raw = raw.trim();
  if (!/^[1-9]+$/.test(raw)) {
    throw new Error('Expected digits 1-9');
  }
  return { digits: raw, type: 'candidates' };
}

function showEnterDialog(): void {
  if (PropertiesService.getDocumentProperties().getProperty('mathdokuInitialized') !== 'true') {
    SlidesApp.getUi().alert('Please run Mathdoku > Init first.');
    return;
  }
  const html = HtmlService.createHtmlOutputFromFile('EnterDialog')
    .setWidth(340)
    .setHeight(210);
  SlidesApp.getUi().showModalDialog(html, 'Edit Cell');
}
