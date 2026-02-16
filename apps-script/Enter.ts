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

function enter(cellsText: string, opsText: string): void {
  const state = getPuzzleState();
  const n = state.size;
  const slide = getCurrentSlide();

  // Parse cell references
  const tokens = cellsText.trim().split(/\s+/);
  const cellRefs: string[] = [];
  for (const token of tokens) {
    if (token.startsWith('@')) {
      // Expand cage
      const anchor = parseCellRef(token.substring(1));
      const cageCells = getCageCells(anchor.r, anchor.c);
      for (const cell of cageCells) {
        cellRefs.push(cell);
      }
    } else {
      const cell = parseCellRef(token);
      cellRefs.push(cellRefA1(cell.r, cell.c));
    }
  }

  if (cellRefs.length === 0) {
    throw new Error('No cells specified');
  }

  // Parse operations
  const ops = parseOperations(opsText, cellRefs.length);

  // Apply operations
  for (const cellRef of cellRefs) {
    applyOps(slide, cellRef, ops, n);
  }
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
