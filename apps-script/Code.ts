/**
 * Mathdoku Solver -- Google Slides bound script.
 *
 * Main entry: Code.ts (menu, helpers, importPuzzle)
 */

// ── Assert helpers ──────────────────────────────────────────────────────────

interface Cage {
  cells: string[];
  label?: string;
  op?: string;
  value?: number;
}

interface CageProfile {
  boxHFrac: number;
  boxWFrac: number;
  font: number;
  insetXFrac: number;
  insetYFrac: number;
}

// ── Interfaces ─────────────────────────────────────────────────────────────

interface CandidatesProfile {
  digitMargin: number;
  font: number;
  hFrac: number;
  wFrac: number;
  xFrac: number;
  yFrac: number;
}

interface CellRef {
  c: number;
  r: number;
}

interface GridBoundaries {
  horizontalBounds: boolean[][];
  verticalBounds: boolean[][];
}

interface LayoutProfile {
  axisFont: number;
  axisLabelH: number;
  axisLabelW: number;
  axisSideOffset: number;
  axisTopOffset: number;
  cage: CageProfile;
  candidates: CandidatesProfile;
  gridLeftIn: number;
  gridSizeIn: number;
  gridTopIn: number;
  metaSz: number;
  solve: SolveProfile;
  thickPt: number;
  thinPt: number;
  titleHIn: number;
  titleSz: number;
  value: ValueProfile;
}

interface PuzzleJson {
  cages: Cage[];
  meta?: string;
  operations?: boolean;
  size: number;
  title?: string;
}

interface PuzzleState {
  cages: Cage[];
  operations: boolean;
  size: number;
}

interface SolveProfile {
  colGapIn: number;
  cols: number;
  colWIn: number;
  font: number;
  leftIn: number;
}

interface ValueProfile {
  font: number;
  hFrac: number;
  yFrac: number;
}

function assertNonNullable<T>(value: T, errorOrMessage?: Error | string): asserts value is NonNullable<T> {
  if (value !== null && value !== undefined) {
    return;
  }
  errorOrMessage ??= value === null ? 'Value is null' : 'Value is undefined';
  const error = typeof errorOrMessage === 'string' ? new Error(errorOrMessage) : errorOrMessage;
  throw error;
}

function ensureNonNullable<T>(value: T, errorOrMessage?: Error | string): NonNullable<T> {
  assertNonNullable(value, errorOrMessage);
  return value;
}

// ── Colors ──────────────────────────────────────────────────────────────────

const AXIS_LABEL_MAGENTA = '#C800C8';
const CAGE_LABEL_BLUE = '#3232C8';
const VALUE_GRAY = '#3C414B';
const CAND_DARK_RED = '#8B0000';
const CANDIDATES_FONT = 'Consolas';
const THIN_GRAY = '#AAAAAA';
const BLACK = '#000000';
const GREEN = '#00B050';
const LIGHT_GRAY_BORDER = '#C8C8C8';
const FOOTER_COLOR = '#6E7887';

const FOOTER_TEXT = '@mnaoumov';

// ── Slide dimensions ────────────────────────────────────────────────────────
// Layout matches CLAUDE.md "Slide layout spec (pixel-perfect reference)".
// All positions/sizes from LAYOUT_PROFILES; convert inches to pt (1 in = 72 pt), round to integer.
// Reference size 960×540 pt; scale only when page is smaller or content would overflow.

const REF_W = 960;
const REF_H = 540;

// Google Slides text boxes have ~0.05 in default top padding that cannot be zeroed via API.
// Compensate by shifting boxes up. See CLAUDE.md "Google Slides API quirks" #3.
const TEXT_BOX_TOP_PAD_PT = 0.05 * 72;

// ── Layout profiles ─────────────────────────────────────────────────────────

const LAYOUT_PROFILES: Record<number, LayoutProfile> = {
  4: {
    axisFont: 24, axisLabelH: 0.34, axisLabelW: 0.30,
    axisSideOffset: 0.36, axisTopOffset: 0.42, cage: { boxHFrac: 0.35, boxWFrac: 0.65, font: 28, insetXFrac: 0.07, insetYFrac: 0.05 },
    candidates: { digitMargin: 12, font: 22, hFrac: 0.60, wFrac: 0.80, xFrac: 0.15, yFrac: 0.38 }, gridLeftIn: 0.65,
    gridSizeIn: 4.75, gridTopIn: 1.35, metaSz: 20,
    solve: { colGapIn: 0.25, cols: 2, colWIn: 3.25, font: 16, leftIn: 6.20 }, thickPt: 5.0,
    thinPt: 1.0,
    titleHIn: 0.85,
    titleSz: 30,
    value: { font: 52, hFrac: 0.70, yFrac: 0.30 }
  },
  5: {
    axisFont: 26, axisLabelH: 0.37, axisLabelW: 0.32,
    axisSideOffset: 0.38, axisTopOffset: 0.45, cage: { boxHFrac: 0.33, boxWFrac: 0.65, font: 24, insetXFrac: 0.07, insetYFrac: 0.05 },
    candidates: { digitMargin: 7, font: 20, hFrac: 0.62, wFrac: 0.88, xFrac: 0.05, yFrac: 0.36 }, gridLeftIn: 0.65,
    gridSizeIn: 5.20, gridTopIn: 1.25, metaSz: 18,
    solve: { colGapIn: 0.25, cols: 2, colWIn: 3.10, font: 16, leftIn: 6.55 }, thickPt: 5.0,
    thinPt: 1.0,
    titleHIn: 0.70,
    titleSz: 26,
    value: { font: 44, hFrac: 0.72, yFrac: 0.28 }
  },
  6: {
    axisFont: 22, axisLabelH: 0.32, axisLabelW: 0.28,
    axisSideOffset: 0.34, axisTopOffset: 0.41, cage: { boxHFrac: 0.30, boxWFrac: 0.65, font: 22, insetXFrac: 0.07, insetYFrac: 0.05 },
    candidates: { digitMargin: 5, font: 19, hFrac: 0.65, wFrac: 0.86, xFrac: 0.07, yFrac: 0.33 }, gridLeftIn: 0.65,
    gridSizeIn: 5.70, gridTopIn: 1.15, metaSz: 16,
    solve: { colGapIn: 0.25, cols: 2, colWIn: 2.95, font: 16, leftIn: 6.85 }, thickPt: 6.5,
    thinPt: 1.0,
    titleHIn: 0.65,
    titleSz: 24,
    value: { font: 38, hFrac: 0.75, yFrac: 0.25 }
  },
  7: {
    axisFont: 28, axisLabelH: 0.40, axisLabelW: 0.35,
    axisSideOffset: 0.42, axisTopOffset: 0.49, cage: { boxHFrac: 0.28, boxWFrac: 0.65, font: 20, insetXFrac: 0.07, insetYFrac: 0.05 },
    candidates: { digitMargin: 1, font: 18, hFrac: 0.67, wFrac: 0.84, xFrac: 0.08, yFrac: 0.31 }, gridLeftIn: 0.65,
    gridSizeIn: 6.05, gridTopIn: 1.10, metaSz: 14,
    solve: { colGapIn: 0.25, cols: 2, colWIn: 2.85, font: 16, leftIn: 7.05 }, thickPt: 6.5,
    thinPt: 1.0,
    titleHIn: 0.55,
    titleSz: 22,
    value: { font: 32, hFrac: 0.77, yFrac: 0.23 }
  },
  8: {
    axisFont: 28, axisLabelH: 0.40, axisLabelW: 0.35,
    axisSideOffset: 0.42, axisTopOffset: 0.49, cage: { boxHFrac: 0.26, boxWFrac: 0.65, font: 18, insetXFrac: 0.07, insetYFrac: 0.05 },
    candidates: { digitMargin: 1, font: 17, hFrac: 0.69, wFrac: 0.84, xFrac: 0.08, yFrac: 0.29 }, gridLeftIn: 0.65,
    gridSizeIn: 6.20, gridTopIn: 1.10, metaSz: 14,
    solve: { colGapIn: 0.25, cols: 2, colWIn: 2.80, font: 16, leftIn: 7.15 }, thickPt: 6.5,
    thinPt: 1.0,
    titleHIn: 0.55,
    titleSz: 22,
    value: { font: 30, hFrac: 0.78, yFrac: 0.22 }
  },
  9: {
    axisFont: 28, axisLabelH: 0.40, axisLabelW: 0.35,
    axisSideOffset: 0.42, axisTopOffset: 0.49, cage: { boxHFrac: 0.24, boxWFrac: 0.70, font: 16, insetXFrac: 0.07, insetYFrac: 0.05 },
    candidates: { digitMargin: 0, font: 14, hFrac: 0.71, wFrac: 0.88, xFrac: 0.09, yFrac: 0.27 }, gridLeftIn: 0.65,
    gridSizeIn: 6.30, gridTopIn: 1.10, metaSz: 14,
    solve: { colGapIn: 0.25, cols: 2, colWIn: 2.75, font: 16, leftIn: 7.25 }, thickPt: 6.5,
    thinPt: 1.0,
    titleHIn: 0.55,
    titleSz: 22,
    value: { font: 28, hFrac: 0.80, yFrac: 0.20 }
  }
};

// ── Menu ────────────────────────────────────────────────────────────────────

function cellRefA1(r: number, c: number): string {
  return String.fromCharCode(65 + c) + String(r + 1);
}

/**
 * Extract hex color string from a Slides Color object for comparison.
 * Returns uppercase hex like "#00B050".
 */
function colorToHex(color: GoogleAppsScript.Slides.Color): string {
  return color.asRgbColor().asHexString().toUpperCase();
}

function computeGridBoundaries(cages: Cage[], gridDimension: number): GridBoundaries {
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

// ── Helpers ─────────────────────────────────────────────────────────────────

function drawAxisLabels(
  slide: GoogleAppsScript.Slides.Slide,
  gridLeft: number,
  gridTop: number,
  cellWidth: number,
  gridDimension: number,
  profile: LayoutProfile
): void {
  const axisFont = profile.axisFont;
  const labelHeight = pt(in2pt(profile.axisLabelH));
  const labelWidth = pt(in2pt(profile.axisLabelW));
  const topOffset = in2pt(profile.axisTopOffset);
  const sideOffset = in2pt(profile.axisSideOffset);
  const topY = pt(gridTop - topOffset - TEXT_BOX_TOP_PAD_PT);
  const sideX = pt(gridLeft - sideOffset);

  // Column labels (A, B, C, ...): use running x so alignment doesn't drift from rounding
  let columnLeft = gridLeft;
  for (let c = 0; c < gridDimension; c++) {
    const boxWidth = pt(cellWidth);
    const box = slide.insertTextBox(String.fromCharCode(65 + c), pt(columnLeft), topY, boxWidth, labelHeight);
    box.getText().getTextStyle()
      .setFontFamily('Segoe UI').setFontSize(axisFont).setBold(true).setForegroundColor(AXIS_LABEL_MAGENTA);
    box.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    box.getBorder().setTransparent();
    columnLeft += cellWidth;
  }

  // Row labels (1, 2, 3, ...)
  for (let r = 0; r < gridDimension; r++) {
    const y = pt(gridTop + r * cellWidth + (cellWidth - labelHeight) / 2);
    const box = slide.insertTextBox(String(r + 1), sideX, y, labelWidth, labelHeight);
    box.getText().getTextStyle()
      .setFontFamily('Segoe UI').setFontSize(axisFont).setBold(true).setForegroundColor(AXIS_LABEL_MAGENTA);
    box.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    box.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
    box.getBorder().setTransparent();
  }
}

function drawCageBoundaries(
  slide: GoogleAppsScript.Slides.Slide,
  gridLeft: number,
  gridTop: number,
  gridSize: number,
  gridDimension: number,
  thickPt: number,
  verticalBounds: boolean[][],
  horizontalBounds: boolean[][]
): void {
  const cellWidth = gridSize / gridDimension;
  const inset = thickPt / 2;

  // Vertical runs
  for (let c = 1; c < gridDimension; c++) {
    let startRow = 0;
    while (startRow < gridDimension) {
      if (!ensureNonNullable(verticalBounds[startRow])[c - 1]) {
        startRow++;
        continue;
      }
      let endRow = startRow;
      while (endRow + 1 < gridDimension && ensureNonNullable(verticalBounds[endRow + 1])[c - 1]) {
        endRow++;
      }
      const x = gridLeft + c * cellWidth;
      let y1 = gridTop + startRow * cellWidth;
      let y2 = gridTop + (endRow + 1) * cellWidth;
      if (startRow === 0) {
        y1 += inset;
      }
      if (endRow === gridDimension - 1) {
        y2 -= inset;
      }
      drawThickRect(slide, x - thickPt / 2, y1, thickPt, y2 - y1);
      startRow = endRow + 1;
    }
  }

  // Horizontal runs
  for (let r = 1; r < gridDimension; r++) {
    let startCol = 0;
    while (startCol < gridDimension) {
      if (!ensureNonNullable(horizontalBounds[r - 1])[startCol]) {
        startCol++;
        continue;
      }
      let endCol = startCol;
      while (endCol + 1 < gridDimension && ensureNonNullable(horizontalBounds[r - 1])[endCol + 1]) {
        endCol++;
      }
      const y = gridTop + r * cellWidth;
      let x1 = gridLeft + startCol * cellWidth;
      let x2 = gridLeft + (endCol + 1) * cellWidth;
      if (startCol === 0) {
        x1 += inset;
      }
      if (endCol === gridDimension - 1) {
        x2 -= inset;
      }
      drawThickRect(slide, x1, y - thickPt / 2, x2 - x1, thickPt);
      startCol = endCol + 1;
    }
  }
}

function drawCageLabels(
  slide: GoogleAppsScript.Slides.Slide,
  gridLeft: number,
  gridTop: number,
  cellWidth: number,
  cages: Cage[],
  operations: boolean,
  profile: LayoutProfile
): void {
  const cageProfile = profile.cage;
  const insetX = cageProfile.insetXFrac * cellWidth;
  const insetY = cageProfile.insetYFrac * cellWidth;
  const labelBoxWidth = cageProfile.boxWFrac * cellWidth;
  // Cap box height so it never extends into the candidates region
  const labelBoxHeight = Math.min(
    cageProfile.boxHFrac * cellWidth,
    profile.candidates.yFrac * cellWidth - insetY
  );

  for (let i = 0; i < cages.length; i++) {
    const cage = ensureNonNullable(cages[i]);

    // Find geometric top-left cell (smallest row, then smallest column)
    const parsed = cage.cells.map((ref) => ({ ref, ...parseCellRef(ref) }));
    parsed.sort((a, b) => a.r === b.r ? a.c - b.c : a.r - b.r);
    const topLeftCell = ensureNonNullable(parsed[0]);
    const topLeftCellRef = topLeftCell.ref;

    // Build label
    let label = cage.label ?? '';
    if (!label && cage.value !== undefined) {
      label = operations && cage.cells.length > 1 && cage.op
        ? String(cage.value) + opSymbol(cage.op)
        : String(cage.value);
    }

    const x = pt(gridLeft + topLeftCell.c * cellWidth + insetX);
    const y = pt(gridTop + topLeftCell.r * cellWidth + insetY - TEXT_BOX_TOP_PAD_PT);
    // Size font for usable area (box minus default top padding from Google Slides)
    const usableHeight = Math.max(7, labelBoxHeight - TEXT_BOX_TOP_PAD_PT);
    const actualFont = fitFontSize(label, cageProfile.font, labelBoxWidth / 72, usableHeight / 72);
    const box = slide.insertTextBox(label, x, y, pt(labelBoxWidth), pt(labelBoxHeight));
    box.setTitle(`CAGE_${String(i)}_${topLeftCellRef}`);
    box.getText().getTextStyle()
      .setFontFamily('Segoe UI').setFontSize(actualFont).setBold(true).setForegroundColor(CAGE_LABEL_BLUE);
    box.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
    box.setContentAlignment(SlidesApp.ContentAlignment.TOP);
  }
}

function drawJoinSquares(
  slide: GoogleAppsScript.Slides.Slide,
  gridLeft: number,
  gridTop: number,
  gridSize: number,
  gridDimension: number,
  thickPt: number,
  verticalBounds: boolean[][],
  horizontalBounds: boolean[][]
): void {
  const cellWidth = gridSize / gridDimension;
  const clamp = (value: number, lo: number, hi: number): number =>
    Math.max(lo, Math.min(hi, value));

  // Draw join squares at every interior vertex where any cage boundary touches (spec: "if any of ... is true")
  for (let vertexRow = 1; vertexRow < gridDimension; vertexRow++) {
    for (let vertexCol = 1; vertexCol < gridDimension; vertexCol++) {
      const boundAbove = ensureNonNullable(verticalBounds[vertexRow - 1])[vertexCol - 1];
      const boundBelow = ensureNonNullable(verticalBounds[vertexRow])[vertexCol - 1];
      const boundLeft = ensureNonNullable(horizontalBounds[vertexRow - 1])[vertexCol - 1];
      const boundRight = ensureNonNullable(horizontalBounds[vertexRow - 1])[vertexCol];
      if (!boundAbove && !boundBelow && !boundLeft && !boundRight) {
        continue;
      }
      const x = gridLeft + vertexCol * cellWidth;
      const y = gridTop + vertexRow * cellWidth;
      const left = clamp(x - thickPt / 2, gridLeft, gridLeft + gridSize - thickPt);
      const top = clamp(y - thickPt / 2, gridTop, gridTop + gridSize - thickPt);
      drawThickRect(slide, left, top, pt(thickPt), pt(thickPt));
    }
  }
}

function drawOuterBorder(
  slide: GoogleAppsScript.Slides.Slide,
  gridLeft: number,
  gridTop: number,
  gridSize: number,
  thickPt: number
): void {
  const halfThick = thickPt / 2;
  // Top
  drawThickRect(slide, gridLeft - halfThick, gridTop - halfThick, gridSize + thickPt, thickPt);
  // Bottom
  drawThickRect(slide, gridLeft - halfThick, gridTop + gridSize - halfThick, gridSize + thickPt, thickPt);
  // Left
  drawThickRect(slide, gridLeft - halfThick, gridTop + halfThick, thickPt, Math.max(0, gridSize - thickPt));
  // Right
  drawThickRect(slide, gridLeft + gridSize - halfThick, gridTop + halfThick, thickPt, Math.max(0, gridSize - thickPt));
}

function drawThickRect(
  slide: GoogleAppsScript.Slides.Slide,
  left: number,
  top: number,
  width: number,
  height: number
): void {
  const w = pt(width);
  const h = pt(height);
  if (w <= 0 || h <= 0) {
    return;
  }
  const rect = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, pt(left), pt(top), w, h);
  rect.getFill().setSolidFill(BLACK);
  rect.getBorder().setTransparent();
}

function drawThinGrid(
  slide: GoogleAppsScript.Slides.Slide,
  gridLeft: number,
  gridTop: number,
  gridSize: number,
  gridDimension: number,
  thinWidth: number,
  verticalBounds: boolean[][],
  horizontalBounds: boolean[][]
): void {
  const cellWidth = gridSize / gridDimension;
  const halfThinWidth = thinWidth / 2;

  // Vertical thin lines: between columns c-1 and c, for each row
  for (let c = 1; c < gridDimension; c++) {
    for (let r = 0; r < gridDimension; r++) {
      if (ensureNonNullable(verticalBounds[r])[c - 1]) {
        continue;
      }
      const x = gridLeft + c * cellWidth;
      const y = gridTop + r * cellWidth;
      const rect = slide.insertShape(SlidesApp.ShapeType.RECTANGLE,
        pt(x - halfThinWidth), pt(y), pt(thinWidth), pt(cellWidth));
      rect.getFill().setSolidFill(THIN_GRAY);
      rect.getBorder().setTransparent();
    }
  }

  // Horizontal thin lines: between rows r-1 and r, for each column
  for (let r = 1; r < gridDimension; r++) {
    for (let c = 0; c < gridDimension; c++) {
      if (ensureNonNullable(horizontalBounds[r - 1])[c]) {
        continue;
      }
      const x = gridLeft + c * cellWidth;
      const y = gridTop + r * cellWidth;
      const rect = slide.insertShape(SlidesApp.ShapeType.RECTANGLE,
        pt(x), pt(y - halfThinWidth), pt(cellWidth), pt(thinWidth));
      rect.getFill().setSolidFill(THIN_GRAY);
      rect.getBorder().setTransparent();
    }
  }
}

function fitFontSize(text: string, basePt: number, boxWidthIn: number, boxHeightIn: number): number {
  const trimmedText = text.trim();
  if (!trimmedText) {
    return basePt;
  }
  const maxWidthPt = Math.max(1, boxWidthIn * 72 - 2);
  const maxHeightPt = Math.max(1, boxHeightIn * 72 - 1);
  const charCount = Math.max(1, trimmedText.length);
  const widthBased = Math.floor(maxWidthPt / (0.60 * charCount));
  const heightBased = Math.floor(maxHeightPt / 1.15);
  return Math.max(7, Math.min(basePt, widthBased, heightBased));
}

function formatCandidates(digits: string, gridSize: number): string {
  // Fixed-position layout: each digit at its slot, missing digits replaced with space.
  // No inter-digit spaces — original used Font.Spacing for visual separation,
  // which Google Slides API doesn't support (see CLAUDE.md deviations).
  const firstRowCount = Math.ceil(gridSize / 2);
  let line1 = '';
  let line2 = '';
  for (let d = 1; d <= gridSize; d++) {
    const ch = digits.includes(String(d)) ? String(d) : ' ';
    if (d <= firstRowCount) {
      line1 += ch;
    } else {
      line2 += ch;
    }
  }
  return `${line1}\n${line2}`;
}

function getCageCells(r: number, c: number): string[] {
  const state = getPuzzleState();
  const ref = cellRefA1(r, c);
  for (const cage of state.cages) {
    if (cage.cells.includes(ref)) {
      return cage.cells;
    }
  }
  throw new Error(`Cell ${ref} not found in any cage`);
}

function getCurrentSlide(): GoogleAppsScript.Slides.Slide {
  const pres = SlidesApp.getActivePresentation();
  const selection = pres.getSelection();
  const page = selection.getCurrentPage() as GoogleAppsScript.Slides.Page | null;
  if (page) {
    return page.asSlide();
  }
  return getLastSlide();
}

function getLastSlide(): GoogleAppsScript.Slides.Slide {
  const slides = SlidesApp.getActivePresentation().getSlides();
  const last = slides[slides.length - 1];
  if (!last) {
    throw new Error('Presentation has no slides');
  }
  return last;
}

function getPuzzleState(): PuzzleState {
  const raw = PropertiesService.getDocumentProperties().getProperty('mathdokuState');
  if (!raw) {
    throw new Error('No puzzle state found in document properties');
  }
  return JSON.parse(raw) as PuzzleState;
}

// ── Import puzzle (grid generation) ─────────────────────────────────────────

function getShapeByTitle(slide: GoogleAppsScript.Slides.Slide, title: string): GoogleAppsScript.Slides.Shape | null {
  for (const el of slide.getPageElements()) {
    if (el.getTitle() === title) {
      return el.asShape();
    }
  }
  return null;
}

function importPuzzle(puzzleJson: PuzzleJson | string, presId?: string): void {
  const puzzle: PuzzleJson = (typeof puzzleJson === 'string') ? JSON.parse(puzzleJson) as PuzzleJson : puzzleJson;
  const gridDimension = puzzle.size;
  const cages = puzzle.cages;
  const operations = puzzle.operations !== false;
  const title = puzzle.title ?? '';
  const meta = puzzle.meta ?? '';

  const profile = LAYOUT_PROFILES[gridDimension];
  if (!profile) {
    throw new Error(`Unsupported size: ${String(gridDimension)}`);
  }

  // Store puzzle state
  const state: PuzzleState = { cages, operations, size: gridDimension };
  PropertiesService.getDocumentProperties().setProperty('mathdokuState', JSON.stringify(state));

  const pres = presId === undefined
    ? SlidesApp.getActivePresentation()
    : SlidesApp.openById(presId);

  // Remove default slides
  for (const existingSlide of pres.getSlides()) {
    existingSlide.remove();
  }

  const slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);

  // Remove any default placeholder elements from the blank layout
  for (const element of slide.getPageElements()) {
    element.remove();
  }

  // Ensure slide background is white (covers full page area)
  slide.getBackground().setSolidFill('#FFFFFF');

  // Log actual page dimensions for diagnostics
  Logger.log(`Page dimensions: ${String(pres.getPageWidth())}x${String(pres.getPageHeight())} pt (expected ${String(REF_W)}x${String(REF_H)})`);

  // All coordinates use reference dimensions (960x540 pt); round to integer pt for pixel-perfect match to PowerPoint
  const gridLeft = pt(in2pt(profile.gridLeftIn));
  const gridTop = pt(in2pt(profile.gridTopIn));
  const gridSize = pt(in2pt(profile.gridSizeIn));
  const cellWidth = gridSize / gridDimension;
  const thinPt = profile.thinPt;
  const thickPt = profile.thickPt;

  // Compute cage boundaries
  const { horizontalBounds, verticalBounds } = computeGridBoundaries(cages, gridDimension);

  // ── Title (full slide width, match Python: 0.2" left, SLIDE_W - 0.4" width) ──
  const titleLeft = pt(in2pt(0.2));
  const titleTop = pt(in2pt(0.05));
  const titleWidth = pt(REF_W - in2pt(0.4));
  const titleHeight = pt(in2pt(profile.titleHIn));
  const titleBox = slide.insertTextBox(`${title}\n${meta}`,
    titleLeft, titleTop, titleWidth, titleHeight);
  const titleRange = titleBox.getText();
  titleRange.getRange(0, title.length).getTextStyle()
    .setFontFamily('Segoe UI').setFontSize(profile.titleSz).setBold(true).setForegroundColor(BLACK);
  if (meta.length > 0) {
    titleRange.getRange(title.length + 1, title.length + 1 + meta.length).getTextStyle()
      .setFontFamily('Segoe UI').setFontSize(profile.metaSz).setBold(false).setForegroundColor(VALUE_GRAY);
  }
  titleRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // ── Value + Candidates boxes ──
  const valueProfile = profile.value;
  const candidatesProfile = profile.candidates;

  for (let r = 0; r < gridDimension; r++) {
    for (let c = 0; c < gridDimension; c++) {
      const cellLeft = gridLeft + c * cellWidth;
      const cellTop = gridTop + r * cellWidth;
      const ref = cellRefA1(r, c);

      // VALUE box
      const valueBox = slide.insertTextBox(' ',
        pt(cellLeft), pt(cellTop + valueProfile.yFrac * cellWidth),
        pt(cellWidth), pt(valueProfile.hFrac * cellWidth));
      valueBox.setTitle(`VALUE_${ref}`);
      valueBox.getText().getTextStyle()
        .setFontFamily('Segoe UI').setFontSize(valueProfile.font).setBold(true).setForegroundColor(VALUE_GRAY);
      valueBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      valueBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

      // CANDIDATES box (extend height to compensate for default text box padding;
      // BOTTOM-anchored text is pushed up by bottom padding, extra height prevents upward overflow)
      const candidatesBox = slide.insertTextBox(' ',
        pt(cellLeft + candidatesProfile.xFrac * cellWidth), pt(cellTop + candidatesProfile.yFrac * cellWidth),
        pt(candidatesProfile.wFrac * cellWidth), pt(candidatesProfile.hFrac * cellWidth + 2 * TEXT_BOX_TOP_PAD_PT));
      candidatesBox.setTitle(`CANDIDATES_${ref}`);
      candidatesBox.getText().getTextStyle()
        .setFontFamily(CANDIDATES_FONT).setFontSize(candidatesProfile.font).setBold(false).setForegroundColor(CAND_DARK_RED);
      candidatesBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
      candidatesBox.setContentAlignment(SlidesApp.ContentAlignment.BOTTOM);
    }
  }

  // ── Draw order: thin grid, cage boundaries, join squares, outer border, axis labels, cage labels ──
  drawThinGrid(slide, gridLeft, gridTop, gridSize, gridDimension, thinPt, verticalBounds, horizontalBounds);
  drawCageBoundaries(slide, gridLeft, gridTop, gridSize, gridDimension, thickPt, verticalBounds, horizontalBounds);
  drawJoinSquares(slide, gridLeft, gridTop, gridSize, gridDimension, thickPt, verticalBounds, horizontalBounds);
  drawOuterBorder(slide, gridLeft, gridTop, gridSize, thickPt);
  drawAxisLabels(slide, gridLeft, gridTop, cellWidth, gridDimension, profile);

  // ── Cage labels ──
  drawCageLabels(slide, gridLeft, gridTop, cellWidth, cages, operations, profile);

  // ── Footer ──
  const footerBox = slide.insertTextBox(FOOTER_TEXT,
    pt(in2pt(0.4)), pt(REF_H - in2pt(0.45)), pt(REF_W - in2pt(0.8)), pt(in2pt(0.3)));
  footerBox.getText().getTextStyle()
    .setFontFamily('Segoe UI').setFontSize(14).setBold(false).setForegroundColor(FOOTER_COLOR);
  footerBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END);

  // ── Solve notes columns ──
  const solveProfile = profile.solve;
  const notesLeft = in2pt(solveProfile.leftIn);
  const columnWidth = in2pt(solveProfile.colWIn);
  const columnGap = in2pt(solveProfile.colGapIn);
  for (let i = 0; i < solveProfile.cols; i++) {
    const noteBox = slide.insertTextBox(' ',
      pt(notesLeft + i * (columnWidth + columnGap)), pt(gridTop), pt(columnWidth), pt(gridSize));
    noteBox.setTitle(`SOLVE_NOTES_COL${String(i + 1)}`);
    noteBox.getText().getTextStyle()
      .setFontFamily('Segoe UI').setFontSize(solveProfile.font).setBold(false).setForegroundColor(VALUE_GRAY);
    noteBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
    noteBox.setContentAlignment(SlidesApp.ContentAlignment.TOP);
    noteBox.getBorder().getLineFill().setSolidFill(LIGHT_GRAY_BORDER);
    noteBox.getBorder().setWeight(1);
  }

  // ── Post-scale: fit all content on slide (horizontal and vertical) ──
  const pageWidth = pres.getPageWidth();
  const pageHeight = pres.getPageHeight();
  const contentRight = notesLeft + solveProfile.cols * columnWidth + (solveProfile.cols - 1) * columnGap;
  const contentBottom = Math.max(gridTop + gridSize, REF_H - in2pt(0.45));
  const margin = 20;
  const needScale = pageWidth < REF_W - 0.5 || pageHeight < REF_H - 0.5
    || contentRight > pageWidth - margin || contentBottom > pageHeight - margin;
  const horizontalScale = (pageWidth - margin) / contentRight;
  const verticalScale = (pageHeight - margin) / contentBottom;
  const finalScale = needScale ? Math.min(1, pageWidth / REF_W, pageHeight / REF_H, horizontalScale, verticalScale) : 1;
  if (needScale && finalScale < 1) {
    scaleSlideElements(slide, finalScale);
  }

  Logger.log(`Import complete: ${String(gridDimension)}x${String(gridDimension)} grid, pageW=${String(pageWidth)} pageH=${String(pageHeight)} scale=${finalScale.toFixed(3)}`);
}

function in2pt(inches: number): number {
  return inches * 72;
}

/**
 * Check whether a Slides Color object matches a given hex string constant.
 */
function isColorEqual(color: GoogleAppsScript.Slides.Color, hex: string): boolean {
  return colorToHex(color) === hex.toUpperCase();
}

/** Round to integer pt for pixel-perfect layout (match Python EMU rounding). */
function pt(x: number): number {
  return Math.round(x);
}

const PUZZLE_INIT_OBJECT_ID = 'PuzzleInitData';

function addMathdokuMenu(): void {
  const menu = SlidesApp.getUi().createMenu('Mathdoku');
  const initialized = PropertiesService.getDocumentProperties().getProperty('mathdokuInitialized') === 'true';
  if (!initialized) {
    menu.addItem('Init', 'runInit');
  }
  menu.addItem('Enter', 'showEnterDialog');
  menu.addItem('MakeNextSlide', 'makeNextSlide');
  menu.addToUi();
}

function onOpen(): void {
  addMathdokuMenu();
}

function opSymbol(op: string): string {
  const map: Record<string, string> = {
    '-': '\u2212', '*': 'x', '/': '/',
    '+': '+', '\u00f7': '/', '\u2212': '\u2212',
    'x': 'x', 'X': 'x'
  };
  return map[op.trim()] ?? op.trim();
}

function parseCellRef(token: string): CellRef {
  const m = /^(?<col>[A-Z])(?<row>[1-9]\d*)$/.exec(token.trim().toUpperCase());
  if (!m) {
    throw new Error(`Bad cell ref: ${token}`);
  }
  const groups = ensureNonNullable(m.groups);
  return {
    c: ensureNonNullable(groups['col']).charCodeAt(0) - 65,
    r: parseInt(ensureNonNullable(groups['row']), 10) - 1
  };
}

function runInit(): void {
  const props = PropertiesService.getDocumentProperties();
  if (props.getProperty('mathdokuInitialized') === 'true') {
    SlidesApp.getUi().alert('Already initialized.');
    return;
  }
  const pres = SlidesApp.getActivePresentation();
  const slides = pres.getSlides();
  if (slides.length === 0) {
    SlidesApp.getUi().alert('No slides found.');
    return;
  }
  const slide = ensureNonNullable(slides[0]);
  const el = slide.getPageElementById(PUZZLE_INIT_OBJECT_ID);
  if (el.getPageElementType() !== SlidesApp.PageElementType.SHAPE) {
    SlidesApp.getUi().alert('No puzzle data found. Run the generator again.');
    return;
  }
  const initShape = el.asShape();
  if (initShape.getShapeType() !== SlidesApp.ShapeType.TEXT_BOX) {
    SlidesApp.getUi().alert('No puzzle data found. Run the generator again.');
    return;
  }
  const text = initShape.getText().asString();
  let puzzleJson: PuzzleJson;
  try {
    puzzleJson = JSON.parse(text) as PuzzleJson;
  } catch (_e) {
    SlidesApp.getUi().alert('Invalid puzzle data.');
    return;
  }
  initShape.remove();
  importPuzzle(puzzleJson, pres.getId());
  props.setProperty('mathdokuInitialized', 'true');
  addMathdokuMenu();
}

function scaleSlideElements(slide: GoogleAppsScript.Slides.Slide, scale: number): void {
  for (const element of slide.getPageElements()) {
    element.setLeft(element.getLeft() * scale);
    element.setTop(element.getTop() * scale);
    element.setWidth(element.getWidth() * scale);
    element.setHeight(element.getHeight() * scale);

    const type = element.getPageElementType();
    if (type === SlidesApp.PageElementType.SHAPE) {
      const shape = element.asShape();
      for (const run of shape.getText().getRuns()) {
        const fontSize = run.getTextStyle().getFontSize();
        if (fontSize) {
          run.getTextStyle().setFontSize(Math.max(7, Math.round(fontSize * scale)));
        }
      }
      try {
        const borderWeight = shape.getBorder().getWeight();
        if (borderWeight > 0) {
          shape.getBorder().setWeight(borderWeight * scale);
        }
      } catch (_e) { /* no border */ }
    } else if (type === SlidesApp.PageElementType.LINE) {
      const line = element.asLine();
      const lineWeight = line.getWeight();
      line.setWeight(Math.max(1, lineWeight * scale));
    }
  }
}
