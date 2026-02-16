/**
 * Mathdoku Solver -- Google Slides bound script.
 *
 * Main entry: Code.ts (menu, helpers, importPuzzle)
 */

// ── Assert helpers ──────────────────────────────────────────────────────────

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

// ── Interfaces ─────────────────────────────────────────────────────────────

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

// ── Helpers ─────────────────────────────────────────────────────────────────

function drawAxisLabels(
  slide: GoogleAppsScript.Slides.Slide,
  gridLeft: number,
  gridTop: number,
  cellW: number,
  n: number,
  prof: LayoutProfile
): void {
  const axisFont = prof.axisFont;
  const labelH = pt(in2pt(prof.axisLabelH));
  const labelW = pt(in2pt(prof.axisLabelW));
  const topOffset = in2pt(prof.axisTopOffset);
  const sideOffset = in2pt(prof.axisSideOffset);
  const topY = pt(gridTop - topOffset - TEXT_BOX_TOP_PAD_PT);
  const sideX = pt(gridLeft - sideOffset);

  // Column labels (A, B, C, ...): use running x so alignment doesn't drift from rounding
  let colLeft = gridLeft;
  for (let c = 0; c < n; c++) {
    const boxW = pt(cellW);
    const box = slide.insertTextBox(String.fromCharCode(65 + c), pt(colLeft), topY, boxW, labelH);
    box.getText().getTextStyle()
      .setFontFamily('Segoe UI').setFontSize(axisFont).setBold(true).setForegroundColor(AXIS_LABEL_MAGENTA);
    box.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    box.getBorder().setTransparent();
    colLeft += cellW;
  }

  // Row labels (1, 2, 3, ...)
  for (let r = 0; r < n; r++) {
    const y = pt(gridTop + r * cellW + (cellW - labelH) / 2);
    const box = slide.insertTextBox(String(r + 1), sideX, y, labelW, labelH);
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
  n: number,
  thickPt: number,
  vBound: boolean[][],
  hBound: boolean[][]
): void {
  const cellW = gridSize / n;
  const inset = thickPt / 2;

  // Vertical runs
  for (let c = 1; c < n; c++) {
    let r0 = 0;
    while (r0 < n) {
      if (!ensureNonNullable(vBound[r0])[c - 1]) {
        r0++;
        continue;
      }
      let r1 = r0;
      while (r1 + 1 < n && ensureNonNullable(vBound[r1 + 1])[c - 1]) {
        r1++;
      }
      const x = gridLeft + c * cellW;
      let y1 = gridTop + r0 * cellW;
      let y2 = gridTop + (r1 + 1) * cellW;
      if (r0 === 0) {
        y1 += inset;
      }
      if (r1 === n - 1) {
        y2 -= inset;
      }
      drawThickRect(slide, x - thickPt / 2, y1, thickPt, y2 - y1);
      r0 = r1 + 1;
    }
  }

  // Horizontal runs
  for (let r = 1; r < n; r++) {
    let c0 = 0;
    while (c0 < n) {
      if (!ensureNonNullable(hBound[r - 1])[c0]) {
        c0++;
        continue;
      }
      let c1 = c0;
      while (c1 + 1 < n && ensureNonNullable(hBound[r - 1])[c1 + 1]) {
        c1++;
      }
      const y = gridTop + r * cellW;
      let x1 = gridLeft + c0 * cellW;
      let x2 = gridLeft + (c1 + 1) * cellW;
      if (c0 === 0) {
        x1 += inset;
      }
      if (c1 === n - 1) {
        x2 -= inset;
      }
      drawThickRect(slide, x1, y - thickPt / 2, x2 - x1, thickPt);
      c0 = c1 + 1;
    }
  }
}

function drawCageLabels(
  slide: GoogleAppsScript.Slides.Slide,
  gridLeft: number,
  gridTop: number,
  cellW: number,
  cages: Cage[],
  operations: boolean,
  prof: LayoutProfile
): void {
  const cageProf = prof.cage;
  const insetX = cageProf.insetXFrac * cellW;
  const insetY = cageProf.insetYFrac * cellW;
  const boxW = cageProf.boxWFrac * cellW;
  const boxH = cageProf.boxHFrac * cellW;

  for (let i = 0; i < cages.length; i++) {
    const cage = ensureNonNullable(cages[i]);

    // Find geometric top-left cell (smallest row, then smallest column)
    const parsed = cage.cells.map((ref) => ({ ref, ...parseCellRef(ref) }));
    parsed.sort((a, b) => a.r !== b.r ? a.r - b.r : a.c - b.c);
    const tl = ensureNonNullable(parsed[0]);
    const tlRef = tl.ref;

    // Build label
    let label = cage.label ?? '';
    if (!label && cage.value !== undefined) {
      label = operations && cage.cells.length > 1 && cage.op
        ? String(cage.value) + opSymbol(cage.op)
        : String(cage.value);
    }

    const x = pt(gridLeft + tl.c * cellW + insetX);
    const y = pt(gridTop + tl.r * cellW + insetY - TEXT_BOX_TOP_PAD_PT);
    // Size font for usable area (box minus default top+bottom padding)
    const effectiveBoxH = Math.max(7, boxH - 2 * TEXT_BOX_TOP_PAD_PT);
    const actualFont = fitFontSize(label, cageProf.font, boxW / 72, effectiveBoxH / 72);
    // Keep original box height (text is smaller so it fits within padded area)
    const box = slide.insertTextBox(label, x, y, pt(boxW), pt(boxH));
    box.setTitle(`CAGE_${String(i)}_${tlRef}`);
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
  n: number,
  thickPt: number,
  vBound: boolean[][],
  hBound: boolean[][]
): void {
  const cellW = gridSize / n;
  const clamp = (v: number, lo: number, hi: number): number =>
    Math.max(lo, Math.min(hi, v));

  // Draw join squares at every interior vertex where any cage boundary touches (spec: "if any of ... is true")
  for (let vr = 1; vr < n; vr++) {
    for (let vc = 1; vc < n; vc++) {
      const vAbove = ensureNonNullable(vBound[vr - 1])[vc - 1];
      const vBelow = ensureNonNullable(vBound[vr])[vc - 1];
      const hLeft = ensureNonNullable(hBound[vr - 1])[vc - 1];
      const hRight = ensureNonNullable(hBound[vr - 1])[vc];
      if (!vAbove && !vBelow && !hLeft && !hRight) {
        continue;
      }
      const x = gridLeft + vc * cellW;
      const y = gridTop + vr * cellW;
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
  const half = thickPt / 2;
  // Top
  drawThickRect(slide, gridLeft - half, gridTop - half, gridSize + thickPt, thickPt);
  // Bottom
  drawThickRect(slide, gridLeft - half, gridTop + gridSize - half, gridSize + thickPt, thickPt);
  // Left
  drawThickRect(slide, gridLeft - half, gridTop + half, thickPt, Math.max(0, gridSize - thickPt));
  // Right
  drawThickRect(slide, gridLeft + gridSize - half, gridTop + half, thickPt, Math.max(0, gridSize - thickPt));
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
  n: number,
  thinW: number,
  vBound: boolean[][],
  hBound: boolean[][]
): void {
  const cellW = gridSize / n;
  const halfThin = thinW / 2;

  // Vertical thin lines: between columns c-1 and c, for each row
  for (let c = 1; c < n; c++) {
    for (let r = 0; r < n; r++) {
      if (ensureNonNullable(vBound[r])[c - 1]) {
        continue;
      }
      const x = gridLeft + c * cellW;
      const y = gridTop + r * cellW;
      const rect = slide.insertShape(SlidesApp.ShapeType.RECTANGLE,
        pt(x - halfThin), pt(y), pt(thinW), pt(cellW));
      rect.getFill().setSolidFill(THIN_GRAY);
      rect.getBorder().setTransparent();
    }
  }

  // Horizontal thin lines: between rows r-1 and r, for each column
  for (let r = 1; r < n; r++) {
    for (let c = 0; c < n; c++) {
      if (ensureNonNullable(hBound[r - 1])[c]) {
        continue;
      }
      const x = gridLeft + c * cellW;
      const y = gridTop + r * cellW;
      const rect = slide.insertShape(SlidesApp.ShapeType.RECTANGLE,
        pt(x), pt(y - halfThin), pt(cellW), pt(thinW));
      rect.getFill().setSolidFill(THIN_GRAY);
      rect.getBorder().setTransparent();
    }
  }
}

function fitFontSize(text: string, basePt: number, boxWIn: number, boxHIn: number): number {
  const t = text.trim();
  if (!t) {
    return basePt;
  }
  const maxWPt = Math.max(1, boxWIn * 72 - 2);
  const maxHPt = Math.max(1, boxHIn * 72 - 1);
  const charCount = Math.max(1, t.length);
  const widthBased = Math.floor(maxWPt / (0.60 * charCount));
  const heightBased = Math.floor(maxHPt / 1.15);
  return Math.max(7, Math.min(basePt, widthBased, heightBased));
}

function formatCandidates(digits: string, gridSize: number): string {
  // Fixed-position layout: each digit at its slot, missing digits replaced with space.
  // No inter-digit spaces — original used Font.Spacing for visual separation,
  // which Google Slides API doesn't support (see CLAUDE.md deviations).
  const half = Math.ceil(gridSize / 2);
  let line1 = '';
  let line2 = '';
  for (let d = 1; d <= gridSize; d++) {
    const ch = digits.includes(String(d)) ? String(d) : ' ';
    if (d <= half) {
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
  const n = puzzle.size;
  const cages = puzzle.cages;
  const operations = puzzle.operations !== false;
  const title = puzzle.title ?? '';
  const meta = puzzle.meta ?? '';

  const prof = LAYOUT_PROFILES[n];
  if (!prof) {
    throw new Error(`Unsupported size: ${String(n)}`);
  }

  // Store puzzle state
  const state: PuzzleState = { cages, operations, size: n };
  PropertiesService.getDocumentProperties().setProperty('mathdokuState', JSON.stringify(state));

  const pres = presId === undefined
    ? SlidesApp.getActivePresentation()
    : SlidesApp.openById(presId);

  // Remove default slides
  for (const s of pres.getSlides()) {
    s.remove();
  }

  const slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);

  // Remove any default placeholder elements from the blank layout
  for (const el of slide.getPageElements()) {
    el.remove();
  }

  // Ensure slide background is white (covers full page area)
  slide.getBackground().setSolidFill('#FFFFFF');

  // Log actual page dimensions for diagnostics
  Logger.log(`Page dimensions: ${String(pres.getPageWidth())}x${String(pres.getPageHeight())} pt (expected ${String(REF_W)}x${String(REF_H)})`);

  // All coordinates use reference dimensions (960x540 pt); round to integer pt for pixel-perfect match to PowerPoint
  const gridLeft = pt(in2pt(prof.gridLeftIn));
  const gridTop = pt(in2pt(prof.gridTopIn));
  const gridSize = pt(in2pt(prof.gridSizeIn));
  const cellW = gridSize / n;
  const thinPt = prof.thinPt;
  const thickPt = prof.thickPt;

  // Build cell-to-cage mapping
  const cellToCage: Record<string, number> = {};
  for (let ci = 0; ci < cages.length; ci++) {
    const cage = ensureNonNullable(cages[ci]);
    for (const cell of cage.cells) {
      cellToCage[cell] = ci;
    }
  }

  // Compute cage boundaries
  const vBound: boolean[][] = [];
  const hBound: boolean[][] = [];
  for (let r = 0; r < n; r++) {
    const row: boolean[] = [];
    vBound[r] = row;
    for (let c = 1; c < n; c++) {
      const leftRef = cellRefA1(r, c - 1);
      const rightRef = cellRefA1(r, c);
      row[c - 1] = cellToCage[leftRef] !== cellToCage[rightRef];
    }
  }
  for (let r = 1; r < n; r++) {
    const row: boolean[] = [];
    hBound[r - 1] = row;
    for (let c = 0; c < n; c++) {
      const topRef = cellRefA1(r - 1, c);
      const botRef = cellRefA1(r, c);
      row[c] = cellToCage[topRef] !== cellToCage[botRef];
    }
  }

  // ── Title (full slide width, match Python: 0.2" left, SLIDE_W - 0.4" width) ──
  const titleLeft = pt(in2pt(0.2));
  const titleTop = pt(in2pt(0.05));
  const titleWidth = pt(REF_W - in2pt(0.4));
  const titleHeight = pt(in2pt(prof.titleHIn));
  const titleBox = slide.insertTextBox(`${title}\n${meta}`,
    titleLeft, titleTop, titleWidth, titleHeight);
  const titleRange = titleBox.getText();
  titleRange.getRange(0, title.length).getTextStyle()
    .setFontFamily('Segoe UI').setFontSize(prof.titleSz).setBold(true).setForegroundColor(BLACK);
  if (meta.length > 0) {
    titleRange.getRange(title.length + 1, title.length + 1 + meta.length).getTextStyle()
      .setFontFamily('Segoe UI').setFontSize(prof.metaSz).setBold(false).setForegroundColor(VALUE_GRAY);
  }
  titleRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // ── Value + Candidates boxes ──
  const valProf = prof.value;
  const candProf = prof.candidates;

  for (let r = 0; r < n; r++) {
    for (let c = 0; c < n; c++) {
      const cellLeft = gridLeft + c * cellW;
      const cellTop = gridTop + r * cellW;
      const ref = cellRefA1(r, c);

      // VALUE box
      const valBox = slide.insertTextBox(' ',
        pt(cellLeft), pt(cellTop + valProf.yFrac * cellW),
        pt(cellW), pt(valProf.hFrac * cellW));
      valBox.setTitle(`VALUE_${ref}`);
      valBox.getText().getTextStyle()
        .setFontFamily('Segoe UI').setFontSize(valProf.font).setBold(true).setForegroundColor(VALUE_GRAY);
      valBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      valBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

      // CANDIDATES box
      const candBox = slide.insertTextBox(' ',
        pt(cellLeft + candProf.xFrac * cellW), pt(cellTop + candProf.yFrac * cellW),
        pt(candProf.wFrac * cellW), pt(candProf.hFrac * cellW));
      candBox.setTitle(`CANDIDATES_${ref}`);
      candBox.getText().getTextStyle()
        .setFontFamily(CANDIDATES_FONT).setFontSize(candProf.font).setBold(false).setForegroundColor(CAND_DARK_RED);
      candBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
      candBox.setContentAlignment(SlidesApp.ContentAlignment.BOTTOM);
    }
  }

  // ── Draw order: thin grid, cage boundaries, join squares, outer border, axis labels, cage labels ──
  drawThinGrid(slide, gridLeft, gridTop, gridSize, n, thinPt, vBound, hBound);
  drawCageBoundaries(slide, gridLeft, gridTop, gridSize, n, thickPt, vBound, hBound);
  drawJoinSquares(slide, gridLeft, gridTop, gridSize, n, thickPt, vBound, hBound);
  drawOuterBorder(slide, gridLeft, gridTop, gridSize, thickPt);
  drawAxisLabels(slide, gridLeft, gridTop, cellW, n, prof);

  // ── Cage labels ──
  drawCageLabels(slide, gridLeft, gridTop, cellW, cages, operations, prof);

  // ── Footer ──
  const footerBox = slide.insertTextBox(FOOTER_TEXT,
    pt(in2pt(0.4)), pt(REF_H - in2pt(0.45)), pt(REF_W - in2pt(0.8)), pt(in2pt(0.3)));
  footerBox.getText().getTextStyle()
    .setFontFamily('Segoe UI').setFontSize(14).setBold(false).setForegroundColor(FOOTER_COLOR);
  footerBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END);

  // ── Solve notes columns ──
  const solveProf = prof.solve;
  const notesLeft = in2pt(solveProf.leftIn);
  const colW = in2pt(solveProf.colWIn);
  const colGap = in2pt(solveProf.colGapIn);
  for (let i = 0; i < solveProf.cols; i++) {
    const noteBox = slide.insertTextBox(' ',
      pt(notesLeft + i * (colW + colGap)), pt(gridTop), pt(colW), pt(gridSize));
    noteBox.setTitle(`SOLVE_NOTES_COL${String(i + 1)}`);
    noteBox.getText().getTextStyle()
      .setFontFamily('Segoe UI').setFontSize(solveProf.font).setBold(false).setForegroundColor(VALUE_GRAY);
    noteBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
    noteBox.setContentAlignment(SlidesApp.ContentAlignment.TOP);
    noteBox.getBorder().getLineFill().setSolidFill(LIGHT_GRAY_BORDER);
    noteBox.getBorder().setWeight(1);
  }

  // ── Post-scale: fit all content on slide (horizontal and vertical) ──
  const pageW = pres.getPageWidth();
  const pageH = pres.getPageHeight();
  const contentRight = notesLeft + solveProf.cols * colW + (solveProf.cols - 1) * colGap;
  const contentBottom = Math.max(gridTop + gridSize, REF_H - in2pt(0.45));
  const margin = 20;
  const needScale = pageW < REF_W - 0.5 || pageH < REF_H - 0.5
    || contentRight > pageW - margin || contentBottom > pageH - margin;
  const scaleW = (pageW - margin) / contentRight;
  const scaleH = (pageH - margin) / contentBottom;
  const finalScale = needScale ? Math.min(1, pageW / REF_W, pageH / REF_H, scaleW, scaleH) : 1;
  if (needScale && finalScale < 1) {
    scaleSlideElements(slide, finalScale);
  }

  Logger.log(`Import complete: ${String(n)}x${String(n)} grid, pageW=${String(pageW)} pageH=${String(pageH)} scale=${finalScale.toFixed(3)}`);
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
  if (el?.getPageElementType() !== SlidesApp.PageElementType.SHAPE) {
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
  for (const el of slide.getPageElements()) {
    el.setLeft(el.getLeft() * scale);
    el.setTop(el.getTop() * scale);
    el.setWidth(el.getWidth() * scale);
    el.setHeight(el.getHeight() * scale);

    const type = el.getPageElementType();
    if (type === SlidesApp.PageElementType.SHAPE) {
      const shape = el.asShape();
      for (const run of shape.getText().getRuns()) {
        const fs = run.getTextStyle().getFontSize();
        if (fs) {
          run.getTextStyle().setFontSize(Math.max(7, Math.round(fs * scale)));
        }
      }
      try {
        const bw = shape.getBorder().getWeight();
        if (bw > 0) {
          shape.getBorder().setWeight(bw * scale);
        }
      } catch (_e) { /* no border */ }
    } else if (type === SlidesApp.PageElementType.LINE) {
      const line = el.asLine();
      const w = line.getWeight();
      line.setWeight(Math.max(1, w * scale));
    }
  }
}
