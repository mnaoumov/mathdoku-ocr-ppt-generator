/**
 * View.ts -- Google Slides rendering layer.
 *
 * Contains: SlidesRenderer class (implements PuzzleRenderer), all Google
 * Slides code, color/layout constants, global entry points.
 */

interface CageProfile {
  readonly boxHeightFraction: number;
  readonly boxWidthFraction: number;
  readonly font: number;
  readonly insetLeftFraction: number;
  readonly insetTopFraction: number;
}

interface CandidatesProfile {
  readonly digitMargin: number;
  readonly font: number;
  readonly heightFraction: number;
  readonly leftFraction: number;
  readonly topFraction: number;
  readonly widthFraction: number;
}

interface LayoutProfile {
  readonly axisFont: number;
  readonly axisLabelHeight: number;
  readonly axisLabelWidth: number;
  readonly axisSideOffset: number;
  readonly axisTopOffset: number;
  readonly cage: CageProfile;
  readonly candidates: CandidatesProfile;
  readonly gridLeftInches: number;
  readonly gridSizeInches: number;
  readonly gridTopInches: number;
  readonly metaFontSize: number;
  readonly solve: SolveProfile;
  readonly thickPt: number;
  readonly thinPt: number;
  readonly titleFontSize: number;
  readonly titleHeightInches: number;
  readonly value: ValueProfile;
}

interface SolveProfile {
  readonly columnCount: number;
  readonly columnGapInches: number;
  readonly columnWidthInches: number;
  readonly font: number;
  readonly leftInches: number;
}

interface ValueProfile {
  readonly font: number;
  readonly heightFraction: number;
  readonly topFraction: number;
}

class SlidesRenderer implements PuzzleRenderer {
  private currentGridSize = 0;
  private currentSlide: GoogleAppsScript.Slides.Slide | null = null;

  public beginPendingRender(gridSize: number): void {
    makeNextSlide(gridSize);
    this.currentSlide = getCurrentSlide();
    this.currentGridSize = gridSize;
  }

  public renderCommittedChanges(gridSize: number): void {
    makeNextSlide(gridSize);
    this.currentSlide = null;
  }

  public renderPendingCandidates(change: CandidatesChange): void {
    const slide = ensureNonNullable(this.currentSlide);
    applyPendingCandidates(slide, change.cell.ref, change.values, this.currentGridSize);
  }

  public renderPendingClearance(change: CellClearance): void {
    const slide = ensureNonNullable(this.currentSlide);
    clearShapeText(slide, `VALUE_${change.cell.ref}`);
    clearShapeText(slide, `CANDIDATES_${change.cell.ref}`);
  }

  public renderPendingStrikethrough(change: CandidatesStrikethrough): void {
    const slide = ensureNonNullable(this.currentSlide);
    applyPendingStrikethrough(slide, change.cell.ref, change.values);
  }

  public renderPendingValue(change: ValueChange): void {
    const slide = ensureNonNullable(this.currentSlide);
    applyPendingValue(slide, change.cell.ref, String(change.value));
  }
}

function addChanges(): void {
  try {
    if (PropertiesService.getDocumentProperties().getProperty('mathdokuInitialized') !== 'true') {
      SlidesApp.getUi().alert('Please run Mathdoku > Init first.');
      return;
    }
    const html = HtmlService.createHtmlOutputFromFile('EnterDialog')
      .setWidth(ENTER_DIALOG_WIDTH_PX)
      .setHeight(ENTER_DIALOG_HEIGHT_PX);
    SlidesApp.getUi().showModelessDialog(html, 'Edit Cell');
  } catch (e: unknown) {
    showError('Add changes', e);
  }
}

function addChangesFromInput(input: string): void {
  const puzzle = buildPuzzleFromSlide();
  puzzle.enter(input);
  puzzle.commit();
}

function addChangesFromStrategies(): number {
  const puzzle = buildPuzzleFromSlide();
  return puzzle.applyEasyStrategies();
}

function addMathdokuMenu(): void {
  const menu = SlidesApp.getUi().createMenu('Mathdoku');
  const initialized = PropertiesService.getDocumentProperties().getProperty('mathdokuInitialized') === 'true';
  if (initialized) {
    menu.addItem('Add changes', 'addChanges');
    menu.addItem('Apply easy strategies', 'applyEasyStrategies');
  } else {
    menu.addItem('Init', 'init');
  }
  menu.addToUi();
}

function applyEasyStrategies(): void {
  try {
    const steps = addChangesFromStrategies();
    if (steps === 0) {
      SlidesApp.getUi().alert('No further solving steps found.');
    }
  } catch (e: unknown) {
    showError('Apply easy strategies', e);
  }
}

function applyPendingCandidates(
  slide: GoogleAppsScript.Slides.Slide,
  cellRef: string,
  values: readonly number[],
  gridSize: number
): void {
  const shape = getShapeByTitle(slide, `CANDIDATES_${cellRef}`);
  if (!shape) {
    throw new Error(`Candidates shape not found: CANDIDATES_${cellRef}`);
  }

  const formatted = formatCandidates(values, gridSize);
  const textRange = shape.getText();
  textRange.setText(formatted);
  textRange.getTextStyle()
    .setFontFamily(CANDIDATES_FONT)
    .setForegroundColor(GREEN);

  clearShapeText(slide, `VALUE_${cellRef}`);
}

function applyPendingStrikethrough(
  slide: GoogleAppsScript.Slides.Slide,
  cellRef: string,
  values: readonly number[]
): void {
  const shape = getShapeByTitle(slide, `CANDIDATES_${cellRef}`);
  if (!shape) {
    throw new Error(`Candidates shape not found: CANDIDATES_${cellRef}`);
  }

  const valueChars = new Set(values.map(String));
  const textRange = shape.getText();
  const fullText = textRange.asString();

  for (let i = 0; i < fullText.length; i++) {
    const ch = fullText.charAt(i);
    if (valueChars.has(ch)) {
      textRange.getRange(i, i + 1).getTextStyle()
        .setForegroundColor(GREEN)
        .setStrikethrough(true);
    }
  }
}

function applyPendingValue(
  slide: GoogleAppsScript.Slides.Slide,
  cellRef: string,
  valueText: string
): void {
  const shape = getShapeByTitle(slide, `VALUE_${cellRef}`);
  if (!shape) {
    throw new Error(`Value shape not found: VALUE_${cellRef}`);
  }

  const textRange = shape.getText();
  textRange.setText(valueText);
  textRange.getTextStyle()
    .setFontFamily('Segoe UI')
    .setBold(true)
    .setForegroundColor(GREEN);

  clearShapeText(slide, `CANDIDATES_${cellRef}`);
}

function buildPuzzleFromSlide(): Puzzle {
  const state = getPuzzleState();
  const slide = getCurrentSlide();
  const values = new Map<string, number>();
  const candidates = new Map<string, Set<number>>();

  for (let r = 0; r < state.size; r++) {
    for (let c = 0; c < state.size; c++) {
      const ref = cellRefA1(r, c);
      readCellFromSlide(slide, ref, values, candidates);
    }
  }

  return new Puzzle(
    new SlidesRenderer(),
    state.size,
    state.cages,
    state.hasOperators,
    '',
    '',
    values,
    candidates
  );
}

function clamp(value: number, lo: number, hi: number): number {
  return Math.max(lo, Math.min(hi, value));
}

function clearShapeText(slide: GoogleAppsScript.Slides.Slide, title: string): void {
  const shape = getShapeByTitle(slide, title);
  if (shape) {
    try {
      shape.getText().setText(' ');
    } catch (_e) { /* Shape may not support text */ }
  }
}

function colorToHex(color: GoogleAppsScript.Slides.Color): string {
  return color.asRgbColor().asHexString().toUpperCase();
}

function drawAxisLabels(
  slide: GoogleAppsScript.Slides.Slide,
  gridLeft: number,
  gridTop: number,
  cellWidth: number,
  gridDimension: number,
  profile: LayoutProfile
): void {
  const axisFont = profile.axisFont;
  const labelHeight = pt(in2pt(profile.axisLabelHeight));
  const labelWidth = pt(in2pt(profile.axisLabelWidth));
  const topOffset = in2pt(profile.axisTopOffset);
  const sideOffset = in2pt(profile.axisSideOffset);
  const topY = pt(gridTop - topOffset - TEXT_BOX_TOP_PADDING_PT);
  const sideX = pt(gridLeft - sideOffset);

  let columnLeft = gridLeft;
  for (let c = 0; c < gridDimension; c++) {
    const boxWidth = pt(cellWidth);
    const box = slide.insertTextBox(String.fromCharCode(CHAR_CODE_A + c), pt(columnLeft), topY, boxWidth, labelHeight);
    box.getText().getTextStyle()
      .setFontFamily('Segoe UI').setFontSize(axisFont).setBold(true).setForegroundColor(AXIS_LABEL_MAGENTA);
    box.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    box.getBorder().setTransparent();
    columnLeft += cellWidth;
  }

  for (let r = 0; r < gridDimension; r++) {
    const y = pt(gridTop + r * cellWidth + (cellWidth - labelHeight) / SIDE_COUNT);
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
  verticalBounds: readonly (readonly boolean[])[],
  horizontalBounds: readonly (readonly boolean[])[]
): void {
  const cellWidth = gridSize / gridDimension;
  const inset = thickPt / SIDE_COUNT;

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
      drawThickRect(slide, x - thickPt / SIDE_COUNT, y1, thickPt, y2 - y1);
      startRow = endRow + 1;
    }
  }

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
      drawThickRect(slide, x1, y - thickPt / SIDE_COUNT, x2 - x1, thickPt);
      startCol = endCol + 1;
    }
  }
}

function drawCageLabels(
  slide: GoogleAppsScript.Slides.Slide,
  gridLeft: number,
  gridTop: number,
  cellWidth: number,
  cages: readonly CageRaw[],
  hasOperators: boolean,
  profile: LayoutProfile
): void {
  const cageProfile = profile.cage;
  const insetX = cageProfile.insetLeftFraction * cellWidth;
  const insetY = cageProfile.insetTopFraction * cellWidth;
  const labelBoxWidth = cageProfile.boxWidthFraction * cellWidth;
  const labelBoxHeight = Math.min(
    cageProfile.boxHeightFraction * cellWidth,
    profile.candidates.topFraction * cellWidth - insetY
  );

  for (let i = 0; i < cages.length; i++) {
    const cage = ensureNonNullable(cages[i]);
    const parsed = cage.cells.map((ref) => ({ ref, ...parseCellRef(ref) }));
    parsed.sort((a, b) => a.rowId === b.rowId ? a.columnId - b.columnId : a.rowId - b.rowId);
    const topLeftCell = ensureNonNullable(parsed[0]);
    const topLeftCellRef = topLeftCell.ref;

    let label = cage.label ?? '';
    if (!label && cage.value !== undefined) {
      label = hasOperators && cage.cells.length > 1 && cage.operator
        ? String(cage.value) + opSymbol(cage.operator)
        : String(cage.value);
    }

    const x = pt(gridLeft + (topLeftCell.columnId - 1) * cellWidth + insetX);
    const y = pt(gridTop + (topLeftCell.rowId - 1) * cellWidth + insetY - TEXT_BOX_TOP_PADDING_PT);
    const usableHeight = Math.max(MIN_FONT_SIZE, labelBoxHeight - TEXT_BOX_TOP_PADDING_PT);
    const actualFont = fitFontSize(label, cageProfile.font, labelBoxWidth / POINTS_PER_INCH, usableHeight / POINTS_PER_INCH);
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
  verticalBounds: readonly (readonly boolean[])[],
  horizontalBounds: readonly (readonly boolean[])[]
): void {
  const cellWidth = gridSize / gridDimension;

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
      const left = clamp(x - thickPt / SIDE_COUNT, gridLeft, gridLeft + gridSize - thickPt);
      const top = clamp(y - thickPt / SIDE_COUNT, gridTop, gridTop + gridSize - thickPt);
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
  const halfThick = thickPt / SIDE_COUNT;
  drawThickRect(slide, gridLeft - halfThick, gridTop - halfThick, gridSize + thickPt, thickPt);
  drawThickRect(slide, gridLeft - halfThick, gridTop + gridSize - halfThick, gridSize + thickPt, thickPt);
  drawThickRect(slide, gridLeft - halfThick, gridTop + halfThick, thickPt, Math.max(0, gridSize - thickPt));
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
  verticalBounds: readonly (readonly boolean[])[],
  horizontalBounds: readonly (readonly boolean[])[]
): void {
  const cellWidth = gridSize / gridDimension;
  const halfThinWidth = thinWidth / SIDE_COUNT;

  for (let c = 1; c < gridDimension; c++) {
    for (let r = 0; r < gridDimension; r++) {
      if (ensureNonNullable(verticalBounds[r])[c - 1]) {
        continue;
      }
      const x = gridLeft + c * cellWidth;
      const y = gridTop + r * cellWidth;
      const rect = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, pt(x - halfThinWidth), pt(y), pt(thinWidth), pt(cellWidth));
      rect.getFill().setSolidFill(THIN_GRAY);
      rect.getBorder().setTransparent();
    }
  }

  for (let r = 1; r < gridDimension; r++) {
    for (let c = 0; c < gridDimension; c++) {
      if (ensureNonNullable(horizontalBounds[r - 1])[c]) {
        continue;
      }
      const x = gridLeft + c * cellWidth;
      const y = gridTop + r * cellWidth;
      const rect = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, pt(x), pt(y - halfThinWidth), pt(cellWidth), pt(thinWidth));
      rect.getFill().setSolidFill(THIN_GRAY);
      rect.getBorder().setTransparent();
    }
  }
}

function finalizeCandidatesShape(shape: GoogleAppsScript.Slides.Shape, gridSize: number): void {
  const textRange = shape.getText();
  const fullText = textRange.asString().replace(/\n$/, '');
  if (fullText.trim().length === 0 || fullText.trim() === ' ') {
    return;
  }

  let hasGreenStrike = false;
  let hasGreen = false;
  const survivingValues: number[] = [];

  for (let i = 0; i < fullText.length; i++) {
    const ch = fullText.charAt(i);
    if (ch < '1' || ch > '9') {
      continue;
    }

    const style = textRange.getRange(i, i + 1).getTextStyle();
    const color = style.getForegroundColor();
    const strike = style.isStrikethrough();
    const isGreen = isColorEqual(color, GREEN);

    if (isGreen && strike) {
      hasGreenStrike = true;
      continue;
    }

    if (isGreen) {
      hasGreen = true;
    }

    survivingValues.push(parseInt(ch, 10));
  }

  if (hasGreenStrike) {
    const newFormatted = formatCandidates(survivingValues, gridSize);
    textRange.setText(newFormatted);
    textRange.getTextStyle()
      .setFontFamily(CANDIDATES_FONT)
      .setForegroundColor(CANDIDATES_DARK_RED)
      .setStrikethrough(false);
    return;
  }

  if (hasGreen) {
    for (let i = 0; i < fullText.length; i++) {
      const ch = fullText.charAt(i);
      if (ch < '1' || ch > '9') {
        continue;
      }
      const style = textRange.getRange(i, i + 1).getTextStyle();
      const color = style.getForegroundColor();
      if (isColorEqual(color, GREEN)) {
        style.setForegroundColor(CANDIDATES_DARK_RED);
      }
    }
  }
}

function finalizeValueShape(shape: GoogleAppsScript.Slides.Shape): void {
  const textRange = shape.getText();
  const fullText = textRange.asString().replace(/\n$/, '');
  if (fullText.trim().length === 0 || fullText.trim() === ' ') {
    return;
  }

  for (let i = 0; i < fullText.length; i++) {
    const ch = fullText.charAt(i);
    if (ch < '1' || ch > '9') {
      continue;
    }
    const style = textRange.getRange(i, i + 1).getTextStyle();
    const color = style.getForegroundColor();
    if (isColorEqual(color, GREEN)) {
      style.setForegroundColor(VALUE_GRAY);
    }
  }
}

function fitFontSize(text: string, basePt: number, boxWidthIn: number, boxHeightIn: number): number {
  const trimmedText = text.trim();
  if (!trimmedText) {
    return basePt;
  }
  const maxWidthPt = Math.max(1, boxWidthIn * POINTS_PER_INCH - FONT_FIT_PADDING_PT);
  const maxHeightPt = Math.max(1, boxHeightIn * POINTS_PER_INCH - 1);
  const charCount = Math.max(1, trimmedText.length);
  const widthBased = Math.floor(maxWidthPt / (FONT_FIT_WIDTH_RATIO * charCount));
  const heightBased = Math.floor(maxHeightPt / FONT_FIT_HEIGHT_RATIO);
  return Math.max(MIN_FONT_SIZE, Math.min(basePt, widthBased, heightBased));
}

function formatCandidates(values: readonly number[], gridSize: number): string {
  const valueSet = new Set(values);
  const firstRowCount = Math.ceil(gridSize / CANDIDATE_ROW_COUNT);
  let line1 = '';
  let line2 = '';
  for (let d = 1; d <= gridSize; d++) {
    const ch = valueSet.has(d) ? String(d) : ' ';
    if (d <= firstRowCount) {
      line1 += ch;
    } else {
      line2 += ch;
    }
  }
  return `${line1}\n${line2}`;
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

function getShapeByTitle(slide: GoogleAppsScript.Slides.Slide, title: string): GoogleAppsScript.Slides.Shape | null {
  for (const el of slide.getPageElements()) {
    if (el.getTitle() === title) {
      return el.asShape();
    }
  }
  return null;
}

function importPuzzle(puzzleJson: PuzzleJson | string, presId?: string): void {
  const parsed: PuzzleJson = (typeof puzzleJson === 'string') ? JSON.parse(puzzleJson) as PuzzleJson : puzzleJson;
  const gridDimension = parsed.size;
  const cages = parsed.cages;
  const hasOperators = parsed.hasOperators !== false;
  const title = parsed.title ?? '';
  const meta = parsed.meta ?? '';

  const profile = LAYOUT_PROFILES[gridDimension];
  if (!profile) {
    throw new Error(`Unsupported size: ${String(gridDimension)}`);
  }

  const state: PuzzleState = { cages, hasOperators, size: gridDimension };
  const docProps = PropertiesService.getDocumentProperties();
  docProps.setProperty('mathdokuState', JSON.stringify(state));
  docProps.setProperty('mathdokuInitialized', 'true');

  const pres = presId === undefined
    ? SlidesApp.getActivePresentation()
    : SlidesApp.openById(presId);

  for (const existingSlide of pres.getSlides()) {
    existingSlide.remove();
  }

  const slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);

  for (const element of slide.getPageElements()) {
    element.remove();
  }

  slide.getBackground().setSolidFill('#FFFFFF');

  Logger.log(
    `Page dimensions: ${String(pres.getPageWidth())}x${String(pres.getPageHeight())} pt (expected ${String(SLIDE_WIDTH_PT)}x${String(SLIDE_HEIGHT_PT)})`
  );

  const gridLeft = pt(in2pt(profile.gridLeftInches));
  const gridTop = pt(in2pt(profile.gridTopInches));
  const gridSize = pt(in2pt(profile.gridSizeInches));
  const cellWidth = gridSize / gridDimension;
  const thinPt = profile.thinPt;
  const thickPt = profile.thickPt;

  const { horizontalBounds, verticalBounds } = computeGridBoundaries(cages, gridDimension);

  // Title
  const titleLeft = pt(in2pt(TITLE_HORIZONTAL_MARGIN_INCHES / SIDE_COUNT));
  const titleTop = pt(in2pt(TEXT_BOX_TOP_PADDING_INCHES));
  const titleWidth = pt(SLIDE_WIDTH_PT - in2pt(TITLE_HORIZONTAL_MARGIN_INCHES));
  const titleHeight = pt(in2pt(profile.titleHeightInches));
  const titleBox = slide.insertTextBox(`${title}\n${meta}`, titleLeft, titleTop, titleWidth, titleHeight);
  const titleRange = titleBox.getText();
  titleRange.getRange(0, title.length).getTextStyle()
    .setFontFamily('Segoe UI').setFontSize(profile.titleFontSize).setBold(true).setForegroundColor(BLACK);
  if (meta.length > 0) {
    titleRange.getRange(title.length + 1, title.length + 1 + meta.length).getTextStyle()
      .setFontFamily('Segoe UI').setFontSize(profile.metaFontSize).setBold(false).setForegroundColor(VALUE_GRAY);
  }
  titleRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  // Value + Candidates boxes
  renderValueAndCandidateBoxes(slide, gridLeft, gridTop, cellWidth, gridDimension, profile);

  // Draw order
  drawThinGrid(slide, gridLeft, gridTop, gridSize, gridDimension, thinPt, verticalBounds, horizontalBounds);
  drawCageBoundaries(slide, gridLeft, gridTop, gridSize, gridDimension, thickPt, verticalBounds, horizontalBounds);
  drawJoinSquares(slide, gridLeft, gridTop, gridSize, gridDimension, thickPt, verticalBounds, horizontalBounds);
  drawOuterBorder(slide, gridLeft, gridTop, gridSize, thickPt);
  drawAxisLabels(slide, gridLeft, gridTop, cellWidth, gridDimension, profile);
  drawCageLabels(slide, gridLeft, gridTop, cellWidth, cages, hasOperators, profile);

  // Footer
  const footerLeft = pt(in2pt(TITLE_HORIZONTAL_MARGIN_INCHES));
  const footerTop = pt(SLIDE_HEIGHT_PT - in2pt(FOOTER_OFFSET_INCHES));
  const footerWidth = pt(SLIDE_WIDTH_PT - in2pt(SIDE_COUNT * TITLE_HORIZONTAL_MARGIN_INCHES));
  const footerHeight = pt(in2pt(FOOTER_HEIGHT_INCHES));
  const footerBox = slide.insertTextBox(FOOTER_TEXT, footerLeft, footerTop, footerWidth, footerHeight);
  footerBox.getText().getTextStyle()
    .setFontFamily('Segoe UI').setFontSize(FOOTER_FONT_SIZE).setBold(false).setForegroundColor(FOOTER_COLOR);
  footerBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.END);

  // Solve notes columns
  renderSolveNotesColumns(slide, gridTop, gridSize, profile);

  // Post-scale
  scaleIfNeeded(pres, slide, gridTop, gridSize, profile);

  // Init slides via Puzzle
  const renderer = new SlidesRenderer();
  const puzzleObj = new Puzzle(renderer, gridDimension, cages, hasOperators, title, meta);
  const initBatches = puzzleObj.buildInitChanges();
  for (const batch of initBatches) {
    puzzleObj.applyChanges(batch);
    puzzleObj.commit();
  }

  Logger.log(
    `Import complete: ${String(gridDimension)}x${String(gridDimension)} grid, pageW=${String(pres.getPageWidth())} pageH=${String(pres.getPageHeight())}`
  );
}

function in2pt(inches: number): number {
  return inches * POINTS_PER_INCH;
}

function init(): void {
  try {
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
    let puzzleData: PuzzleJson;
    try {
      puzzleData = JSON.parse(text) as PuzzleJson;
    } catch (_e) {
      SlidesApp.getUi().alert('Invalid puzzle data.');
      return;
    }
    initShape.remove();
    importPuzzle(puzzleData);
    props.setProperty('mathdokuInitialized', 'true');
    addMathdokuMenu();

    const allSlides = pres.getSlides();
    const lastInitSlide = allSlides[allSlides.length - 1];
    if (lastInitSlide) {
      lastInitSlide.selectAsCurrentPage();
    }
  } catch (e: unknown) {
    showError('Init', e);
  }
}

function isColorEqual(color: GoogleAppsScript.Slides.Color, hex: string): boolean {
  return colorToHex(color) === hex.toUpperCase();
}

function makeNextSlide(gridSize: number): void {
  const slide = getCurrentSlide();
  const newSlide = slide.duplicate();

  for (const element of newSlide.getPageElements()) {
    const title = element.getTitle();
    if (!title) {
      continue;
    }

    try {
      if (title.startsWith('VALUE_')) {
        finalizeValueShape(element.asShape());
      } else if (title.startsWith('CANDIDATES_')) {
        finalizeCandidatesShape(element.asShape(), gridSize);
      }
    } catch (_e) { /* Skip shapes that don't support text */ }
  }

  newSlide.selectAsCurrentPage();
}

function onOpen(): void {
  addMathdokuMenu();
}

function opSymbol(op: string): string {
  const map: Record<string, string> = {
    '-': '\u2212',
    '*': 'x',
    '/': '/',
    '+': '+',
    '\u00f7': '/',
    '\u2212': '\u2212',
    'X': 'x',
    'x': 'x'
  };
  return map[op.trim()] ?? op.trim();
}

function pt(x: number): number {
  return Math.round(x);
}

function readCellFromSlide(
  slide: GoogleAppsScript.Slides.Slide,
  ref: string,
  values: Map<string, number>,
  candidates: Map<string, Set<number>>
): void {
  const valueShape = getShapeByTitle(slide, `VALUE_${ref}`);
  if (valueShape) {
    const text = valueShape.getText().asString().replace(/\n$/, '').trim();
    if (/^[1-9]$/.test(text)) {
      values.set(ref, parseInt(text, 10));
    }
  }

  const candShape = getShapeByTitle(slide, `CANDIDATES_${ref}`);
  if (!candShape) {
    return;
  }

  const textRange = candShape.getText();
  const fullText = textRange.asString().replace(/\n$/, '');
  if (fullText.trim().length === 0 || fullText.trim() === ' ') {
    return;
  }

  const cands = new Set<number>();
  for (let i = 0; i < fullText.length; i++) {
    const ch = fullText.charAt(i);
    if (ch < '1' || ch > '9') {
      continue;
    }
    const style = textRange.getRange(i, i + 1).getTextStyle();
    if (!style.isStrikethrough()) {
      cands.add(parseInt(ch, 10));
    }
  }

  if (cands.size > 0) {
    candidates.set(ref, cands);
  }
}

function renderSolveNotesColumns(
  slide: GoogleAppsScript.Slides.Slide,
  gridTop: number,
  gridSize: number,
  profile: LayoutProfile
): void {
  const solveProfile = profile.solve;
  const notesLeft = in2pt(solveProfile.leftInches);
  const columnWidth = in2pt(solveProfile.columnWidthInches);
  const columnGap = in2pt(solveProfile.columnGapInches);
  for (let i = 0; i < solveProfile.columnCount; i++) {
    const noteBox = slide.insertTextBox(' ', pt(notesLeft + i * (columnWidth + columnGap)), pt(gridTop), pt(columnWidth), pt(gridSize));
    noteBox.setTitle(`SOLVE_NOTES_COL${String(i + 1)}`);
    noteBox.getText().getTextStyle()
      .setFontFamily('Segoe UI').setFontSize(solveProfile.font).setBold(false).setForegroundColor(VALUE_GRAY);
    noteBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
    noteBox.setContentAlignment(SlidesApp.ContentAlignment.TOP);
    noteBox.getBorder().getLineFill().setSolidFill(LIGHT_GRAY_BORDER);
    noteBox.getBorder().setWeight(1);
  }
}

function renderValueAndCandidateBoxes(
  slide: GoogleAppsScript.Slides.Slide,
  gridLeft: number,
  gridTop: number,
  cellWidth: number,
  gridDimension: number,
  profile: LayoutProfile
): void {
  const valueProfile = profile.value;
  const candidatesProfile = profile.candidates;

  for (let r = 0; r < gridDimension; r++) {
    for (let c = 0; c < gridDimension; c++) {
      const cellLeft = gridLeft + c * cellWidth;
      const cellTop = gridTop + r * cellWidth;
      const ref = cellRefA1(r, c);

      const valueBox = slide.insertTextBox(
        ' ',
        pt(cellLeft),
        pt(cellTop + valueProfile.topFraction * cellWidth),
        pt(cellWidth),
        pt(valueProfile.heightFraction * cellWidth)
      );
      valueBox.setTitle(`VALUE_${ref}`);
      valueBox.getText().getTextStyle()
        .setFontFamily('Segoe UI').setFontSize(valueProfile.font).setBold(true).setForegroundColor(VALUE_GRAY);
      valueBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      valueBox.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

      const candidatesBox = slide.insertTextBox(
        ' ',
        pt(cellLeft + candidatesProfile.leftFraction * cellWidth),
        pt(cellTop + candidatesProfile.topFraction * cellWidth),
        pt(candidatesProfile.widthFraction * cellWidth),
        pt(candidatesProfile.heightFraction * cellWidth + SIDE_COUNT * TEXT_BOX_TOP_PADDING_PT)
      );
      candidatesBox.setTitle(`CANDIDATES_${ref}`);
      candidatesBox.getText().getTextStyle()
        .setFontFamily(CANDIDATES_FONT).setFontSize(candidatesProfile.font).setBold(false).setForegroundColor(CANDIDATES_DARK_RED);
      candidatesBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
      candidatesBox.setContentAlignment(SlidesApp.ContentAlignment.BOTTOM);
    }
  }
}

function scaleIfNeeded(
  pres: GoogleAppsScript.Slides.Presentation,
  slide: GoogleAppsScript.Slides.Slide,
  gridTop: number,
  gridSize: number,
  profile: LayoutProfile
): void {
  const pageWidth = pres.getPageWidth();
  const pageHeight = pres.getPageHeight();
  const solveProfile = profile.solve;
  const notesLeft = in2pt(solveProfile.leftInches);
  const columnWidth = in2pt(solveProfile.columnWidthInches);
  const columnGap = in2pt(solveProfile.columnGapInches);
  const contentRight = notesLeft + solveProfile.columnCount * columnWidth + (solveProfile.columnCount - 1) * columnGap;
  const contentBottom = Math.max(gridTop + gridSize, SLIDE_HEIGHT_PT - in2pt(FOOTER_OFFSET_INCHES));
  const margin = SCALE_MARGIN_PT;
  const needScale = pageWidth < SLIDE_WIDTH_PT - SCALE_TOLERANCE_PT || pageHeight < SLIDE_HEIGHT_PT - SCALE_TOLERANCE_PT
    || contentRight > pageWidth - margin || contentBottom > pageHeight - margin;
  const horizontalScale = (pageWidth - margin) / contentRight;
  const verticalScale = (pageHeight - margin) / contentBottom;
  const finalScale = needScale ? Math.min(1, pageWidth / SLIDE_WIDTH_PT, pageHeight / SLIDE_HEIGHT_PT, horizontalScale, verticalScale) : 1;
  if (needScale && finalScale < 1) {
    scaleSlideElements(slide, finalScale);
  }
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
      try {
        for (const run of shape.getText().getRuns()) {
          const fontSize = run.getTextStyle().getFontSize();
          if (fontSize) {
            run.getTextStyle().setFontSize(Math.max(MIN_FONT_SIZE, Math.round(fontSize * scale)));
          }
        }
      } catch (_e) { /* GetText throws on shapes without text (e.g. grid rectangles) */ }
      try {
        const borderWeight = shape.getBorder().getWeight();
        if (borderWeight > 0) {
          shape.getBorder().setWeight(borderWeight * scale);
        }
      } catch (_e) { /* GetBorder may throw when no border is set */ }
    } else if (type === SlidesApp.PageElementType.LINE) {
      const line = element.asLine();
      const lineWeight = line.getWeight();
      line.setWeight(Math.max(1, lineWeight * scale));
    }
  }
}

function showError(source: string, e: unknown): void {
  const message = e instanceof Error ? e.message : String(e);
  const stack = e instanceof Error ? (e.stack ?? '') : '';
  Logger.log(`${source}: ${message}\n${stack}`);
  SlidesApp.getUi().alert(`${source}: ${message}`);
}

const AXIS_LABEL_MAGENTA = '#C800C8';
const BLACK = '#000000';
const CAGE_LABEL_BLUE = '#3232C8';
const CANDIDATES_DARK_RED = '#8B0000';
const CANDIDATES_FONT = 'Consolas';
const CANDIDATE_ROW_COUNT = 2;
const ENTER_DIALOG_HEIGHT_PX = 140;
const ENTER_DIALOG_WIDTH_PX = 400;
const FONT_FIT_HEIGHT_RATIO = 1.15;
const FONT_FIT_PADDING_PT = 2;
const FONT_FIT_WIDTH_RATIO = 0.60;
const FOOTER_COLOR = '#6E7887';
const FOOTER_FONT_SIZE = 14;
const FOOTER_HEIGHT_INCHES = 0.3;
const FOOTER_OFFSET_INCHES = 0.45;
const FOOTER_TEXT = '@mnaoumov';
const GREEN = '#00B050';
/* eslint-disable no-magic-numbers -- Canonical layout spec values. */
const LAYOUT_PROFILES: Record<number, LayoutProfile> = {
  4: {
    axisFont: 24,
    axisLabelHeight: 0.34,
    axisLabelWidth: 0.30,
    axisSideOffset: 0.36,
    axisTopOffset: 0.42,
    cage: { boxHeightFraction: 0.35, boxWidthFraction: 0.65, font: 28, insetLeftFraction: 0.07, insetTopFraction: 0.05 },
    candidates: { digitMargin: 12, font: 22, heightFraction: 0.60, leftFraction: 0.15, topFraction: 0.38, widthFraction: 0.80 },
    gridLeftInches: 0.65,
    gridSizeInches: 4.75,
    gridTopInches: 1.35,
    metaFontSize: 20,
    solve: { columnCount: 1, columnGapInches: 0.25, columnWidthInches: 3.25, font: 16, leftInches: 6.20 },
    thickPt: 5.0,
    thinPt: 1.0,
    titleFontSize: 30,
    titleHeightInches: 0.85,
    value: { font: 52, heightFraction: 0.70, topFraction: 0.30 }
  },
  5: {
    axisFont: 26,
    axisLabelHeight: 0.37,
    axisLabelWidth: 0.32,
    axisSideOffset: 0.38,
    axisTopOffset: 0.45,
    cage: { boxHeightFraction: 0.33, boxWidthFraction: 0.65, font: 24, insetLeftFraction: 0.07, insetTopFraction: 0.05 },
    candidates: { digitMargin: 7, font: 20, heightFraction: 0.62, leftFraction: 0.05, topFraction: 0.36, widthFraction: 0.88 },
    gridLeftInches: 0.65,
    gridSizeInches: 5.20,
    gridTopInches: 1.25,
    metaFontSize: 18,
    solve: { columnCount: 1, columnGapInches: 0.25, columnWidthInches: 3.10, font: 16, leftInches: 6.55 },
    thickPt: 5.0,
    thinPt: 1.0,
    titleFontSize: 26,
    titleHeightInches: 0.70,
    value: { font: 44, heightFraction: 0.72, topFraction: 0.28 }
  },
  6: {
    axisFont: 22,
    axisLabelHeight: 0.32,
    axisLabelWidth: 0.28,
    axisSideOffset: 0.34,
    axisTopOffset: 0.41,
    cage: { boxHeightFraction: 0.30, boxWidthFraction: 0.65, font: 22, insetLeftFraction: 0.07, insetTopFraction: 0.05 },
    candidates: { digitMargin: 5, font: 19, heightFraction: 0.65, leftFraction: 0.07, topFraction: 0.33, widthFraction: 0.86 },
    gridLeftInches: 0.65,
    gridSizeInches: 5.70,
    gridTopInches: 1.15,
    metaFontSize: 16,
    solve: { columnCount: 1, columnGapInches: 0.25, columnWidthInches: 2.95, font: 16, leftInches: 6.85 },
    thickPt: 6.5,
    thinPt: 1.0,
    titleFontSize: 24,
    titleHeightInches: 0.65,
    value: { font: 38, heightFraction: 0.75, topFraction: 0.25 }
  },
  7: {
    axisFont: 28,
    axisLabelHeight: 0.40,
    axisLabelWidth: 0.35,
    axisSideOffset: 0.42,
    axisTopOffset: 0.49,
    cage: { boxHeightFraction: 0.28, boxWidthFraction: 0.65, font: 20, insetLeftFraction: 0.07, insetTopFraction: 0.05 },
    candidates: { digitMargin: 1, font: 18, heightFraction: 0.67, leftFraction: 0.08, topFraction: 0.31, widthFraction: 0.84 },
    gridLeftInches: 0.65,
    gridSizeInches: 6.05,
    gridTopInches: 1.10,
    metaFontSize: 14,
    solve: { columnCount: 1, columnGapInches: 0.25, columnWidthInches: 2.85, font: 16, leftInches: 7.05 },
    thickPt: 6.5,
    thinPt: 1.0,
    titleFontSize: 22,
    titleHeightInches: 0.55,
    value: { font: 32, heightFraction: 0.77, topFraction: 0.23 }
  },
  8: {
    axisFont: 28,
    axisLabelHeight: 0.40,
    axisLabelWidth: 0.35,
    axisSideOffset: 0.42,
    axisTopOffset: 0.49,
    cage: { boxHeightFraction: 0.26, boxWidthFraction: 0.65, font: 18, insetLeftFraction: 0.07, insetTopFraction: 0.05 },
    candidates: { digitMargin: 1, font: 17, heightFraction: 0.69, leftFraction: 0.08, topFraction: 0.29, widthFraction: 0.84 },
    gridLeftInches: 0.65,
    gridSizeInches: 6.20,
    gridTopInches: 1.10,
    metaFontSize: 14,
    solve: { columnCount: 1, columnGapInches: 0.25, columnWidthInches: 2.80, font: 16, leftInches: 7.15 },
    thickPt: 6.5,
    thinPt: 1.0,
    titleFontSize: 22,
    titleHeightInches: 0.55,
    value: { font: 30, heightFraction: 0.78, topFraction: 0.22 }
  },
  9: {
    axisFont: 28,
    axisLabelHeight: 0.40,
    axisLabelWidth: 0.35,
    axisSideOffset: 0.42,
    axisTopOffset: 0.49,
    cage: { boxHeightFraction: 0.24, boxWidthFraction: 0.70, font: 16, insetLeftFraction: 0.07, insetTopFraction: 0.05 },
    candidates: { digitMargin: 0, font: 14, heightFraction: 0.71, leftFraction: 0.09, topFraction: 0.27, widthFraction: 0.88 },
    gridLeftInches: 0.65,
    gridSizeInches: 6.30,
    gridTopInches: 1.10,
    metaFontSize: 14,
    solve: { columnCount: 1, columnGapInches: 0.25, columnWidthInches: 2.75, font: 16, leftInches: 7.25 },
    thickPt: 6.5,
    thinPt: 1.0,
    titleFontSize: 22,
    titleHeightInches: 0.55,
    value: { font: 28, heightFraction: 0.80, topFraction: 0.20 }
  }
};
/* eslint-enable no-magic-numbers -- End layout profiles. */
const LIGHT_GRAY_BORDER = '#C8C8C8';
const MIN_FONT_SIZE = 7;
const POINTS_PER_INCH = 72;
const PUZZLE_INIT_OBJECT_ID = 'PuzzleInitData';
const SCALE_MARGIN_PT = 20;
const SCALE_TOLERANCE_PT = 0.5;
const SIDE_COUNT = 2;
const SLIDE_HEIGHT_PT = 540;
const SLIDE_WIDTH_PT = 960;
const TEXT_BOX_TOP_PADDING_INCHES = 0.05;
const TEXT_BOX_TOP_PADDING_PT = TEXT_BOX_TOP_PADDING_INCHES * POINTS_PER_INCH;
const THIN_GRAY = '#AAAAAA';
const TITLE_HORIZONTAL_MARGIN_INCHES = 0.4;
const VALUE_GRAY = '#3C414B';
