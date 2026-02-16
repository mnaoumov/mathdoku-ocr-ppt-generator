/**
 * MakeNextSlide.ts -- Duplicate current slide and finalize green changes.
 *
 * Green value      -> VALUE_GRAY
 * Green candidates -> CAND_DARK_RED
 * Green+strike     -> remove character, rebuild candidate text
 */

function finalizeCandidatesShape(shape: GoogleAppsScript.Slides.Shape, gridSize: number): void {
  const textRange = shape.getText();
  const fullText = textRange.asString().replace(/\n$/, '');
  if (fullText.trim().length === 0 || fullText.trim() === ' ') {
    return;
  }

  let hasGreenStrike = false;
  let hasGreen = false;
  const survivingDigits: string[] = [];

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

    survivingDigits.push(ch);
  }

  // If we had green+strikethrough removals, rebuild the text
  if (hasGreenStrike) {
    const newFormatted = formatCandidates(survivingDigits.join(''), gridSize);
    textRange.setText(newFormatted);
    textRange.getTextStyle()
      .setFontFamily(CANDIDATES_FONT)
      .setForegroundColor(CAND_DARK_RED)
      .setStrikethrough(false);
    return;
  }

  // If we had plain green (no strike), just change green -> CAND_DARK_RED
  if (hasGreen) {
    for (let i = 0; i < fullText.length; i++) {
      const ch = fullText.charAt(i);
      if (ch < '1' || ch > '9') {
        continue;
      }
      const style = textRange.getRange(i, i + 1).getTextStyle();
      const color = style.getForegroundColor();
      if (isColorEqual(color, GREEN)) {
        style.setForegroundColor(CAND_DARK_RED);
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

function makeNextSlide(): void {
  if (PropertiesService.getDocumentProperties().getProperty('mathdokuInitialized') !== 'true') {
    SlidesApp.getUi().alert('Please run Mathdoku > Init first.');
    return;
  }
  const slide = getCurrentSlide();
  const state = getPuzzleState();
  const n = state.size;

  // Duplicate the slide
  const newSlide = slide.duplicate();

  // Process all shapes on the NEW slide
  for (const el of newSlide.getPageElements()) {
    const title = el.getTitle();
    if (!title) {
      continue;
    }

    if (title.startsWith('VALUE_')) {
      finalizeValueShape(el.asShape());
    } else if (title.startsWith('CANDIDATES_')) {
      finalizeCandidatesShape(el.asShape(), n);
    }
  }
}
