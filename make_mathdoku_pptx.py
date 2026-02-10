from __future__ import annotations

import re
import sys
from dataclasses import dataclass
from pathlib import Path

import yaml
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE, MSO_CONNECTOR
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Emu, Inches, Pt

# Slide size: widescreen 16:9
SLIDE_W_IN = 13.333
SLIDE_H_IN = 7.5

# Colors
AXIS_LABEL_MAGENTA = RGBColor(200, 0, 200)
CAGE_LABEL_BLUE = RGBColor(30, 90, 200)
VALUE_GRAY = RGBColor(60, 65, 75)
DEFAULT_CANDIDATES_DARK_RED = RGBColor(139, 0, 0)

THIN_GRID_PT_DEFAULT = 1.0
THICK_BORDER_PT_DEFAULT = 6.5


@dataclass(frozen=True)
class Cage:
    cells: tuple[tuple[int, int], ...]  # (r, c) 0-based
    label: str


CELL_RE = re.compile(r"^([A-Z])([1-9][0-9]*)$")


def _font(run, *, name: str = "Segoe UI", size: int, bold: bool = False, rgb: RGBColor) -> None:
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = rgb


def _set_slide_size(prs: Presentation) -> None:
    prs.slide_width = Inches(SLIDE_W_IN)
    prs.slide_height = Inches(SLIDE_H_IN)


def _lock_shape(shape, *, move: bool, resize: bool, rotate: bool, select: bool, text_edit: bool) -> None:
    """
    Apply DrawingML `a:spLocks` to shape.

    NOTE: `select=True` means "lock selection" (noSelect=1). Avoid for shapes VBA must access.
    """
    sp = getattr(shape, "_element", None)
    if sp is None:
        return
    cNvSpPr = sp.find(".//" + qn("p:cNvSpPr"))
    if cNvSpPr is None:
        return

    spLocks = cNvSpPr.find(qn("a:spLocks"))
    if spLocks is None:
        spLocks = OxmlElement("a:spLocks")
        cNvSpPr.append(spLocks)

    spLocks.set("noMove", "1" if move else "0")
    spLocks.set("noResize", "1" if resize else "0")
    spLocks.set("noRot", "1" if rotate else "0")
    spLocks.set("noSelect", "1" if select else "0")
    spLocks.set("noTextEdit", "1" if text_edit else "0")


def _add_title(slide, *, title: str, meta: str, n: int) -> None:
    if n >= 7:
        top, h, title_sz, meta_sz = 0.05, 0.55, 22, 14
    else:
        top, h, title_sz, meta_sz = 0.05, 0.90, 30, 20

    box = slide.shapes.add_textbox(Inches(0.2), Inches(top), Inches(SLIDE_W_IN - 0.4), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.vertical_anchor = MSO_ANCHOR.TOP

    p1 = tf.paragraphs[0]
    p1.text = title
    p1.alignment = PP_ALIGN.CENTER
    _font(p1.runs[0], size=title_sz, bold=True, rgb=RGBColor(0, 0, 0))

    p2 = tf.add_paragraph()
    p2.text = meta
    p2.alignment = PP_ALIGN.CENTER
    _font(p2.runs[0], size=meta_sz, bold=False, rgb=RGBColor(0, 0, 0))


def _add_footer(slide, *, text: str) -> None:
    box = slide.shapes.add_textbox(Inches(0.4), Inches(SLIDE_H_IN - 0.45), Inches(SLIDE_W_IN - 0.8), Inches(0.3))
    tf = box.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = text
    p.alignment = PP_ALIGN.RIGHT
    _font(p.runs[0], size=14, bold=False, rgb=RGBColor(110, 120, 135))


def _add_hidden_meta(slide, *, puzzle_id: str, n: int, operations: bool) -> None:
    """
    Store metadata off-slide for VBA.
    Must NOT be locked for selection, or VBA access may error.
    """
    meta = slide.shapes.add_textbox(Inches(SLIDE_W_IN + 1.0), Inches(SLIDE_H_IN + 1.0), Inches(0.2), Inches(0.2))
    meta.name = "MATHDOKU_META"
    meta.fill.background()
    meta.line.fill.background()
    tf = meta.text_frame
    tf.clear()
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.text = (
        "mathdoku:\n"
        f"id: {puzzle_id}\n"
        f"size: {n}\n"
        f"operations: {str(bool(operations)).lower()}\n"
    )
    p.alignment = PP_ALIGN.LEFT
    if not p.runs:
        p.add_run()
    _font(p.runs[0], size=1, bold=False, rgb=RGBColor(255, 255, 255))
    _lock_shape(meta, move=True, resize=True, rotate=True, select=False, text_edit=True)


def _add_axis_labels(slide, *, grid_left: float, grid_top: float, cell_w: float, n: int) -> None:
    if n >= 7:
        top_offset = min(0.55, max(0.38, cell_w * 0.28))
        side_offset = min(0.48, max(0.28, cell_w * 0.22))
    else:
        top_offset = min(0.45, max(0.26, cell_w * 0.22))
        side_offset = min(0.40, max(0.22, cell_w * 0.20))

    for c in range(n):
        x = grid_left + c * cell_w
        box = slide.shapes.add_textbox(Inches(x), Inches(grid_top - top_offset), Inches(cell_w), Inches(0.25))
        tf = box.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        p.text = chr(ord("A") + c)
        p.alignment = PP_ALIGN.CENTER
        _font(p.runs[0], size=16, bold=True, rgb=AXIS_LABEL_MAGENTA)
        box.line.fill.background()
        box.fill.background()
        _lock_shape(box, move=True, resize=True, rotate=True, select=True, text_edit=True)

    for r in range(n):
        y = grid_top + r * cell_w + (cell_w - 0.25) / 2
        box = slide.shapes.add_textbox(Inches(grid_left - side_offset), Inches(y), Inches(0.25), Inches(0.25))
        tf = box.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        p.text = str(r + 1)
        p.alignment = PP_ALIGN.CENTER
        _font(p.runs[0], size=16, bold=True, rgb=AXIS_LABEL_MAGENTA)
        box.line.fill.background()
        box.fill.background()
        _lock_shape(box, move=True, resize=True, rotate=True, select=True, text_edit=True)


def _add_cell_value_box(slide, *, left: float, top: float, size: float, name: str, font_size: int) -> None:
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(size), Inches(size))
    box.name = name
    box.fill.background()
    box.line.fill.background()

    tf = box.text_frame
    tf.clear()
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = Inches(0.0)
    tf.margin_right = Inches(0.0)
    tf.margin_top = Inches(0.0)
    tf.margin_bottom = Inches(0.0)
    tf.word_wrap = False

    p = tf.paragraphs[0]
    p.text = " "
    p.alignment = PP_ALIGN.CENTER
    _font(p.runs[0], size=font_size, bold=True, rgb=VALUE_GRAY)

    # Keep accessible to VBA. Prevent accidental movement.
    _lock_shape(box, move=True, resize=True, rotate=True, select=False, text_edit=False)


def _add_cell_candidates_box(
    slide,
    *,
    left: float,
    top: float,
    width: float,
    height: float,
    name: str,
    font_size: int,
    rgb: RGBColor,
) -> None:
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    box.name = name
    box.fill.background()
    box.line.fill.background()

    tf = box.text_frame
    tf.clear()
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = Inches(0.0)
    tf.margin_right = Inches(0.0)
    tf.margin_top = Inches(0.0)
    tf.margin_bottom = Inches(0.0)

    p = tf.paragraphs[0]
    p.text = " "
    p.alignment = PP_ALIGN.LEFT
    _font(p.runs[0], name="Consolas", size=font_size, bold=False, rgb=rgb)

    _lock_shape(box, move=True, resize=True, rotate=True, select=False, text_edit=False)


def _add_cell_hitbox(slide, *, left: float, top: float, size: float, name: str) -> None:
    box = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(left), Inches(top), Inches(size), Inches(size))
    box.name = name
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(255, 255, 255)
    try:
        box.fill.transparency = 1.0
    except Exception:
        pass
    box.line.fill.background()
    _lock_shape(box, move=True, resize=True, rotate=True, select=False, text_edit=True)


def _op_symbol(op: str) -> str:
    op = op.strip()
    return {
        "+": "+",
        "-": "−",
        "−": "−",
        "*": "×",
        "x": "×",
        "X": "×",
        "×": "×",
        "/": "÷",
        "÷": "÷",
    }.get(op, op)


def _parse_cell(token: str, *, n: int) -> tuple[int, int]:
    m = CELL_RE.match(token.strip().upper())
    if not m:
        raise ValueError(f"Bad cell ref: {token!r} (expected like A1)")
    col_ch, row_s = m.group(1), m.group(2)
    c = ord(col_ch) - ord("A")
    r = int(row_s) - 1
    if not (0 <= r < n and 0 <= c < n):
        raise ValueError(f"Cell out of range for {n}x{n}: {token!r}")
    return (r, c)


def _fit_font_size_for_box(*, text: str, base_pt: int, box_w_in: float, box_h_in: float, min_pt: int = 7) -> int:
    t = (text or "").strip()
    if not t:
        return base_pt
    max_w_pt = max(1.0, box_w_in * 72.0 - 2.0)
    max_h_pt = max(1.0, box_h_in * 72.0 - 1.0)
    n = max(1, len(t))
    width_based = int(max_w_pt / (0.60 * n))
    height_based = int(max_h_pt / 1.15)
    return max(min_pt, min(base_pt, width_based, height_based))


def _compute_boundaries(*, cell_to_cage: dict[tuple[int, int], int], n: int) -> tuple[list[list[bool]], list[list[bool]]]:
    v_bound: list[list[bool]] = [[False for _c in range(1, n)] for _r in range(n)]
    h_bound: list[list[bool]] = [[False for _c in range(n)] for _r in range(1, n)]
    for r in range(n):
        for c in range(1, n):
            v_bound[r][c - 1] = cell_to_cage[(r, c - 1)] != cell_to_cage[(r, c)]
    for r in range(1, n):
        for c in range(n):
            h_bound[r - 1][c] = cell_to_cage[(r - 1, c)] != cell_to_cage[(r, c)]
    return v_bound, h_bound


def _draw_segment(slide, *, x1: float, y1: float, x2: float, y2: float, width_pt: float, rgb: RGBColor) -> None:
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    ln.line.color.rgb = rgb
    ln.line.width = Pt(width_pt)
    _lock_shape(ln, move=True, resize=True, rotate=True, select=True, text_edit=True)


def _draw_thick_rect(slide, *, left: float, top: float, width: float, height: float) -> None:
    EMU_PER_INCH = 914400
    l = int(round(left * EMU_PER_INCH))
    t = int(round(top * EMU_PER_INCH))
    w = int(round(width * EMU_PER_INCH))
    h = int(round(height * EMU_PER_INCH))
    if w <= 0 or h <= 0:
        return
    rect = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Emu(l), Emu(t), Emu(w), Emu(h))
    rect.fill.solid()
    rect.fill.fore_color.rgb = RGBColor(0, 0, 0)
    rect.line.fill.background()
    _lock_shape(rect, move=True, resize=True, rotate=True, select=True, text_edit=True)


def _draw_thin_internal_grid(
    slide,
    *,
    grid_left: float,
    grid_top: float,
    grid_size: float,
    n: int,
    thin_pt: float,
    v_bound: list[list[bool]],
    h_bound: list[list[bool]],
) -> None:
    cell_w = grid_size / n
    thin_rgb = RGBColor(140, 140, 140)
    for r in range(n):
        for c in range(1, n):
            if v_bound[r][c - 1]:
                continue
            x = grid_left + c * cell_w
            y1 = grid_top + r * cell_w
            y2 = y1 + cell_w
            _draw_segment(slide, x1=x, y1=y1, x2=x, y2=y2, width_pt=thin_pt, rgb=thin_rgb)
    for r in range(1, n):
        for c in range(n):
            if h_bound[r - 1][c]:
                continue
            y = grid_top + r * cell_w
            x1 = grid_left + c * cell_w
            x2 = x1 + cell_w
            _draw_segment(slide, x1=x1, y1=y, x2=x2, y2=y, width_pt=thin_pt, rgb=thin_rgb)


def _draw_cage_boundaries(
    slide,
    *,
    grid_left: float,
    grid_top: float,
    grid_size: float,
    n: int,
    cage_pt: float,
    v_bound: list[list[bool]],
    h_bound: list[list[bool]],
) -> None:
    cell_w = grid_size / n
    thick_w = cage_pt / 72.0
    inset = cage_pt / 144.0

    # vertical runs
    for c in range(1, n):
        r0 = 0
        while r0 < n:
            if not v_bound[r0][c - 1]:
                r0 += 1
                continue
            r1 = r0
            while r1 + 1 < n and v_bound[r1 + 1][c - 1]:
                r1 += 1
            x = grid_left + c * cell_w
            y1 = grid_top + r0 * cell_w
            y2 = grid_top + (r1 + 1) * cell_w
            if r0 == 0:
                y1 += inset
            if r1 == n - 1:
                y2 -= inset
            _draw_thick_rect(slide, left=x - thick_w / 2, top=y1, width=thick_w, height=(y2 - y1))
            r0 = r1 + 1

    # horizontal runs
    for r in range(1, n):
        c0 = 0
        while c0 < n:
            if not h_bound[r - 1][c0]:
                c0 += 1
                continue
            c1 = c0
            while c1 + 1 < n and h_bound[r - 1][c1 + 1]:
                c1 += 1
            y = grid_top + r * cell_w
            x1 = grid_left + c0 * cell_w
            x2 = grid_left + (c1 + 1) * cell_w
            if c0 == 0:
                x1 += inset
            if c1 == n - 1:
                x2 -= inset
            _draw_thick_rect(slide, left=x1, top=y - thick_w / 2, width=(x2 - x1), height=thick_w)
            c0 = c1 + 1


def _draw_thick_join_squares(
    slide,
    *,
    grid_left: float,
    grid_top: float,
    grid_size: float,
    n: int,
    cage_pt: float,
    v_bound: list[list[bool]],
    h_bound: list[list[bool]],
) -> None:
    cell_w = grid_size / n
    thick_w = cage_pt / 72.0

    def clamp(v: float, lo: float, hi: float) -> float:
        return lo if v < lo else hi if v > hi else v

    for vr in range(1, n):
        for vc in range(1, n):
            touch = False
            if v_bound[vr - 1][vc - 1] or v_bound[vr][vc - 1]:
                touch = True
            if h_bound[vr - 1][vc - 1] or h_bound[vr - 1][vc]:
                touch = True
            if not touch:
                continue
            x = grid_left + vc * cell_w
            y = grid_top + vr * cell_w
            left = clamp(x - thick_w / 2, grid_left, grid_left + grid_size - thick_w)
            top = clamp(y - thick_w / 2, grid_top, grid_top + grid_size - thick_w)
            _draw_thick_rect(slide, left=left, top=top, width=thick_w, height=thick_w)


def _draw_outer_border(
    slide,
    *,
    grid_left: float,
    grid_top: float,
    grid_size: float,
    cage_pt: float,
) -> None:
    thick_w = cage_pt / 72.0
    half = thick_w / 2.0
    # Top & bottom include corners
    _draw_thick_rect(slide, left=grid_left - half, top=grid_top - half, width=grid_size + thick_w, height=thick_w)
    _draw_thick_rect(
        slide,
        left=grid_left - half,
        top=grid_top + grid_size - half,
        width=grid_size + thick_w,
        height=thick_w,
    )
    # Left/right exclude corners
    _draw_thick_rect(slide, left=grid_left - half, top=grid_top + half, width=thick_w, height=max(0.0, grid_size - thick_w))
    _draw_thick_rect(
        slide,
        left=grid_left + grid_size - half,
        top=grid_top + half,
        width=thick_w,
        height=max(0.0, grid_size - thick_w),
    )


def _compute_layout(n: int, *, solve_enabled: bool, solve_cols: int, solve_min_col_w: float) -> dict:
    left_margin = 0.65
    right_margin = 0.40
    footer_h = 0.55

    if n >= 7:
        grid_top = 1.10
    else:
        grid_top = 1.35

    max_h = SLIDE_H_IN - footer_h - grid_top
    gap = 0.60 if solve_enabled else 0.0
    min_solve_w = (solve_cols * solve_min_col_w) if solve_enabled else 0.0
    max_w_for_grid = SLIDE_W_IN - left_margin - right_margin - gap - min_solve_w
    grid_size = max(3.0, min(max_h, max_w_for_grid))

    solve_left = left_margin + grid_size + gap
    solve_w = max(0.0, SLIDE_W_IN - solve_left - right_margin) if solve_enabled else 0.0

    return {
        "grid_left": left_margin,
        "grid_top": grid_top,
        "grid_size": grid_size,
        "solve_enabled": solve_enabled and solve_w >= (solve_cols * 0.9),
        "solve_left": solve_left,
        "solve_w": solve_w,
    }


def load_yaml_spec(path: Path) -> dict:
    data = yaml.safe_load(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError("YAML spec must be a mapping at the top level")
    return data


def cages_from_yaml(spec: dict) -> tuple[int, list[Cage], str, str, str, float, float, bool, str, dict, dict, dict]:
    n = int(spec.get("size", 4))
    difficulty = spec.get("difficulty", None)
    operations = bool(spec.get("operations", True))
    puzzle_id = str(spec.get("id", "")).strip()
    if not puzzle_id:
        raise ValueError("YAML must include non-empty 'id'")

    title = str(spec.get("title", "")).strip()
    if not title:
        try:
            title = f"#Mathdoku {int(puzzle_id):02d}"
        except Exception:
            title = f"#Mathdoku {puzzle_id}"

    meta = str(spec.get("meta", "")).strip()
    if not meta:
        parts = [f"Size {n}×{n}"]
        if difficulty is not None:
            parts.append(f"Difficulty {difficulty}")
        parts.append("With operators" if operations else "Without operators")
        meta = " • ".join(parts)

    footer = str(spec.get("footer", "@mnaoumov"))
    out_in_yaml = spec.get("output", None)

    lw = spec.get("lineWidths", {}) or {}
    thin_pt = float(lw.get("thin", THIN_GRID_PT_DEFAULT))
    thick_pt = float(lw.get("thick", THICK_BORDER_PT_DEFAULT))

    fonts = spec.get("fonts", {}) or {}
    colors = spec.get("colors", {}) or {}

    value_font = fonts.get("value", None)
    cand_font = fonts.get("cellCandidates", None)

    cand_color_hex = colors.get("cellCandidates", None)
    if cand_color_hex:
        hx = str(cand_color_hex).strip().lstrip("#")
        if len(hx) != 6:
            raise ValueError("colors.cellCandidates must be 6-digit hex")
        cand_rgb = RGBColor(int(hx[0:2], 16), int(hx[2:4], 16), int(hx[4:6], 16))
    else:
        cand_rgb = DEFAULT_CANDIDATES_DARK_RED

    solve_notes = spec.get("solveNotes", {}) or {}
    solve_enabled = bool(solve_notes.get("enabled", True))
    solve_cols = int(solve_notes.get("columns", 3))
    solve_min_col_w = float(solve_notes.get("minColWidth", 1.35))

    prefill = spec.get("prefill", {}) or {}
    prefill_values = prefill.get("values", {}) or {}
    prefill_candidates = prefill.get("cellCandidates", {}) or {}

    cages_in = spec.get("cages", None)
    if not isinstance(cages_in, list) or not cages_in:
        raise ValueError("cages must be a non-empty list")

    cages: list[Cage] = []
    for idx, item in enumerate(cages_in):
        if not isinstance(item, dict):
            raise ValueError(f"cages[{idx}] must be a mapping")
        cells_raw = item.get("cells", None)
        if not isinstance(cells_raw, list) or not cells_raw:
            raise ValueError(f"cages[{idx}].cells must be a non-empty list")
        cells = tuple(_parse_cell(str(t), n=n) for t in cells_raw)

        label = str(item.get("label", "")).strip() if "label" in item else ""
        if not label:
            if "value" not in item:
                raise ValueError(f"cages[{idx}] must have 'label' or 'value'")
            value = str(item["value"]).strip()
            op = item.get("op", item.get("operator", None))
            if operations and len(cells) > 1:
                if not op:
                    raise ValueError(f"cages[{idx}] multi-cell with operations=true requires op")
                label = f"{value}{_op_symbol(str(op))}"
            else:
                label = value

        cages.append(Cage(cells=cells, label=label))

    ui = {
        "value_font": value_font,
        "cand_font": cand_font,
        "cand_rgb": cand_rgb,
        "solve_enabled": solve_enabled,
        "solve_cols": solve_cols,
        "solve_min_col_w": solve_min_col_w,
        "prefill_values": prefill_values,
        "prefill_candidates": prefill_candidates,
    }

    return n, cages, title, meta, footer, thin_pt, thick_pt, operations, puzzle_id, out_in_yaml, ui, solve_notes


def build_pptx(*, spec_path: Path, spec: dict) -> Path:
    (
        n,
        cages,
        title,
        meta,
        footer,
        thin_pt,
        thick_pt,
        operations,
        puzzle_id,
        out_in_yaml,
        ui,
        _solve_notes,
    ) = cages_from_yaml(spec)

    out_path = None
    if out_in_yaml:
        out_path = Path(str(out_in_yaml))
        if not out_path.is_absolute():
            out_path = (spec_path.parent / out_path).resolve()
    else:
        out_path = spec_path.with_suffix(".pptx")

    prs = Presentation()
    _set_slide_size(prs)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    layout = _compute_layout(n, solve_enabled=bool(ui["solve_enabled"]), solve_cols=int(ui["solve_cols"]), solve_min_col_w=float(ui["solve_min_col_w"]))

    _add_title(slide, title=title, meta=meta, n=n)
    _add_hidden_meta(slide, puzzle_id=puzzle_id, n=n, operations=bool(operations))

    grid_left = float(layout["grid_left"])
    grid_top = float(layout["grid_top"])
    grid_size = float(layout["grid_size"])
    cell_w = grid_size / n

    # Determine fonts
    cell_pts = cell_w * 72.0
    value_font = int(ui["value_font"]) if ui["value_font"] is not None else int(max(16, min(72, round(cell_pts * 0.62))))
    cand_font = int(ui["cand_font"]) if ui["cand_font"] is not None else (16 if n <= 4 else 11)
    cand_rgb: RGBColor = ui["cand_rgb"]

    cand_pad_x = min(0.10, cell_w * 0.14)
    cand_top_y = min(0.22, cell_w * 0.28)
    cand_h = min(0.40, cell_w * 0.45)

    # Validate coverage/non-overlap
    all_cells = {(r, c) for r in range(n) for c in range(n)}
    cell_to_cage: dict[tuple[int, int], int] = {}
    for i, cage in enumerate(cages):
        for cell in cage.cells:
            if cell in cell_to_cage:
                raise ValueError(f"Cell {cell} appears in multiple cages")
            cell_to_cage[cell] = i
    missing = sorted(all_cells - set(cell_to_cage))
    if missing:
        raise ValueError(f"Spec missing {len(missing)} cell(s): {missing}")

    # Create cell content + hitboxes
    for r in range(n):
        for c in range(n):
            cell_left = grid_left + c * cell_w
            cell_top = grid_top + r * cell_w
            _add_cell_value_box(slide, left=cell_left, top=cell_top, size=cell_w, name=f"VALUE_r{r+1}c{c+1}", font_size=value_font)
            _add_cell_candidates_box(
                slide,
                left=cell_left + cand_pad_x,
                top=cell_top + cand_top_y,
                width=cell_w - 2 * cand_pad_x,
                height=cand_h,
                name=f"CANDIDATES_r{r+1}c{c+1}",
                font_size=cand_font,
                rgb=cand_rgb,
            )
            _add_cell_hitbox(slide, left=cell_left, top=cell_top, size=cell_w, name=f"CELL_r{r+1}c{c+1}")

    # Axis labels
    _add_axis_labels(slide, grid_left=grid_left, grid_top=grid_top, cell_w=cell_w, n=n)

    # Borders
    v_bound, h_bound = _compute_boundaries(cell_to_cage=cell_to_cage, n=n)
    _draw_thin_internal_grid(slide, grid_left=grid_left, grid_top=grid_top, grid_size=grid_size, n=n, thin_pt=thin_pt, v_bound=v_bound, h_bound=h_bound)
    _draw_cage_boundaries(slide, grid_left=grid_left, grid_top=grid_top, grid_size=grid_size, n=n, cage_pt=thick_pt, v_bound=v_bound, h_bound=h_bound)
    _draw_thick_join_squares(slide, grid_left=grid_left, grid_top=grid_top, grid_size=grid_size, n=n, cage_pt=thick_pt, v_bound=v_bound, h_bound=h_bound)
    _draw_outer_border(slide, grid_left=grid_left, grid_top=grid_top, grid_size=grid_size, cage_pt=thick_pt)

    # Cage labels (one per cage, top-left cell)
    inset_x = min(0.05, cell_w * 0.12)
    inset_y = min(0.04, cell_w * 0.10)
    cage_font_base = int(max(9, min(18, round(cell_pts * 0.26))))
    for i, cage in enumerate(cages):
        tl = min(cage.cells, key=lambda rc: (rc[0], rc[1]))
        r, c = tl
        x = grid_left + c * cell_w + inset_x
        y = grid_top + r * cell_w + inset_y
        label_box_w = min(1.05, cell_w * 0.90)
        label_box_h = min(0.34, cell_w * 0.30)
        box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(label_box_w), Inches(label_box_h))
        box.name = f"CAGE_{i}_r{r+1}c{c+1}"
        box.fill.background()
        box.line.fill.background()
        tf = box.text_frame
        tf.clear()
        tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        tf.margin_left = Inches(0.0)
        tf.margin_right = Inches(0.0)
        tf.margin_top = Inches(0.0)
        tf.margin_bottom = Inches(0.0)
        p = tf.paragraphs[0]
        p.text = cage.label
        p.alignment = PP_ALIGN.LEFT
        if not p.runs:
            p.add_run()
        fitted = _fit_font_size_for_box(text=cage.label, base_pt=cage_font_base, box_w_in=label_box_w, box_h_in=label_box_h, min_pt=7)
        _font(p.runs[0], size=fitted, bold=True, rgb=CAGE_LABEL_BLUE)
        _lock_shape(box, move=True, resize=True, rotate=True, select=True, text_edit=True)

    _add_footer(slide, text=footer)

    # Solve notes columns (editable)
    if bool(layout["solve_enabled"]):
        notes_left = float(layout["solve_left"])
        notes_right_margin = 0.4
        avail_w = max(0.5, SLIDE_W_IN - notes_left - notes_right_margin)
        col_w = avail_w / int(ui["solve_cols"])
        notes_top = grid_top
        notes_h = grid_size
        for i in range(int(ui["solve_cols"])):
            box = slide.shapes.add_textbox(Inches(notes_left + i * col_w), Inches(notes_top), Inches(col_w - 0.15), Inches(notes_h))
            box.name = f"SOLVE_NOTES_COL{i+1}"
            box.fill.background()
            box.line.color.rgb = RGBColor(200, 200, 200)
            box.line.width = Pt(1)
            tf = box.text_frame
            tf.clear()
            tf.vertical_anchor = MSO_ANCHOR.TOP
            tf.word_wrap = True
            tf.margin_left = Inches(0.08)
            tf.margin_right = Inches(0.08)
            tf.margin_top = Inches(0.08)
            tf.margin_bottom = Inches(0.08)
            p = tf.paragraphs[0]
            p.text = " "
            p.alignment = PP_ALIGN.LEFT
            _font(p.runs[0], size=16, bold=False, rgb=VALUE_GRAY)
            _lock_shape(box, move=True, resize=True, rotate=True, select=False, text_edit=False)

    # Bring hitboxes to front so clicks always select a cell.
    for sh in list(slide.shapes):
        nm = getattr(sh, "name", "") or ""
        if nm.startswith("CELL_r"):
            try:
                sh.zorder(0)  # msoBringToFront
            except Exception:
                pass

    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(out_path)
    return out_path


def main(argv: list[str]) -> None:
    if len(argv) != 2:
        print("Error: this script accepts exactly ONE argument: the YAML spec path.")
        print("Usage: python make_mathdoku_pptx.py <spec.yaml>")
        raise SystemExit(2)
    if argv[1].startswith("-"):
        print("Error: CLI flags/parameters are not supported. Put all settings in the YAML file.")
        raise SystemExit(2)

    spec_path = Path(argv[1])
    spec = load_yaml_spec(spec_path)
    out = build_pptx(spec_path=spec_path, spec=spec)
    print(f"Wrote {out}")


if __name__ == "__main__":
    main(sys.argv)

