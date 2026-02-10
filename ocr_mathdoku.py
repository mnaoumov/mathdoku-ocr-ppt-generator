"""
OCR tool for Mathdoku puzzles: screenshot -> YAML spec.

Usage:
    python ocr_mathdoku.py Blog09.png              -> Blog09.yaml
    python ocr_mathdoku.py Blog09.png --size 4     -> override grid size
    python ocr_mathdoku.py Blog09.png -o out.yaml  -> custom output
    python ocr_mathdoku.py Blog09.png --debug       -> save debug images

Requirements:
    pip install opencv-python numpy pytesseract pyyaml
    Tesseract-OCR engine:
      Windows: https://github.com/UB-Mannheim/tesseract/wiki
      macOS:   brew install tesseract
      Linux:   apt install tesseract-ocr
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

import cv2
import numpy as np
import yaml

try:
    import pytesseract
except ImportError:
    pytesseract = None

_DEBUG = False


# ── YAML formatting ────────────────────────────────────────────────────────

class _FlowList(list):
    """List that serializes in YAML flow style: [A1, B1, C1]."""


def _flow_representer(dumper: yaml.Dumper, data: _FlowList):
    return dumper.represent_sequence("tag:yaml.org,2002:seq", data, flow_style=True)


yaml.add_representer(_FlowList, _flow_representer)


# ── helpers ─────────────────────────────────────────────────────────────────

def _cell_a1(r: int, c: int) -> str:
    return f"{chr(ord('A') + c)}{r + 1}"


def _dbg(msg: str) -> None:
    if _DEBUG:
        print(f"  [debug] {msg}")


def _dbg_save(name: str, img: np.ndarray) -> None:
    if _DEBUG:
        cv2.imwrite(name, img)


# ── grid detection ──────────────────────────────────────────────────────────

def _find_grid_bbox(gray: np.ndarray) -> tuple[int, int, int, int]:
    """Detect the puzzle grid bounding rectangle. Returns (x, y, w, h)."""
    h, w = gray.shape

    def _find_via_lines() -> tuple[int, int, int, int] | None:
        _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
        min_len = min(h, w) // 3
        h_lines = cv2.morphologyEx(
            binary, cv2.MORPH_OPEN,
            cv2.getStructuringElement(cv2.MORPH_RECT, (min_len, 1)),
        )
        v_lines = cv2.morphologyEx(
            binary, cv2.MORPH_OPEN,
            cv2.getStructuringElement(cv2.MORPH_RECT, (1, min_len)),
        )
        combined = cv2.dilate(
            cv2.add(h_lines, v_lines), np.ones((5, 5), np.uint8), iterations=2,
        )
        contours, _ = cv2.findContours(combined, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        best, best_area = None, 0
        for cnt in contours:
            area = cv2.contourArea(cnt)
            if area < h * w * 0.05:
                continue
            bx, by, bw, bh = cv2.boundingRect(cnt)
            aspect = min(bw, bh) / max(bw, bh)
            if aspect < 0.70 or area <= best_area:
                continue
            best_area = area
            best = (bx, by, bw, bh)
        return best

    def _find_via_white_region() -> tuple[int, int, int, int] | None:
        """Fallback for app screenshots: find the largest white rectangular area."""
        white = (gray > 200).astype(np.uint8) * 255
        kernel = np.ones((15, 15), np.uint8)
        closed = cv2.morphologyEx(white, cv2.MORPH_CLOSE, kernel)
        contours, _ = cv2.findContours(closed, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        best, best_area = None, 0
        for cnt in contours:
            area = cv2.contourArea(cnt)
            if area < h * w * 0.03:
                continue
            bx, by, bw, bh = cv2.boundingRect(cnt)
            aspect = min(bw, bh) / max(bw, bh)
            if aspect < 0.70 or area <= best_area:
                continue
            best_area = area
            best = (bx, by, bw, bh)
        return best

    result = _find_via_lines()
    if result is None:
        _dbg("Line-based grid detection failed, trying white-region fallback")
        result = _find_via_white_region()
    if result is None:
        raise ValueError("Could not locate puzzle grid in image.")
    return result


# ── line detection via adaptive threshold + projection ──────────────────────

def _detect_line_positions(
    gray: np.ndarray, gx: int, gy: int, gw: int, gh: int,
) -> tuple[list[int], list[int]]:
    """Return candidate (h_positions, v_positions) of grid lines relative to grid origin."""
    crop = gray[gy:gy + gh, gx:gx + gw]

    # Adaptive threshold catches both thin (gray) and thick (dark) lines
    adaptive = cv2.adaptiveThreshold(
        crop, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY_INV, blockSize=15, C=5,
    )
    _dbg_save("debug_adaptive.png", adaptive)

    # Sum-projection: fraction of dark pixels per row / column
    h_proj = np.sum(adaptive > 0, axis=1).astype(float) / gw
    v_proj = np.sum(adaptive > 0, axis=0).astype(float) / gh

    def find_peaks(proj: np.ndarray, threshold: float = 0.25) -> list[int]:
        peaks: list[int] = []
        in_run, start = False, 0
        for i in range(len(proj)):
            if proj[i] > threshold and not in_run:
                start, in_run = i, True
            elif proj[i] <= threshold and in_run:
                seg = proj[start:i]
                peaks.append(start + int(np.argmax(seg)))
                in_run = False
        if in_run:
            seg = proj[start:]
            peaks.append(start + int(np.argmax(seg)))
        return peaks

    h_peaks = find_peaks(h_proj)
    v_peaks = find_peaks(v_proj)
    _dbg(f"Raw peaks: {len(h_peaks)}h {h_peaks}, {len(v_peaks)}v {v_peaks}")
    return h_peaks, v_peaks


# ── grid size selection ─────────────────────────────────────────────────────

def _regularity_score(candidates: list[int], total: int, n: int) -> float:
    """Score how well candidates fit a regular N-division grid. Lower = better."""
    if len(candidates) < 2:
        return float("inf")
    first, last = candidates[0], candidates[-1]
    spacing = (last - first) / n
    if spacing < 10:
        return float("inf")

    matched, error = 0, 0.0
    for k in range(n + 1):
        expected = first + k * spacing
        min_dist = min(abs(c - expected) for c in candidates)
        if min_dist < spacing * 0.20:
            matched += 1
            error += min_dist
        else:
            error += spacing * 0.5
    return -matched * 1000 + error / (n + 1)


def _fit_lines(candidates: list[int], total: int, n: int) -> list[int]:
    """Select N+1 evenly-spaced grid lines from candidates."""
    if len(candidates) < 2:
        return [round(i * total / n) for i in range(n + 1)]
    first, last = candidates[0], candidates[-1]
    spacing = (last - first) / n

    result: list[int] = []
    for k in range(n + 1):
        expected = first + k * spacing
        dists = [(abs(c - expected), c) for c in candidates]
        min_dist, nearest = min(dists)
        result.append(nearest if min_dist < spacing * 0.20 else int(round(expected)))
    return result


def _detect_size_and_lines(
    h_cands: list[int], v_cands: list[int],
    gh: int, gw: int, n: int | None,
) -> tuple[int, list[int], list[int]]:
    """Determine grid size and final line positions."""
    if n is not None:
        return n, _fit_lines(h_cands, gh, n), _fit_lines(v_cands, gw, n)

    best_n, best_score = 4, float("inf")
    for try_n in range(4, 10):
        score = _regularity_score(h_cands, gh, try_n) + _regularity_score(v_cands, gw, try_n)
        _dbg(f"  n={try_n}: score={score:.1f}")
        if score < best_score:
            best_score, best_n = score, try_n

    _dbg(f"Best size: {best_n}")
    return best_n, _fit_lines(h_cands, gh, best_n), _fit_lines(v_cands, gw, best_n)


# ── border classification ──────────────────────────────────────────────────

def _classify_borders(
    gray: np.ndarray,
    gx: int, gy: int,
    n: int, h_pos: list[int], v_pos: list[int],
) -> tuple[dict[tuple[int, int], bool], dict[tuple[int, int], bool]]:
    """
    Classify internal borders as thick (cage boundary) or thin (same cage).
    Metric: near-darkest pixel (255 - 5th-percentile) in a narrow strip.
    Cage borders have many dark pixels; same-cage borders have only light-gray pixels.
    Using 5th percentile instead of min to be robust to label text contamination.
    """
    gh_end = min(gy + max(h_pos) + 10, gray.shape[0])
    gw_end = min(gx + max(v_pos) + 10, gray.shape[1])
    crop = gray[gy:gh_end, gx:gw_end]

    cell_h = (h_pos[-1] - h_pos[0]) / n
    cell_w = (v_pos[-1] - v_pos[0]) / n
    margin_h = max(5, int(cell_h * 0.25))
    margin_w = max(5, int(cell_w * 0.25))
    radius = max(2, int(min(cell_h, cell_w) * 0.02))

    measurements: list[tuple[str, int, int, float]] = []

    # Horizontal internal borders
    for r in range(1, n):
        y = h_pos[r]
        y0, y1 = max(0, y - radius), min(crop.shape[0], y + radius + 1)
        for c in range(n):
            x0, x1 = v_pos[c] + margin_w, v_pos[c + 1] - margin_w
            if x0 >= x1 or y0 >= y1:
                continue
            strip = crop[y0:y1, x0:x1]
            if strip.size == 0:
                continue
            measurements.append(("h", r, c, 255.0 - float(np.percentile(strip, 10))))

    # Vertical internal borders
    for c in range(1, n):
        x = v_pos[c]
        x0, x1 = max(0, x - radius), min(crop.shape[1], x + radius + 1)
        for r in range(n):
            y0, y1 = h_pos[r] + margin_h, h_pos[r + 1] - margin_h
            if y0 >= y1 or x0 >= x1:
                continue
            strip = crop[y0:y1, x0:x1]
            if strip.size == 0:
                continue
            measurements.append(("v", r, c, 255.0 - float(np.percentile(strip, 10))))

    if not measurements:
        return {}, {}

    # Two-class separation (Otsu on mean darkness values)
    values = np.array([v for *_, v in measurements])
    _dbg(f"Border darkness: min={np.min(values):.1f} max={np.max(values):.1f} "
         f"median={np.median(values):.1f}")

    sorted_v = np.sort(np.unique(values))
    best_thresh, best_var = float(np.median(values)), -1.0
    for t in sorted_v:
        lo, hi = values[values <= t], values[values > t]
        if len(lo) == 0 or len(hi) == 0:
            continue
        var = len(lo) * len(hi) * (np.mean(hi) - np.mean(lo)) ** 2
        if var > best_var:
            best_var, best_thresh = var, float(t)

    best_thresh = max(best_thresh, 3.0)
    _dbg(f"Border threshold: {best_thresh:.1f}")

    h_thick: dict[tuple[int, int], bool] = {}
    v_thick: dict[tuple[int, int], bool] = {}
    for axis, r, c, score in measurements:
        is_thick = score > best_thresh
        (h_thick if axis == "h" else v_thick)[(r, c)] = is_thick

    if _DEBUG:
        for axis, r, c, score in measurements:
            tag = "THICK" if score > best_thresh else "thin "
            _dbg(f"  {axis}({r},{c}): {score:.3f} {tag}")

    return h_thick, v_thick


# ── cage grouping (union-find) ─────────────────────────────────────────────

def _group_cages(
    n: int,
    h_thick: dict[tuple[int, int], bool],
    v_thick: dict[tuple[int, int], bool],
) -> list[list[tuple[int, int]]]:
    parent: dict[tuple[int, int], tuple[int, int]] = {
        (r, c): (r, c) for r in range(n) for c in range(n)
    }

    def find(x: tuple[int, int]) -> tuple[int, int]:
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a: tuple[int, int], b: tuple[int, int]) -> None:
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[rb] = ra

    for r in range(n):
        for c in range(n):
            if c + 1 < n and not v_thick.get((r, c + 1), True):
                union((r, c), (r, c + 1))
            if r + 1 < n and not h_thick.get((r + 1, c), True):
                union((r, c), (r + 1, c))

    groups: dict[tuple[int, int], list[tuple[int, int]]] = {}
    for r in range(n):
        for c in range(n):
            groups.setdefault(find((r, c)), []).append((r, c))
    return [sorted(cells) for cells in groups.values()]


# ── label OCR ──────────────────────────────────────────────────────────────

_LABEL_RE = re.compile(r"^(\d+)([+\-x×÷/])?$")


def _require_tesseract() -> None:
    if pytesseract is None:
        print(
            "Error: pytesseract not installed.\n"
            "  pip install pytesseract\n"
            "  + install Tesseract engine",
            file=sys.stderr,
        )
        raise SystemExit(1)


def _trim_to_text(crop_gray: np.ndarray) -> np.ndarray:
    """Remove border artifacts and crop tightly around label text."""
    h, w = crop_gray.shape

    # 1. Binarize with fixed threshold to pick up dark label text only,
    #    ignoring gray candidate numbers (typically > 160).
    _, binary = cv2.threshold(crop_gray, 160, 255, cv2.THRESH_BINARY_INV)

    # 2. Strip continuous dark columns/rows from left/top edge (grid border
    #    remnants).  Only strip if the column/row is >90% dark across its
    #    full extent — digit strokes never span the full height/width.
    max_strip = min(w, h) // 4  # never strip more than 25% of the crop
    left = 0
    for c in range(min(max_strip, w)):
        if np.mean(binary[:, c] > 0) > 0.9:
            left = c + 1
        else:
            break
    top = 0
    for r in range(min(max_strip, h)):
        if np.mean(binary[r, :] > 0) > 0.9:
            top = r + 1
        else:
            break
    if left > 0 or top > 0:
        _dbg(f"  Stripped border: left={left}px top={top}px")
        crop_gray = crop_gray[top:, left:]
        binary = binary[top:, left:]
        h, w = crop_gray.shape
        if h < 5 or w < 5:
            return crop_gray

    # 3. Pad so contours touching the edge are properly detected
    bp = 2
    padded = cv2.copyMakeBorder(binary, bp, bp, bp, bp,
                                cv2.BORDER_CONSTANT, value=0)
    contours, _ = cv2.findContours(padded, cv2.RETR_EXTERNAL,
                                   cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        return crop_gray

    # 4. Keep contours that look like text (not thin border lines)
    text_cnts: list[np.ndarray] = []
    for cnt in contours:
        x, y, cw, ch = cv2.boundingRect(cnt)
        if cw * ch < 20:
            continue
        aspect = min(cw, ch) / max(cw, ch) if max(cw, ch) > 0 else 0
        if aspect < 0.08:
            continue
        text_cnts.append(cnt)

    if not text_cnts:
        return crop_gray

    all_pts = np.vstack(text_cnts)
    bx, by, tw, th = cv2.boundingRect(all_pts)
    bx, by = bx - bp, by - bp
    pad = 3
    bx, by = max(0, bx - pad), max(0, by - pad)
    tw = min(w - bx, tw + 2 * pad)
    th = min(h - by, th + 2 * pad)
    return crop_gray[by:by + th, bx:bx + tw]


def _prepare_ocr_image(crop_gray: np.ndarray) -> np.ndarray:
    """Prepare a grayscale crop for Tesseract: scale, binarize, pad."""
    h, w = crop_gray.shape
    if h < 80:
        scale = max(2, 80 // h)
        crop_gray = cv2.resize(crop_gray, None, fx=scale, fy=scale,
                               interpolation=cv2.INTER_CUBIC)

    _, binary = cv2.threshold(crop_gray, 0, 255,
                              cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    if np.mean(binary) < 128:
        binary = 255 - binary

    pad = 12
    return cv2.copyMakeBorder(binary, pad, pad, pad, pad,
                              cv2.BORDER_CONSTANT, value=255)


def _run_tesseract(padded: np.ndarray, configs: list[str]) -> str:
    """Try multiple Tesseract configs, pick best result via voting.

    Strategy: prefer longest digit string (avoids '11+' being outvoted by '1+'),
    then majority vote among those, then include operator if any matching result
    has one (catches operators missed by some configs).
    """
    from collections import Counter
    results: list[tuple[str, str | None]] = []  # (digits, op_or_None)
    best_raw = ""
    for cfg in configs:
        try:
            text = pytesseract.image_to_string(padded, config=cfg).strip()
        except Exception:
            continue
        text = text.replace(" ", "").replace("\n", "")
        text = text.replace("×", "x").replace("÷", "/").replace("−", "-")
        text = text.replace("X", "x")
        text = text.replace("O", "0").replace("o", "0").replace("Q", "0")
        text = text.replace("l", "1").replace("I", "1")
        m = _LABEL_RE.match(text)
        if not m:
            # Try to salvage trailing operator from garbled/unknown characters
            m_raw = re.match(r"^(\d+)(.)$", text)
            if m_raw and not m_raw.group(2).isdigit():
                ch = m_raw.group(2)
                if ch in "+-\u2212\u2013":
                    text = m_raw.group(1) + "-"
                elif ch in "/\u00f7|\\":
                    text = m_raw.group(1) + "/"
                else:
                    text = m_raw.group(1) + "x"
                m = _LABEL_RE.match(text)
        if m:
            results.append((m.group(1), m.group(2)))
        elif len(text) > len(best_raw):
            best_raw = text

    if not results:
        return best_raw

    # Group results by digit-string length, then pick the best group.
    # Prefer the longest digit string (more complete reading), but only
    # if it has at least 2 votes.  A single outlier "11+" shouldn't beat
    # many "1+" results, but a genuine "1470x" seen by 2+ configs should
    # beat "470x".
    from itertools import groupby
    digit_counts = Counter(d for d, _ in results)
    by_len: dict[int, list[str]] = {}
    for d, cnt in digit_counts.items():
        by_len.setdefault(len(d), []).append(d)

    best_digits = ""
    for length in sorted(by_len, reverse=True):
        candidates = by_len[length]
        top = max(candidates, key=lambda d: digit_counts[d])
        if digit_counts[top] >= 2 or length == min(by_len):
            best_digits = top
            break
    if not best_digits:
        best_digits = digit_counts.most_common(1)[0][0]

    # Include operator if any result with this digit string has one
    ops = [op for d, op in results if d == best_digits and op]
    if ops:
        op_counts = Counter(ops)
        return best_digits + op_counts.most_common(1)[0][0]
    return best_digits


_OCR_CONFIGS = [
    "--oem 1 --psm 7 -c tessedit_char_whitelist=0123456789+x-/",
    "--oem 1 --psm 8 -c tessedit_char_whitelist=0123456789+x-/",
    "--oem 1 --psm 13 -c tessedit_char_whitelist=0123456789+x-/",
    "--psm 7 -c tessedit_char_whitelist=0123456789+x-/",
    "--psm 8 -c tessedit_char_whitelist=0123456789+x-/",
    "--psm 13 -c tessedit_char_whitelist=0123456789+x-/",
    "--psm 7",
    "--psm 8",
]


def _ocr_crop(crop_gray: np.ndarray) -> str:
    """OCR a small grayscale crop. Returns cleaned text."""
    h, w = crop_gray.shape
    if h < 10 or w < 10:
        return ""

    # Trim to text area (removes border line artifacts)
    trimmed = _trim_to_text(crop_gray)
    th, tw = trimmed.shape
    if th < 5 or tw < 5:
        return ""

    padded = _prepare_ocr_image(trimmed)
    result = _run_tesseract(padded, _OCR_CONFIGS)

    # Cage value "0" is never valid in Mathdoku (all values are positive).
    # The most common OCR confusion is "9" → "0" due to similar shapes.
    # Replace "0" with "9" as a post-processing correction.
    m = _LABEL_RE.match(result)
    if m and m.group(1) == "0":
        op_suffix = m.group(2) or ""
        _dbg(f"  Correcting invalid value '0' -> '9' (likely 9→0 OCR confusion)")
        result = "9" + op_suffix

    return result


def _read_cage_labels(
    gray: np.ndarray,
    gx: int, gy: int,
    n: int,
    h_pos: list[int], v_pos: list[int],
    cages: list[list[tuple[int, int]]],
) -> list[tuple[str, str | None]]:
    """For each cage, read and parse its label. Returns [(value, op), ...]."""
    _require_tesseract()
    results: list[tuple[str, str | None]] = []
    for idx, cells in enumerate(cages):
        tl_r, tl_c = cells[0]
        cx, cy = v_pos[tl_c], h_pos[tl_r]
        cw = v_pos[tl_c + 1] - cx
        ch = h_pos[tl_r + 1] - cy

        margin = 3
        lx = gx + cx + margin
        ly = gy + cy + margin
        lw = int(cw * 0.92)
        lh = int(ch * 0.42)

        crop = gray[ly:ly + lh, lx:lx + lw]
        if crop.size == 0:
            results.append(("?", None))
            continue

        _dbg_save(f"debug_label_{idx}.png", crop)

        raw = _ocr_crop(crop)
        m = _LABEL_RE.match(raw)
        if m:
            results.append((m.group(1), m.group(2)))
        elif raw and raw[0].isdigit():
            digits = re.match(r"\d+", raw)
            val = digits.group(0) if digits else raw
            rest = raw[len(val):]
            op = rest[0] if rest and rest[0] in "+-x/" else None
            results.append((val, op))
        else:
            results.append(("?", None))
    return results


# ── main pipeline ──────────────────────────────────────────────────────────

def ocr_mathdoku(
    img_path: Path,
    *,
    n: int | None = None,
    puzzle_id: str | None = None,
    difficulty: int | None = None,
) -> dict:
    img = cv2.imread(str(img_path))
    if img is None:
        raise FileNotFoundError(f"Cannot read: {img_path}")
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # 1. Grid bounding box
    gx, gy, gw, gh = _find_grid_bbox(gray)
    _dbg(f"Grid bbox: ({gx},{gy}) {gw}x{gh}")

    # 2. Candidate line positions
    h_cands, v_cands = _detect_line_positions(gray, gx, gy, gw, gh)

    # 3. Grid size & line selection
    n_detected, h_pos, v_pos = _detect_size_and_lines(h_cands, v_cands, gh, gw, n)
    if n is None:
        n = n_detected
    print(f"Grid: {n}x{n} at ({gx},{gy}) {gw}x{gh}")
    _dbg(f"h_pos={h_pos}")
    _dbg(f"v_pos={v_pos}")

    # 4. Border classification
    h_thick, v_thick = _classify_borders(gray, gx, gy, n, h_pos, v_pos)

    # 5. Cage grouping
    cages = _group_cages(n, h_thick, v_thick)
    print(f"Cages: {len(cages)}")

    # 6. Label OCR
    labels = _read_cage_labels(gray, gx, gy, n, h_pos, v_pos, cages)

    # 7. Build result
    has_ops = False
    cages_data: list[dict] = []
    for cells, (value, op) in zip(cages, labels):
        refs = _FlowList(_cell_a1(r, c) for r, c in cells)
        entry: dict = {"cells": refs}
        entry["value"] = int(value) if value.isdigit() else value
        if op:
            entry["op"] = op
            has_ops = True
        cages_data.append(entry)
        print(f"  {refs}: {value}{op or ''}")

    if puzzle_id is None:
        stem = img_path.stem
        m = re.search(r"\d+", stem)
        puzzle_id = m.group(0) if m else stem

    result: dict = {"id": puzzle_id, "size": n}
    if difficulty is not None:
        result["difficulty"] = difficulty
    result["operations"] = has_ops
    result["cages"] = cages_data

    try:
        result["id"] = int(result["id"])
    except (ValueError, TypeError):
        pass

    return result


def main() -> None:
    global _DEBUG
    ap = argparse.ArgumentParser(description="OCR a Mathdoku screenshot to YAML")
    ap.add_argument("image", help="puzzle screenshot (PNG, JPG, ...)")
    ap.add_argument("-o", "--output", help="output YAML (default: <stem>.yaml)")
    ap.add_argument("--size", type=int, choices=range(4, 10), metavar="N",
                    help="grid size 4-9 (auto-detected if omitted)")
    ap.add_argument("--id", help="puzzle ID (default: filename stem)")
    ap.add_argument("--difficulty", type=int, help="difficulty level")
    ap.add_argument("--debug", action="store_true", help="save debug images and info")
    args = ap.parse_args()

    _DEBUG = args.debug

    path = Path(args.image)
    if not path.is_file():
        print(f"Error: {path} not found", file=sys.stderr)
        raise SystemExit(1)

    out = Path(args.output) if args.output else path.with_suffix(".yaml")
    result = ocr_mathdoku(path, n=args.size, puzzle_id=args.id, difficulty=args.difficulty)

    with open(out, "w", encoding="utf-8") as f:
        yaml.dump(result, f, default_flow_style=False, allow_unicode=True, sort_keys=False)
    print(f"\nWrote {out}")


if __name__ == "__main__":
    main()
