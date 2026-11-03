# -*- coding: utf-8 -*-
"""
s1_framework_diagram.py
-----------------------
Academic Conceptual Framework Generator (Bagan Kerangka Berpikir)
Designed to replicate the exact visual layout and typography of S1/S2 Indonesian
University Thesis standards:
- Supports 1, 2, 3, 4, and 5 Independent Variables (X) -> Dependent Variable (Y)
- Outer bounding box on left enclosing independent variables (for k >= 2)
- Top-right error term box: "Faktor-faktor Lain (e)" with dashed arrow into Y
- Dependent variable box (Y) on the right
- 100% strictly orthogonal (90-degree) connector routing without slanted lines
- Simultaneous effect connector (Uji F) with loop-around (Image 1 style)
  or direct horizontal (Image 2 style)
- High-resolution anti-aliased rendering with Times New Roman Bold typography.
"""

import os
import math
from typing import List, Dict, Optional, Tuple, Any
from PIL import Image, ImageDraw, ImageFont

FONT_DIR = r"C:\Windows\Fonts"
DEFAULT_OUT_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "output", "figs"
)


def _load_font(font_name: str, size: int) -> ImageFont.FreeTypeFont:
    """Load TrueType font from Windows font directory with fallback."""
    candidates = [
        os.path.join(FONT_DIR, font_name),
        os.path.join(FONT_DIR, "timesbd.ttf"),
        os.path.join(FONT_DIR, "arialbd.ttf"),
    ]
    for p in candidates:
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, size)
            except Exception:
                pass
    return ImageFont.load_default()


def _draw_box(
    draw: ImageDraw.ImageDraw,
    rect: Tuple[int, int, int, int],
    text_lines: List[str],
    font: ImageFont.FreeTypeFont,
    line_width: int = 3,
    border_color: Tuple[int, int, int] = (0, 0, 0),
    fill_color: Tuple[int, int, int] = (255, 255, 255),
    line_spacing: int = 6,
):
    """Draw a sharp rectangular box with centered multi-line text."""
    x1, y1, x2, y2 = rect
    draw.rectangle([x1, y1, x2, y2], fill=fill_color, outline=border_color, width=line_width)

    line_heights = []
    line_widths = []
    for line in text_lines:
        bbox = draw.textbbox((0, 0), line, font=font)
        w = bbox[2] - bbox[0]
        h = bbox[3] - bbox[1]
        line_widths.append(w)
        line_heights.append(h)

    total_h = sum(line_heights) + line_spacing * (len(text_lines) - 1)
    start_y = y1 + (y2 - y1 - total_h) / 2.0

    curr_y = start_y
    for i, line in enumerate(text_lines):
        w = line_widths[i]
        x = x1 + (x2 - x1 - w) / 2.0
        draw.text((x, curr_y), line, font=font, fill=(0, 0, 0))
        curr_y += line_heights[i] + line_spacing


def _draw_arrowhead(
    draw: ImageDraw.ImageDraw,
    tip: Tuple[float, float],
    direction: str,
    size: float = 14.0,
    color: Tuple[int, int, int] = (0, 0, 0),
):
    """
    Draw a sharp solid black triangle arrowhead pointing exactly into tip.
    direction: 'RIGHT', 'LEFT', 'DOWN', 'UP'
    """
    tx, ty = tip
    if direction == "DOWN":
        p1 = (tx - size * 0.52, ty - size)
        p2 = (tx + size * 0.52, ty - size)
        draw.polygon([(tx, ty), p1, p2], fill=color)
    elif direction == "UP":
        p1 = (tx - size * 0.52, ty + size)
        p2 = (tx + size * 0.52, ty + size)
        draw.polygon([(tx, ty), p1, p2], fill=color)
    elif direction == "RIGHT":
        p1 = (tx - size, ty - size * 0.52)
        p2 = (tx - size, ty + size * 0.52)
        draw.polygon([(tx, ty), p1, p2], fill=color)
    elif direction == "LEFT":
        p1 = (tx + size, ty - size * 0.52)
        p2 = (tx + size, ty + size * 0.52)
        draw.polygon([(tx, ty), p1, p2], fill=color)


def _draw_path(
    draw: ImageDraw.ImageDraw,
    points: List[Tuple[float, float]],
    line_width: int = 3,
    color: Tuple[int, int, int] = (0, 0, 0),
    dashed: bool = False,
    dash_len: float = 14.0,
    gap_len: float = 8.0,
    arrow_dir: Optional[str] = None,
    arrow_size: float = 14.0,
):
    """Draw a strictly orthogonal polyline path with optional dashing and final arrowhead."""
    for i in range(len(points) - 1):
        p1 = points[i]
        p2 = points[i + 1]
        is_last_seg = (i == len(points) - 2)

        end_pt = p2
        if is_last_seg and arrow_dir:
            if arrow_dir == "DOWN":
                end_pt = (p2[0], p2[1] - arrow_size * 0.6)
            elif arrow_dir == "UP":
                end_pt = (p2[0], p2[1] + arrow_size * 0.6)
            elif arrow_dir == "RIGHT":
                end_pt = (p2[0] - arrow_size * 0.6, p2[1])
            elif arrow_dir == "LEFT":
                end_pt = (p2[0] + arrow_size * 0.6, p2[1])

        if not dashed:
            draw.line([p1, end_pt], fill=color, width=line_width)
        else:
            x1, y1 = p1
            x2, y2 = end_pt
            dx = x2 - x1
            dy = y2 - y1
            dist = math.hypot(dx, dy)
            if dist > 0:
                ux = dx / dist
                uy = dy / dist
                curr = 0.0
                while curr < dist:
                    seg_len = min(dash_len, dist - curr)
                    sa = (x1 + ux * curr, y1 + uy * curr)
                    sb = (x1 + ux * (curr + seg_len), y1 + uy * (curr + seg_len))
                    draw.line([sa, sb], fill=color, width=line_width)
                    curr += dash_len + gap_len

    if arrow_dir:
        _draw_arrowhead(draw, points[-1], arrow_dir, size=arrow_size, color=color)


def _draw_label(
    draw: ImageDraw.ImageDraw,
    text: str,
    pos: Tuple[float, float],
    font: ImageFont.FreeTypeFont,
    bg_pad: int = 4,
):
    """Draw label text placed cleanly above the horizontal path."""
    cx, cy = pos
    bbox = draw.textbbox((0, 0), text, font=font)
    w = bbox[2] - bbox[0]
    h = bbox[3] - bbox[1]
    x = cx - w / 2.0
    y = cy - h / 2.0
    if bg_pad > 0:
        draw.rectangle([x - bg_pad, y - bg_pad, x + w + bg_pad, y + h + bg_pad], fill=(255, 255, 255))
    draw.text((x, y), text, font=font, fill=(0, 0, 0))


def wrap_text_lines(text: str, max_chars_per_line: int = 18) -> List[str]:
    """Gracefully wrap variable name text into balanced lines."""
    words = text.strip().split()
    if not words:
        return []
    lines = []
    cur_line = []
    cur_len = 0
    for w in words:
        if cur_len + len(w) + (1 if cur_line else 0) <= max_chars_per_line:
            cur_line.append(w)
            cur_len += len(w) + (1 if cur_line else 0)
        else:
            if cur_line:
                lines.append(" ".join(cur_line))
            cur_line = [w]
            cur_len = len(w)
    if cur_line:
        lines.append(" ".join(cur_line))
    return lines


def generate_framework_diagram(
    x_variables: List[Dict[str, str]],
    y_variable: Dict[str, str],
    error_term_label: str = "Faktor-faktor\nLain (e)",
    label_style: str = "H",  # "H" (H1, H2..), "numeric" (1, 2..), or custom list
    custom_labels: Optional[List[str]] = None,
    simultaneous_style: Optional[str] = None,  # "loop", "straight", or "none"
    error_dashed: bool = True,
    supersample: int = 2,
    output_path: Optional[str] = None,
) -> str:
    """
    Generate publication-grade academic Kerangka Berpikir diagram matching Image 1 & 2 styles.

    Features:
    - 1 Variable: clean horizontal arrow from X to Y, with (e) above Y.
    - 2 Variables: X1 -> top Y, X2 -> bottom Y, (e) -> top Y, Simultan -> straight into left Y (Image 2 style).
    - 3 Variables: X1 -> top Y, X2 -> left Y, X3 -> bottom Y, Simultan -> bottom loop into right Y (Image 1 style).
    - 4 Variables: X1 -> top Y, X2 & X3 -> straight into left Y, X4 -> bottom Y, Simultan -> bottom loop into right Y.
    - 5 Variables: X1 & X2 -> top Y, X3 -> left Y, X4 & X5 -> bottom Y, Simultan -> bottom loop into right Y.
    - 100% strictly orthogonal (90-degree) connectors with zero line collisions and no box overlap.
    """
    k = len(x_variables)
    if k < 1 or k > 5:
        raise ValueError("Framework diagram supports between 1 and 5 independent variables.")

    if simultaneous_style is None:
        if k == 1:
            simultaneous_style = "none"
        elif k == 2:
            simultaneous_style = "straight"
        else:
            simultaneous_style = "loop"

    S = supersample

    # Determine Canvas Dimensions
    if k == 1:
        base_w, base_h = 1000, 440
    elif k == 2:
        base_w, base_h = 1080, 560
    elif k == 3:
        base_w, base_h = 1140, 650
    elif k == 4:
        base_w, base_h = 1220, 750
    else:  # k == 5
        base_w, base_h = 1320, 880

    W = base_w * S
    H = base_h * S

    img = Image.new("RGB", (W, H), (255, 255, 255))
    draw = ImageDraw.Draw(img)

    # Fonts
    font_bold_lg = _load_font("timesbd.ttf", 25 * S)
    font_bold_md = _load_font("timesbd.ttf", 22 * S)
    font_bold_sm = _load_font("timesbd.ttf", 19 * S)
    font_label = _load_font("timesbd.ttf", 22 * S)

    line_w = 3 * S
    box_w = 3 * S
    arrow_sz = 15 * S

    # Hypotheses / Path Labels
    if custom_labels:
        labels = {"X": custom_labels[:k], "sim": custom_labels[k] if len(custom_labels) > k else ""}
    elif label_style == "numeric":
        if k == 1:
            labels = {"X": ["1"], "sim": ""}
        else:
            # Image 1 style: Simultan is 1, X1 is 2, X2 is 3, etc.
            labels = {"X": [str(i + 2) for i in range(k)], "sim": "1"}
    else:  # "H"
        if k == 1:
            labels = {"X": ["H1"], "sim": ""}
        else:
            labels = {"X": [f"H{i+1}" for i in range(k)], "sim": f"H{k+1}"}

    x_labels = labels["X"]
    sim_label = labels["sim"]

    # 1. Left Side: Outer Box & X Boxes
    outer_x1 = 65 * S
    outer_w = 360 * S
    outer_x2 = outer_x1 + outer_w

    inner_pad_x = 24 * S
    inner_pad_y = 22 * S
    x_box_x1 = outer_x1 + inner_pad_x
    x_box_x2 = outer_x2 - inner_pad_x

    # Vertical coordinates for X boxes
    if k == 1:
        x_box_w = 320 * S
        x_box_h = 110 * S
        x_box_x1 = 100 * S
        x_box_x2 = x_box_x1 + x_box_w
        x_mid_y = 230 * S
        x_box_rects = [(x_box_x1, x_mid_y - x_box_h / 2.0, x_box_x2, x_mid_y + x_box_h / 2.0)]
    elif k == 2:
        x_box_h = 100 * S
        x_box_gap = 120 * S
        outer_y1 = 85 * S
        outer_y2 = outer_y1 + 2 * inner_pad_y + 2 * x_box_h + x_box_gap
        x_box_rects = [
            (x_box_x1, outer_y1 + inner_pad_y, x_box_x2, outer_y1 + inner_pad_y + x_box_h),
            (x_box_x1, outer_y1 + inner_pad_y + x_box_h + x_box_gap, x_box_x2, outer_y2 - inner_pad_y)
        ]
    elif k == 3:
        x_box_h = 92 * S
        x_box_gap = 48 * S
        outer_y1 = 80 * S
        outer_y2 = outer_y1 + 2 * inner_pad_y + 3 * x_box_h + 2 * x_box_gap
        x_box_rects = []
        curr_y = outer_y1 + inner_pad_y
        for i in range(3):
            x_box_rects.append((x_box_x1, curr_y, x_box_x2, curr_y + x_box_h))
            curr_y += x_box_h + x_box_gap
    elif k == 4:
        x_box_h = 86 * S
        x_box_gap = 35 * S
        outer_y1 = 70 * S
        outer_y2 = outer_y1 + 2 * inner_pad_y + 4 * x_box_h + 3 * x_box_gap
        x_box_rects = []
        curr_y = outer_y1 + inner_pad_y
        for i in range(4):
            x_box_rects.append((x_box_x1, curr_y, x_box_x2, curr_y + x_box_h))
            curr_y += x_box_h + x_box_gap
    else:  # k == 5
        x_box_h = 80 * S
        x_box_gap = 26 * S
        outer_y1 = 60 * S
        outer_y2 = outer_y1 + 2 * inner_pad_y + 5 * x_box_h + 4 * x_box_gap
        x_box_rects = []
        curr_y = outer_y1 + inner_pad_y
        for i in range(5):
            x_box_rects.append((x_box_x1, curr_y, x_box_x2, curr_y + x_box_h))
            curr_y += x_box_h + x_box_gap

    # 2. Right Side: Y Box and Error Term Box (e)
    # Positioning Y box
    y_box_w = 360 * S if k == 5 else 330 * S
    y_box_x1 = (base_w - (430 if k == 5 else 410)) * S
    y_box_x2 = y_box_x1 + y_box_w

    if k == 1:
        y_box_h = 110 * S
        y_box_y1 = x_mid_y - y_box_h / 2.0
        y_box_y2 = y_box_y1 + y_box_h
    elif k == 2:
        y_box_h = 135 * S
        mid_gap = (x_box_rects[0][3] + x_box_rects[1][1]) / 2.0
        y_box_y1 = mid_gap - y_box_h / 2.0
        y_box_y2 = y_box_y1 + y_box_h
    elif k == 3:
        y_box_h = 145 * S
        x2_mid_y = (x_box_rects[1][1] + x_box_rects[1][3]) / 2.0
        y_box_y1 = x2_mid_y - y_box_h / 2.0
        y_box_y2 = y_box_y1 + y_box_h
    elif k == 4:
        # Align Y so that X2 and X3 go straight horizontally into Y's left edge!
        x2_mid_y = (x_box_rects[1][1] + x_box_rects[1][3]) / 2.0
        x3_mid_y = (x_box_rects[2][1] + x_box_rects[2][3]) / 2.0
        y_box_y1 = x2_mid_y - 30 * S
        y_box_y2 = x3_mid_y + 30 * S
        y_box_h = y_box_y2 - y_box_y1
    else:  # k == 5
        y_box_h = 175 * S
        x3_mid_y = (x_box_rects[2][1] + x_box_rects[2][3]) / 2.0
        y_box_y1 = x3_mid_y - y_box_h / 2.0
        y_box_y2 = y_box_y1 + y_box_h

    # Error box (e) placed directly above Y box
    e_box_w = 230 * S
    e_box_h = 75 * S
    # Position (e) box shifted towards the right side of Y
    e_box_x2 = y_box_x2
    e_box_x1 = e_box_x2 - e_box_w
    e_box_y2 = y_box_y1 - 65 * S
    e_box_y1 = e_box_y2 - e_box_h

    # DRAW BOXES
    # 1. Outer Box on Left (for k >= 2)
    if k >= 2:
        draw.rectangle([outer_x1, outer_y1, outer_x2, outer_y2], fill=(255, 255, 255), outline=(0, 0, 0), width=box_w)

    # 2. X Boxes
    for i, xv in enumerate(x_variables):
        r = x_box_rects[i]
        name = xv.get("name", f"Variabel X{i+1}")
        code = xv.get("code", f"X{i+1}" if k > 1 else "X")
        name_lines = wrap_text_lines(name, max_chars_per_line=18)
        lines = name_lines + [f"({code})"]
        fnt = font_bold_md if len(name) <= 20 else font_bold_sm
        _draw_box(draw, r, lines, fnt, line_width=box_w)

    # 3. Y Box
    y_name = y_variable.get("name", "Kinerja Karyawan")
    y_code = y_variable.get("code", "Y")
    y_name_lines = wrap_text_lines(y_name, max_chars_per_line=18)
    y_lines = y_name_lines + [f"({y_code})"]
    fnt_y = font_bold_lg if len(y_name) <= 18 else font_bold_md
    _draw_box(draw, (int(y_box_x1), int(y_box_y1), int(y_box_x2), int(y_box_y2)), y_lines, fnt_y, line_width=box_w)

    # 4. Error Term Box (e)
    e_lines = [line.strip() for line in error_term_label.split("\n") if line.strip()]
    _draw_box(draw, (int(e_box_x1), int(e_box_y1), int(e_box_x2), int(e_box_y2)), e_lines, font_bold_md, line_width=box_w)

    # DRAW CONNECTORS
    # 1. Error Term (e) -> Top of Y
    e_arrow_x = (e_box_x1 + e_box_x2) / 2.0
    _draw_path(
        draw,
        [(e_arrow_x, e_box_y2), (e_arrow_x, y_box_y1)],
        line_width=line_w,
        dashed=error_dashed,
        arrow_dir="DOWN",
        arrow_size=arrow_sz,
    )

    # 2. Independent Variables -> Y
    # Top entry slot is placed safely to the left of the (e) box!
    top_entry_x = y_box_x1 + (e_box_x1 - y_box_x1) * 0.55
    bot_entry_x = top_entry_x

    if k == 1:
        # Perfectly horizontal straight arrow
        x_center_y = (x_box_rects[0][1] + x_box_rects[0][3]) / 2.0
        y_center_y = (y_box_y1 + y_box_y2) / 2.0
        _draw_path(
            draw,
            [(x_box_rects[0][2], x_center_y), (y_box_x1, y_center_y)],
            line_width=line_w,
            arrow_dir="RIGHT",
            arrow_size=arrow_sz,
        )
        mid_x = (x_box_rects[0][2] + y_box_x1) / 2.0
        _draw_label(draw, x_labels[0], (mid_x, x_center_y - 22 * S), font_label)

    elif k == 2:
        # X1 -> Top of Y (90 degree turn down)
        x1_cy = (x_box_rects[0][1] + x_box_rects[0][3]) / 2.0
        _draw_path(
            draw,
            [(x_box_rects[0][2], x1_cy), (top_entry_x, x1_cy), (top_entry_x, y_box_y1)],
            line_width=line_w,
            arrow_dir="DOWN",
            arrow_size=arrow_sz,
        )
        mid_x1 = (x_box_rects[0][2] + top_entry_x) / 2.0
        _draw_label(draw, x_labels[0], (mid_x1, x1_cy - 20 * S), font_label)

        # X2 -> Bottom of Y (90 degree turn up)
        x2_cy = (x_box_rects[1][1] + x_box_rects[1][3]) / 2.0
        _draw_path(
            draw,
            [(x_box_rects[1][2], x2_cy), (bot_entry_x, x2_cy), (bot_entry_x, y_box_y2)],
            line_width=line_w,
            arrow_dir="UP",
            arrow_size=arrow_sz,
        )
        mid_x2 = (x_box_rects[1][2] + bot_entry_x) / 2.0
        _draw_label(draw, x_labels[1], (mid_x2, x2_cy - 20 * S), font_label)

        # Simultaneous Effect
        if simultaneous_style == "straight":
            sim_y = (y_box_y1 + y_box_y2) / 2.0
            _draw_path(
                draw,
                [(outer_x2, sim_y), (y_box_x1, sim_y)],
                line_width=line_w,
                arrow_dir="RIGHT",
                arrow_size=arrow_sz,
            )
            mid_sim = (outer_x2 + y_box_x1) / 2.0
            _draw_label(draw, sim_label, (mid_sim, sim_y - 22 * S), font_label)
        else:
            sim_start_x = outer_x1 + outer_w * 0.45
            loop_y = outer_y2 + 45 * S
            loop_right_x = y_box_x2 + 45 * S
            sim_target_y = (y_box_y1 + y_box_y2) / 2.0
            _draw_path(
                draw,
                [
                    (sim_start_x, outer_y2),
                    (sim_start_x, loop_y),
                    (loop_right_x, loop_y),
                    (loop_right_x, sim_target_y),
                    (y_box_x2, sim_target_y),
                ],
                line_width=line_w,
                arrow_dir="LEFT",
                arrow_size=arrow_sz,
            )
            mid_loop_x = (sim_start_x + loop_right_x) / 2.0
            _draw_label(draw, sim_label, (mid_loop_x, loop_y - 20 * S), font_label)

    elif k == 3:
        # X1 -> Top of Y
        x1_cy = (x_box_rects[0][1] + x_box_rects[0][3]) / 2.0
        _draw_path(
            draw,
            [(x_box_rects[0][2], x1_cy), (top_entry_x, x1_cy), (top_entry_x, y_box_y1)],
            line_width=line_w,
            arrow_dir="DOWN",
            arrow_size=arrow_sz,
        )
        mid_x1 = (x_box_rects[0][2] + top_entry_x) / 2.0
        _draw_label(draw, x_labels[0], (mid_x1, x1_cy - 20 * S), font_label)

        # X2 -> Left center of Y (Straight horizontal)
        x2_cy = (x_box_rects[1][1] + x_box_rects[1][3]) / 2.0
        y_center_y = (y_box_y1 + y_box_y2) / 2.0
        _draw_path(
            draw,
            [(x_box_rects[1][2], x2_cy), (y_box_x1, y_center_y)],
            line_width=line_w,
            arrow_dir="RIGHT",
            arrow_size=arrow_sz,
        )
        mid_x2 = (x_box_rects[1][2] + y_box_x1) / 2.0
        _draw_label(draw, x_labels[1], (mid_x2, x2_cy - 20 * S), font_label)

        # X3 -> Bottom of Y
        x3_cy = (x_box_rects[2][1] + x_box_rects[2][3]) / 2.0
        _draw_path(
            draw,
            [(x_box_rects[2][2], x3_cy), (bot_entry_x, x3_cy), (bot_entry_x, y_box_y2)],
            line_width=line_w,
            arrow_dir="UP",
            arrow_size=arrow_sz,
        )
        mid_x3 = (x_box_rects[2][2] + bot_entry_x) / 2.0
        _draw_label(draw, x_labels[2], (mid_x3, x3_cy - 20 * S), font_label)

        # Simultaneous: Loop around bottom (Image 1 style)
        sim_start_x = outer_x1 + outer_w * 0.45
        loop_y = outer_y2 + 45 * S
        loop_right_x = y_box_x2 + 45 * S
        sim_target_y = (y_box_y1 + y_box_y2) / 2.0
        _draw_path(
            draw,
            [
                (sim_start_x, outer_y2),
                (sim_start_x, loop_y),
                (loop_right_x, loop_y),
                (loop_right_x, sim_target_y),
                (y_box_x2, sim_target_y),
            ],
            line_width=line_w,
            arrow_dir="LEFT",
            arrow_size=arrow_sz,
        )
        mid_loop_x = (sim_start_x + loop_right_x) / 2.0
        _draw_label(draw, sim_label, (mid_loop_x, loop_y - 20 * S), font_label)

    elif k == 4:
        # X1 -> Top of Y (turns down safely to the left of (e) box)
        x1_cy = (x_box_rects[0][1] + x_box_rects[0][3]) / 2.0
        _draw_path(
            draw,
            [(x_box_rects[0][2], x1_cy), (top_entry_x, x1_cy), (top_entry_x, y_box_y1)],
            line_width=line_w,
            arrow_dir="DOWN",
            arrow_size=arrow_sz,
        )
        mid_x1 = (x_box_rects[0][2] + top_entry_x) / 2.0
        _draw_label(draw, x_labels[0], (mid_x1, x1_cy - 20 * S), font_label)

        # X2 & X3 -> 100% Straight horizontal into left of Y!
        x2_cy = (x_box_rects[1][1] + x_box_rects[1][3]) / 2.0
        _draw_path(
            draw,
            [(x_box_rects[1][2], x2_cy), (y_box_x1, x2_cy)],
            line_width=line_w,
            arrow_dir="RIGHT",
            arrow_size=arrow_sz,
        )
        mid_x2 = (x_box_rects[1][2] + y_box_x1) / 2.0
        _draw_label(draw, x_labels[1], (mid_x2, x2_cy - 20 * S), font_label)

        x3_cy = (x_box_rects[2][1] + x_box_rects[2][3]) / 2.0
        _draw_path(
            draw,
            [(x_box_rects[2][2], x3_cy), (y_box_x1, x3_cy)],
            line_width=line_w,
            arrow_dir="RIGHT",
            arrow_size=arrow_sz,
        )
        mid_x3 = (x_box_rects[2][2] + y_box_x1) / 2.0
        _draw_label(draw, x_labels[2], (mid_x3, x3_cy - 20 * S), font_label)

        # X4 -> Bottom of Y
        x4_cy = (x_box_rects[3][1] + x_box_rects[3][3]) / 2.0
        _draw_path(
            draw,
            [(x_box_rects[3][2], x4_cy), (bot_entry_x, x4_cy), (bot_entry_x, y_box_y2)],
            line_width=line_w,
            arrow_dir="UP",
            arrow_size=arrow_sz,
        )
        mid_x4 = (x_box_rects[3][2] + bot_entry_x) / 2.0
        _draw_label(draw, x_labels[3], (mid_x4, x4_cy - 20 * S), font_label)

        # Simultaneous: Loop around bottom (Image 1 style)
        sim_start_x = outer_x1 + outer_w * 0.45
        loop_y = outer_y2 + 45 * S
        loop_right_x = y_box_x2 + 45 * S
        sim_target_y = (y_box_y1 + y_box_y2) / 2.0
        _draw_path(
            draw,
            [
                (sim_start_x, outer_y2),
                (sim_start_x, loop_y),
                (loop_right_x, loop_y),
                (loop_right_x, sim_target_y),
                (y_box_x2, sim_target_y),
            ],
            line_width=line_w,
            arrow_dir="LEFT",
            arrow_size=arrow_sz,
        )
        mid_loop_x = (sim_start_x + loop_right_x) / 2.0
        _draw_label(draw, sim_label, (mid_loop_x, loop_y - 20 * S), font_label)

    elif k == 5:
        # Non-crossing orthogonal distribution:
        # Both top entry points are safely to the left of the (e) box with ample spacing!
        top_entry_x1 = y_box_x1 + 48 * S
        top_entry_x2 = y_box_x1 + 118 * S
        bot_entry_x1 = top_entry_x1
        bot_entry_x2 = top_entry_x2

        # X1 -> Top of Y at entry 1
        x1_cy = (x_box_rects[0][1] + x_box_rects[0][3]) / 2.0
        _draw_path(
            draw,
            [(x_box_rects[0][2], x1_cy), (top_entry_x1, x1_cy), (top_entry_x1, y_box_y1)],
            line_width=line_w,
            arrow_dir="DOWN",
            arrow_size=arrow_sz,
        )
        mid_x1 = (x_box_rects[0][2] + top_entry_x1) / 2.0
        _draw_label(draw, x_labels[0], (mid_x1, x1_cy - 20 * S), font_label)

        # X2 -> Top of Y at entry 2
        x2_cy = (x_box_rects[1][1] + x_box_rects[1][3]) / 2.0
        _draw_path(
            draw,
            [(x_box_rects[1][2], x2_cy), (top_entry_x2, x2_cy), (top_entry_x2, y_box_y1)],
            line_width=line_w,
            arrow_dir="DOWN",
            arrow_size=arrow_sz,
        )
        mid_x2 = (x_box_rects[1][2] + top_entry_x2) / 2.0
        _draw_label(draw, x_labels[1], (mid_x2, x2_cy - 20 * S), font_label)

        # X3 -> Dead straight horizontal into center-left of Y!
        x3_cy = (x_box_rects[2][1] + x_box_rects[2][3]) / 2.0
        y_center_y = (y_box_y1 + y_box_y2) / 2.0
        _draw_path(
            draw,
            [(x_box_rects[2][2], x3_cy), (y_box_x1, y_center_y)],
            line_width=line_w,
            arrow_dir="RIGHT",
            arrow_size=arrow_sz,
        )
        mid_x3 = (x_box_rects[2][2] + y_box_x1) / 2.0
        _draw_label(draw, x_labels[2], (mid_x3, x3_cy - 20 * S), font_label)

        # X4 -> Bottom of Y at entry 2
        x4_cy = (x_box_rects[3][1] + x_box_rects[3][3]) / 2.0
        _draw_path(
            draw,
            [(x_box_rects[3][2], x4_cy), (bot_entry_x2, x4_cy), (bot_entry_x2, y_box_y2)],
            line_width=line_w,
            arrow_dir="UP",
            arrow_size=arrow_sz,
        )
        mid_x4 = (x_box_rects[3][2] + bot_entry_x2) / 2.0
        _draw_label(draw, x_labels[3], (mid_x4, x4_cy - 20 * S), font_label)

        # X5 -> Bottom of Y at entry 1
        x5_cy = (x_box_rects[4][1] + x_box_rects[4][3]) / 2.0
        _draw_path(
            draw,
            [(x_box_rects[4][2], x5_cy), (bot_entry_x1, x5_cy), (bot_entry_x1, y_box_y2)],
            line_width=line_w,
            arrow_dir="UP",
            arrow_size=arrow_sz,
        )
        mid_x5 = (x_box_rects[4][2] + bot_entry_x1) / 2.0
        _draw_label(draw, x_labels[4], (mid_x5, x5_cy - 20 * S), font_label)

        # Simultaneous: Loop around bottom (Image 1 style)
        sim_start_x = outer_x1 + outer_w * 0.45
        loop_y = outer_y2 + 45 * S
        loop_right_x = y_box_x2 + 45 * S
        sim_target_y = (y_box_y1 + y_box_y2) / 2.0
        _draw_path(
            draw,
            [
                (sim_start_x, outer_y2),
                (sim_start_x, loop_y),
                (loop_right_x, loop_y),
                (loop_right_x, sim_target_y),
                (y_box_x2, sim_target_y),
            ],
            line_width=line_w,
            arrow_dir="LEFT",
            arrow_size=arrow_sz,
        )
        mid_loop_x = (sim_start_x + loop_right_x) / 2.0
        _draw_label(draw, sim_label, (mid_loop_x, loop_y - 20 * S), font_label)

    # Output handling
    if not output_path:
        os.makedirs(DEFAULT_OUT_DIR, exist_ok=True)
        tag = f"kerangka_{k}var_{label_style}"
        output_path = os.path.join(DEFAULT_OUT_DIR, f"{tag}.png")
    else:
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)

    img.save(output_path, "PNG", dpi=(300, 300))
    return output_path
