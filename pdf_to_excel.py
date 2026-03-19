"""
PDF → Excel Converter
=====================
• Preserves text alignment (left / center / right) by analyzing word x-positions
• Reproduces tables with full formatting (headers, alternating rows, borders)
• Extracts and places images at their accurate PDF coordinates
• Each PDF page → one Excel sheet

Usage:
    python pdf_to_excel.py input.pdf output.xlsx
"""

import sys, os, io, re, subprocess
from collections import defaultdict

import pdfplumber
from pdf2image import convert_from_path
from PIL import Image as PILImage
from openpyxl import Workbook
from openpyxl.styles import (Font, PatternFill, Alignment,
                              Border, Side, GradientFill)
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter

# ── Constants ──────────────────────────────────────────────────────────────────
EXCEL_COLS   = 16          # number of columns to map page width into
ROW_HEIGHT   = 15          # default Excel row height (pts)
IMG_DIR      = "/tmp/pdf_conv_images"
os.makedirs(IMG_DIR, exist_ok=True)

# ── Colour palette ─────────────────────────────────────────────────────────────
C_HEADER_BG  = "1F3864"
C_HEADER_FG  = "FFFFFF"
C_TOTAL_BG   = "2F5597"
C_TOTAL_FG   = "FFFFFF"
C_ROW_EVEN   = "EBF3FB"
C_ROW_ODD    = "FFFFFF"
C_TITLE_FG   = "1F3864"
C_TITLE_BG   = "DCE6F1"
C_HEADING_FG = "2F5597"

def solid(hex_color):
    return PatternFill("solid", fgColor=hex_color)

THIN = Side(style="thin",   color="BBBBBB")
MED  = Side(style="medium", color="2F5597")
NO   = Side(style=None)

def cell_border(top=THIN, bottom=THIN, left=THIN, right=THIN):
    return Border(top=top, bottom=bottom, left=left, right=right)

# ── Alignment detection ────────────────────────────────────────────────────────
def detect_alignment(words, page_width, margin_frac=0.08):
    """
    Given a list of pdfplumber word dicts on ONE line, decide
    whether the line is LEFT / CENTER / RIGHT aligned.

    Strategy:
      - Compute the line's x-span [x_min .. x_max]
      - Compute how far the text midpoint is from the page midpoint
      - If midpoint is within ±10% of page centre → CENTER
      - If text starts past 55% of page width            → RIGHT
      - Otherwise                                         → LEFT
    """
    if not words:
        return "left"

    x_min   = min(float(w["x0"]) for w in words)
    x_max   = max(float(w["x1"]) for w in words)
    mid_txt = (x_min + x_max) / 2
    mid_pg  = page_width / 2
    margin  = page_width * margin_frac

    if abs(mid_txt - mid_pg) <= margin:
        return "center"
    if x_min > page_width * 0.55:
        return "right"
    return "left"

# ── Helpers ────────────────────────────────────────────────────────────────────
def in_any_table(y, table_bboxes, tol=2):
    for (x0, top, x1, bot) in table_bboxes:
        if top - tol <= y <= bot + tol:
            return True
    return False

def font_size_of_words(words):
    sizes = [float(w.get("size") or 10) for w in words if w.get("size")]
    return sum(sizes) / len(sizes) if sizes else 10

# ── Image extraction ───────────────────────────────────────────────────────────
def extract_images_from_pdf(pdf_path):
    """Extract all embedded images using pdfimages (poppler). Returns sorted file list."""
    subprocess.run(
        ["pdfimages", "-all", pdf_path, os.path.join(IMG_DIR, "img")],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
    )
    files = sorted([
        os.path.join(IMG_DIR, f) for f in os.listdir(IMG_DIR)
        if re.match(r"img-\d+", f)
    ])
    return files

# ── Main conversion ────────────────────────────────────────────────────────────
def convert(pdf_path, output_path):
    print(f"Converting: {pdf_path}")

    # Extract all embedded images once (across whole PDF)
    img_files = extract_images_from_pdf(pdf_path)
    print(f"  → {len(img_files)} embedded image(s) found")

    # Render page thumbnails for reference (not placed in Excel)
    page_renders = convert_from_path(pdf_path, dpi=150)

    wb = Workbook()
    global_img_idx = 0  # pointer into img_files across all pages

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):

            # ── Sheet setup ───────────────────────────────────────────────────
            ws = wb.active if page_num == 0 else wb.create_sheet()
            ws.title = f"Page {page_num + 1}"
            ws.sheet_view.showGridLines = True

            page_w = float(page.width)
            page_h = float(page.height)

            # Set uniform column widths
            col_w = 11
            for c in range(1, EXCEL_COLS + 2):
                ws.column_dimensions[get_column_letter(c)].width = col_w

            current_row = 1  # tracks next free Excel row

            # ── Collect table bounding boxes ──────────────────────────────────
            tbl_objects = page.find_tables()
            table_bboxes = [tuple(t.bbox) for t in tbl_objects]  # (x0,top,x1,bot)

            # ── Group non-table words into lines keyed by rounded y ───────────
            words = page.extract_words(
                keep_blank_chars=True, x_tolerance=3, y_tolerance=3
            )
            line_map = defaultdict(list)
            for w in words:
                y = float(w["top"])
                if not in_any_table(y, table_bboxes):
                    y_key = round(y / 4) * 4   # 4-pt bucket
                    line_map[y_key].append(w)

            # ── Build event list: (pdf_y, type, data) ─────────────────────────
            events = []
            for y_key, wds in line_map.items():
                events.append((y_key, "text", wds))
            for ti, tbl_obj in enumerate(tbl_objects):
                events.append((tbl_obj.bbox[1], "table", (ti, tbl_obj)))
            for ii, img_meta in enumerate(page.images):
                events.append((float(img_meta.get("top", 0)), "image", (ii, img_meta)))

            events.sort(key=lambda e: e[0])

            written_tables = set()
            written_images = set()

            for pdf_y, etype, edata in events:

                # ════════════════════════════════════════════════════════════
                # TEXT LINE
                # ════════════════════════════════════════════════════════════
                if etype == "text":
                    wds = sorted(edata, key=lambda w: float(w["x0"]))
                    line_text = " ".join(w["text"] for w in wds).strip()
                    if not line_text:
                        continue

                    # Detect alignment
                    alignment = detect_alignment(wds, page_w)

                    # Detect font size → classify as title / heading / body
                    avg_size = font_size_of_words(wds)
                    is_title   = avg_size >= 16
                    is_heading = 11 <= avg_size < 16

                    # Write merged cell across all columns
                    cell = ws.cell(row=current_row, column=1, value=line_text)
                    ws.merge_cells(
                        start_row=current_row, start_column=1,
                        end_row=current_row, end_column=EXCEL_COLS
                    )

                    # Alignment mapping
                    xl_align = {"left": "left", "center": "center", "right": "right"}[alignment]

                    if is_title:
                        cell.font      = Font(name="Arial", bold=True, size=16, color=C_TITLE_FG)
                        cell.fill      = solid(C_TITLE_BG)
                        cell.alignment = Alignment(horizontal=xl_align, vertical="center", wrap_text=True)
                        ws.row_dimensions[current_row].height = 32
                    elif is_heading:
                        cell.font      = Font(name="Arial", bold=True, size=12, color=C_HEADING_FG)
                        cell.alignment = Alignment(horizontal=xl_align, vertical="center", wrap_text=True)
                        ws.row_dimensions[current_row].height = 22
                    else:
                        cell.font      = Font(name="Arial", size=10, color="333333")
                        cell.alignment = Alignment(horizontal=xl_align, vertical="center", wrap_text=True)
                        ws.row_dimensions[current_row].height = ROW_HEIGHT

                    current_row += 1

                # ════════════════════════════════════════════════════════════
                # TABLE
                # ════════════════════════════════════════════════════════════
                elif etype == "table":
                    ti, tbl_obj = edata
                    if ti in written_tables:
                        continue
                    written_tables.add(ti)

                    tbl_data = tbl_obj.extract()
                    if not tbl_data:
                        continue

                    current_row += 1  # gap before table

                    num_cols = max(len(r) for r in tbl_data)

                    for r_i, row in enumerate(tbl_data):
                        is_header = (r_i == 0)
                        is_total  = (r_i == len(tbl_data) - 1)

                        for c_i in range(num_cols):
                            val = row[c_i] if c_i < len(row) else ""
                            cell = ws.cell(
                                row=current_row, column=c_i + 1,
                                value=str(val).strip() if val else ""
                            )

                            if is_header:
                                cell.font      = Font(name="Arial", bold=True, size=10, color=C_HEADER_FG)
                                cell.fill      = solid(C_HEADER_BG)
                                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                                cell.border    = cell_border(top=MED, bottom=MED, left=MED, right=MED)
                            elif is_total:
                                cell.font      = Font(name="Arial", bold=True, size=10, color=C_TOTAL_FG)
                                cell.fill      = solid(C_TOTAL_BG)
                                cell.alignment = Alignment(horizontal="center", vertical="center")
                                cell.border    = cell_border(top=MED, bottom=MED, left=MED, right=MED)
                            else:
                                bg = C_ROW_EVEN if r_i % 2 == 0 else C_ROW_ODD
                                cell.font      = Font(name="Arial", size=10, color="1A1A1A")
                                cell.fill      = solid(bg)
                                cell.alignment = Alignment(horizontal="center", vertical="center")
                                cell.border    = cell_border()

                        ws.row_dimensions[current_row].height = 20
                        current_row += 1

                    current_row += 1  # gap after table

                # ════════════════════════════════════════════════════════════
                # IMAGE
                # ════════════════════════════════════════════════════════════
                elif etype == "image":
                    ii, img_meta = edata
                    if ii in written_images:
                        continue
                    written_images.add(ii)

                    if global_img_idx >= len(img_files):
                        continue

                    img_path = img_files[global_img_idx]
                    global_img_idx += 1

                    if not os.path.exists(img_path):
                        continue

                    try:
                        pil_img = PILImage.open(img_path).convert("RGBA")

                        # Original PDF coords
                        pdf_top   = float(img_meta.get("top",    0))
                        pdf_left  = float(img_meta.get("x0",     0))
                        pdf_bot   = float(img_meta.get("bottom", page_h))
                        pdf_right = float(img_meta.get("x1",     page_w))

                        img_w_pts = pdf_right - pdf_left
                        img_h_pts = pdf_bot   - pdf_top

                        # Map PDF width/height (in pts) → pixels
                        # A4 ≈ 595 pts wide; render at 150 dpi → 595/72*150 ≈ 1240px
                        scale = 150 / 72  # dpi / pts_per_inch
                        target_px_w = int(img_w_pts * scale)
                        target_px_h = int(img_h_pts * scale)
                        target_px_w = max(target_px_w, 40)
                        target_px_h = max(target_px_h, 20)

                        pil_img = pil_img.resize(
                            (target_px_w, target_px_h), PILImage.LANCZOS
                        )

                        # Convert RGBA → RGB for JPEG-safe output
                        bg = PILImage.new("RGB", pil_img.size, (255, 255, 255))
                        bg.paste(pil_img, mask=pil_img.split()[3] if pil_img.mode == "RGBA" else None)
                        pil_img = bg

                        buf = io.BytesIO()
                        pil_img.save(buf, format="PNG")
                        buf.seek(0)

                        xl_img        = XLImage(buf)
                        xl_img.width  = target_px_w
                        xl_img.height = target_px_h

                        # Accurate anchor: map PDF (top, left) → Excel (row, col)
                        anchor_row = max(1, int((pdf_top  / page_h) * current_row) + 1)
                        anchor_col = max(1, int((pdf_left / page_w) * EXCEL_COLS)  + 1)

                        anchor_cell = f"{get_column_letter(anchor_col)}{anchor_row}"
                        ws.add_image(xl_img, anchor_cell)

                        # Ensure enough rows exist below anchor for the image
                        rows_needed = int(target_px_h / ROW_HEIGHT) + 1
                        end_row = anchor_row + rows_needed
                        while current_row < end_row:
                            ws.row_dimensions[current_row].height = ROW_HEIGHT
                            current_row += 1

                        print(f"  [Page {page_num+1}] Image {ii+1} → anchor {anchor_cell} "
                              f"({target_px_w}×{target_px_h}px, "
                              f"PDF pos top={pdf_top:.0f} left={pdf_left:.0f})")

                    except Exception as exc:
                        print(f"  [Page {page_num+1}] Image {ii+1} skipped: {exc}")

    wb.save(output_path)
    print(f"\n✓ Saved → {output_path}")


# ── Entry point ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python pdf_to_excel.py  input.pdf  output.xlsx")
        sys.exit(1)
    convert(sys.argv[1], sys.argv[2])
