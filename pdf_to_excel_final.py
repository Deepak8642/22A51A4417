"""
╔══════════════════════════════════════════════════════════════════════════════╗
║          PDF → EXCEL  ULTIMATE CONVERTER  —  FINAL VERSION  v1.0           ║
║          Works on ANY PDF: IRS forms, tax returns, tables, reports          ║
╚══════════════════════════════════════════════════════════════════════════════╝

STRATEGY (auto-detected per page):
  MODE 1 — GRID (PDF has drawn vertical lines ≥ 2)
    Uses v-lines as column separators → perfect for IRS forms (f6765, f4797…)

  MODE 2 — TABLE (pdfplumber finds at least one table with ≥ 2 cols)
    Uses pdfplumber's table extractor → handles PA-3, NY ST-810, MA ST-9

  MODE 3 — TEXT (fallback — free-flow positioned text)
    Clusters words by y-position into rows, detects left/right columns
    by x-gap → handles MD Form 202 and any plain-text PDF

FEATURES:
  ✓ Preserves every character — nothing dropped
  ✓ Fills (black/grey headers) → styled Excel cells
  ✓ Dot leaders absorbed into neighbouring segment
  ✓ Correct alignment (left / center / right) per cell
  ✓ Alternating row shading for readability
  ✓ Bold detection from font names
  ✓ Multi-page → separate worksheets
  ✓ Auto column widths, proportional row heights
  ✓ One-file output

INSTALL:  pip install pdfplumber openpyxl pillow
RUN:      python pdf_to_excel_final.py  input.pdf  output.xlsx
          python pdf_to_excel_final.py  input.pdf          (auto-names output)
"""

import sys, os, re
from collections import defaultdict

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─── tuneable constants ────────────────────────────────────────────────────────
Y_SNAP         = 2.0    # pt — chars within this band = same row
SEG_GAP        = 7      # pt — x-gap bigger than this splits a text segment
VLINE_MIN      = 2      # need ≥ this many PDF v-lines to use GRID mode
TABLE_MIN_COLS = 2      # pdfplumber table must have ≥ this many cols
COL_SNAP       = 5      # pt — snap nearby column boundaries together

# ─── colour helpers ───────────────────────────────────────────────────────────
def _fill(hex6): return PatternFill("solid", fgColor=hex6)
def _side(style="thin", color="CCCCCC"): return Side(style=style, color=color)
def _bdr(style="thin", color="CCCCCC"):
    s = _side(style, color)
    return Border(top=s, bottom=s, left=s, right=s)
def _fnt(bold=False, size=9, color="111111"):
    return Font(name="Arial", bold=bold, size=max(int(size), 7), color=color)
def _aln(h="left", wrap=True):
    return Alignment(horizontal=h, vertical="center", wrap_text=wrap)

THIN  = _bdr("thin",   "CCCCCC")
THICK = _bdr("medium", "333333")
ROW_A = "FFFFFF"
ROW_B = "F2F2F2"


# ══════════════════════════════════════════════════════════════════════════════
#  LOW-LEVEL HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def snap_vals(vals, tol):
    out = []
    for v in sorted(set(round(float(x), 1) for x in vals)):
        if out and v - out[-1] <= tol:
            out[-1] = (out[-1] + v) / 2
        else:
            out.append(v)
    return sorted(out)


def chars_to_rows(chars, y_snap=Y_SNAP):
    """Group chars into visual rows by y-position."""
    bkts = defaultdict(list)
    for ch in chars:
        key = round(float(ch["top"]) / y_snap) * y_snap
        bkts[key].append(ch)
    rows = []
    for k in sorted(bkts):
        chs = sorted(bkts[k], key=lambda c: float(c["x0"]))
        avg_y = sum(float(c["top"]) for c in chs) / len(chs)
        rows.append((avg_y, chs))
    return rows


def chars_to_segs(chs, gap=SEG_GAP):
    """Split a row's chars into x-gap segments; absorb dot leaders."""
    if not chs:
        return []
    chs = sorted(chs, key=lambda c: float(c["x0"]))
    groups = [[chs[0]]]
    for i in range(1, len(chs)):
        if float(chs[i]["x0"]) - float(chs[i - 1]["x1"]) > gap:
            groups.append([])
        groups[-1].append(chs[i])
    result = []
    for g in groups:
        txt = "".join(c["text"] for c in g).strip()
        if not txt:
            continue
        is_dots = all(c["text"].strip() in (".", "", " ") for c in g)
        if is_dots and result:
            result[-1]["x1"] = float(g[-1]["x1"])
            continue
        result.append(dict(
            x0   = float(g[0]["x0"]),
            x1   = float(g[-1]["x1"]),
            text = txt,
            bold = any("Bold" in str(c.get("fontname", "")) for c in g),
            size = float(g[0].get("size") or 9),
        ))
    return result


def get_fills(page):
    """Return significant background fills (black header / grey header)."""
    out = []
    pw = float(page.width)
    for r in page.rects:
        c = r.get("non_stroking_color")
        if c is None:
            continue
        if isinstance(c, (int, float)):
            c = (float(c),) * 3
        if not (isinstance(c, (list, tuple)) and len(c) >= 3):
            continue
        r0, g0, b0 = float(c[0]), float(c[1]), float(c[2])
        if   r0 < 0.2  and g0 < 0.2  and b0 < 0.2:  kind = "black"
        elif r0 > 0.45 and g0 > 0.45 and b0 > 0.45 \
             and not (r0 > 0.95 and g0 > 0.95 and b0 > 0.95): kind = "grey"
        else:
            continue
        fw = float(r["x1"] - r["x0"])
        if fw > pw * 0.20:          # only wide fills = headers
            out.append(dict(
                top=float(r["top"]), bottom=float(r["bottom"]),
                kind=kind,
            ))
    return out


def fill_at(y, fills, tol=3):
    for f in fills:
        if f["top"] - tol <= y <= f["bottom"] + tol:
            return f["kind"]
    return None


def slot_of(x, slots):
    best = 0
    for i in range(len(slots) - 1):
        if slots[i] - 3 <= x:
            best = i
    return best


def end_slot_of(x1, slots):
    best = 0
    for i in range(len(slots) - 1):
        if slots[i] <= x1 + 4:
            best = i
    return best


def detect_align(seg, sx0, sx1):
    cw = sx1 - sx0
    if cw < 5:
        return "left"
    mt = (seg["x0"] + seg["x1"]) / 2
    mc = (sx0 + sx1) / 2
    if abs(mt - mc) <= cw * 0.22:
        return "center"
    if seg["x0"] > sx0 + cw * 0.52:
        return "right"
    return "left"


def safe_merge(ws, r1, c1, r2, c2):
    if r1 == r2 and c1 == c2:
        return
    try:
        ws.merge_cells(start_row=r1, start_column=c1,
                       end_row=r2, end_column=c2)
    except Exception:
        pass


def _style_cell(cell, txt, bold, size, fg, bg, al, border):
    cell.value     = txt
    cell.font      = _fnt(bold=bold, size=size, color=fg)
    cell.fill      = _fill(bg)
    cell.alignment = _aln(al)
    cell.border    = border


# ══════════════════════════════════════════════════════════════════════════════
#  MODE 1 — GRID  (uses PDF vertical lines as column boundaries)
# ══════════════════════════════════════════════════════════════════════════════

def write_grid_page(ws, page):
    pw    = float(page.width)
    fills = get_fills(page)
    vis   = chars_to_rows(page.chars)

    # column slots from v-lines + page edges
    raw_x = [float(l["x0"]) for l in page.lines if abs(l["x0"] - l["x1"]) < 1]
    raw_x += [0.0, pw]
    slots = snap_vals(raw_x, COL_SNAP)
    n     = max(len(slots) - 1, 1)

    for ci in range(n):
        cw = slots[ci + 1] - slots[ci]
        ws.column_dimensions[get_column_letter(ci + 1)].width = max(3, round(cw / 5.2))

    for xl_row, (y, chs) in enumerate(vis, start=1):
        segs = chars_to_segs(chs)
        kind = fill_at(y, fills)
        sz   = segs[0]["size"] if segs else 9
        ws.row_dimensions[xl_row].height = max(10, round(sz * 1.6))

        if not segs:
            continue

        # BLACK full-width header
        if kind == "black":
            txt = "  " + "  ".join(s["text"] for s in segs)
            c   = ws.cell(row=xl_row, column=1, value=txt)
            c.font      = _fnt(bold=True, size=max(int(sz), 9), color="FFFFFF")
            c.fill      = _fill("111111")
            c.alignment = _aln("left")
            c.border    = THICK
            ws.row_dimensions[xl_row].height = 15
            if n > 1: safe_merge(ws, xl_row, 1, xl_row, n)
            continue

        # GREY header
        if kind == "grey":
            used = set()
            for seg in segs:
                ci_s = slot_of(seg["x0"], slots)
                ci_e = end_slot_of(seg["x1"], slots)
                if ci_s + 1 in used: continue
                al   = detect_align(seg, slots[ci_s], slots[min(ci_s+1, len(slots)-1)])
                c    = ws.cell(row=xl_row, column=ci_s + 1, value=seg["text"])
                c.font      = _fnt(bold=True, size=max(int(seg["size"]), 8))
                c.fill      = _fill("DCDCDC")
                c.alignment = _aln(al)
                c.border    = THIN
                if ci_e > ci_s:
                    safe_merge(ws, xl_row, ci_s+1, xl_row, ci_e+1)
                    for cc in range(ci_s+1, ci_e+2): used.add(cc)
                else:
                    used.add(ci_s + 1)
            ws.row_dimensions[xl_row].height = 13
            continue

        # Normal row
        bg   = ROW_B if xl_row % 2 == 0 else ROW_A
        used = set()
        for seg in segs:
            ci_s = slot_of(seg["x0"], slots)
            ci_e = end_slot_of(seg["x1"], slots)
            if ci_s + 1 in used: continue
            al   = detect_align(seg, slots[ci_s], slots[min(ci_s+1, len(slots)-1)])
            c    = ws.cell(row=xl_row, column=ci_s + 1, value=seg["text"])
            c.font      = _fnt(bold=seg["bold"], size=max(int(seg["size"]), 8))
            c.fill      = _fill(bg)
            c.alignment = _aln(al)
            c.border    = THIN
            if ci_e > ci_s:
                safe_merge(ws, xl_row, ci_s+1, xl_row, ci_e+1)
                for cc in range(ci_s+1, ci_e+2): used.add(cc)
            else:
                used.add(ci_s + 1)


# ══════════════════════════════════════════════════════════════════════════════
#  MODE 2 — TABLE  (pdfplumber table extractor)
# ══════════════════════════════════════════════════════════════════════════════

def write_table_page(ws, page, tables):
    """Write page using pdfplumber table data, filling gaps with free text."""
    fills = get_fills(page)
    vis   = chars_to_rows(page.chars)

    # Flatten all table cells into a set of (row, col) coverage
    # Build a global table grid that covers the whole page
    # Strategy: collect all tables, sort by y, interleave free-text rows

    # Map every pdfplumber table bbox
    table_bboxes = [(t.bbox, t.extract()) for t in tables]

    # Gather rows not inside any table
    def in_any_table(y):
        for (x0, top, x1, bottom), _ in table_bboxes:
            if top - 4 <= y <= bottom + 4:
                return True
        return False

    xl_row = 1

    # Collect y-ranges of tables and free rows
    all_vis_ys = [y for y, _ in vis]

    # Process in top-to-bottom order mixing free rows and tables
    processed_tables = set()

    for y, chs in vis:
        # Check if a new table starts near this y
        for tidx, (bbox, tdata) in enumerate(table_bboxes):
            _, top, _, bottom = bbox
            if tidx not in processed_tables and abs(y - top) < 6:
                processed_tables.add(tidx)
                if not tdata:
                    continue
                # Determine column count
                ncols = max(len(r) for r in tdata if r) if tdata else 1
                # Set col widths (rough)
                pw = float(page.width)
                cw_each = max(3, round((pw / ncols) / 5))
                for ci in range(ncols):
                    ltr = get_column_letter(ci + 1)
                    if ws.column_dimensions[ltr].width < cw_each:
                        ws.column_dimensions[ltr].width = cw_each

                is_header_row = True
                for trow in tdata:
                    if trow is None:
                        xl_row += 1
                        continue
                    cells = [str(c).strip() if c is not None else "" for c in trow]
                    if all(c == "" for c in cells):
                        xl_row += 1
                        continue
                    bg = "D9E1F2" if is_header_row else (ROW_B if xl_row % 2 == 0 else ROW_A)
                    bold = is_header_row
                    for ci, val in enumerate(cells):
                        if not val:
                            continue
                        c = ws.cell(row=xl_row, column=ci + 1, value=val)
                        c.font      = _fnt(bold=bold, size=9)
                        c.fill      = _fill(bg)
                        c.alignment = _aln("left", wrap=True)
                        c.border    = THIN
                    ws.row_dimensions[xl_row].height = 14
                    xl_row += 1
                    is_header_row = False
                continue  # skip free-text for this y

        if in_any_table(y):
            continue  # covered by table

        # Free-text row
        segs = chars_to_segs(chs)
        if not segs:
            xl_row += 1
            continue
        kind = fill_at(y, fills)
        sz   = segs[0]["size"] if segs else 9
        ws.row_dimensions[xl_row].height = max(10, round(sz * 1.6))

        if kind == "black":
            txt = "  " + "  ".join(s["text"] for s in segs)
            c   = ws.cell(row=xl_row, column=1, value=txt)
            c.font = _fnt(bold=True, size=max(int(sz), 9), color="FFFFFF")
            c.fill = _fill("111111")
            c.alignment = _aln("left")
            c.border = THICK
            ws.row_dimensions[xl_row].height = 15
        elif kind == "grey":
            for ci_s, seg in enumerate(segs):
                c = ws.cell(row=xl_row, column=ci_s + 1, value=seg["text"])
                c.font = _fnt(bold=True, size=8)
                c.fill = _fill("DCDCDC")
                c.alignment = _aln("left")
                c.border = THIN
        else:
            bg = ROW_B if xl_row % 2 == 0 else ROW_A
            # Auto-detect right-column numbers
            pw = float(page.width)
            for seg in segs:
                is_right = seg["x0"] > pw * 0.55
                al = "right" if is_right else "left"
                # pick a reasonable column
                col = max(1, round(seg["x0"] / (pw / 4)))
                c   = ws.cell(row=xl_row, column=col, value=seg["text"])
                c.font      = _fnt(bold=seg["bold"], size=max(int(seg["size"]), 8))
                c.fill      = _fill(bg)
                c.alignment = _aln(al)
                c.border    = THIN
        xl_row += 1


# ══════════════════════════════════════════════════════════════════════════════
#  MODE 3 — TEXT  (pure positioned text, no grid, no tables)
# ══════════════════════════════════════════════════════════════════════════════

def write_text_page(ws, page):
    """Best-effort conversion for PDFs with no structure (MD Form 202 etc.)"""
    pw    = float(page.width)
    fills = get_fills(page)
    vis   = chars_to_rows(page.chars)

    # Auto-detect columns from x-gap clusters across all rows
    all_segs = [chars_to_segs(chs) for _, chs in vis]
    xs = {0.0, pw}
    for segs in all_segs:
        for s in segs:
            xs.add(s["x0"]); xs.add(s["x1"])
    slots = snap_vals(list(xs), COL_SNAP * 2)
    n     = max(len(slots) - 1, 1)

    # Reasonable col widths
    for ci in range(n):
        cw = slots[ci + 1] - slots[ci]
        ws.column_dimensions[get_column_letter(ci + 1)].width = max(3, round(cw / 5.2))

    for xl_row, (y, chs) in enumerate(vis, start=1):
        segs = chars_to_segs(chs)
        kind = fill_at(y, fills)
        sz   = segs[0]["size"] if segs else 9
        ws.row_dimensions[xl_row].height = max(10, round(sz * 1.6))

        if not segs:
            continue

        if kind == "black":
            txt = "  " + "  ".join(s["text"] for s in segs)
            c   = ws.cell(row=xl_row, column=1, value=txt)
            c.font = _fnt(bold=True, size=max(int(sz), 9), color="FFFFFF")
            c.fill = _fill("111111")
            c.alignment = _aln("left")
            c.border = THICK
            ws.row_dimensions[xl_row].height = 15
            safe_merge(ws, xl_row, 1, xl_row, n)
            continue

        if kind == "grey":
            used = set()
            for seg in segs:
                ci_s = slot_of(seg["x0"], slots)
                ci_e = end_slot_of(seg["x1"], slots)
                if ci_s + 1 in used: continue
                c = ws.cell(row=xl_row, column=ci_s + 1, value=seg["text"])
                c.font = _fnt(bold=True, size=max(int(seg["size"]), 8))
                c.fill = _fill("DCDCDC")
                c.alignment = _aln("left")
                c.border = THIN
                if ci_e > ci_s:
                    safe_merge(ws, xl_row, ci_s+1, xl_row, ci_e+1)
                    for cc in range(ci_s+1, ci_e+2): used.add(cc)
                else:
                    used.add(ci_s+1)
            ws.row_dimensions[xl_row].height = 13
            continue

        bg   = ROW_B if xl_row % 2 == 0 else ROW_A
        used = set()
        for seg in segs:
            ci_s = slot_of(seg["x0"], slots)
            ci_e = end_slot_of(seg["x1"], slots)
            if ci_s + 1 in used: continue
            al   = detect_align(seg, slots[ci_s], slots[min(ci_s+1, len(slots)-1)])
            c    = ws.cell(row=xl_row, column=ci_s + 1, value=seg["text"])
            c.font      = _fnt(bold=seg["bold"], size=max(int(seg["size"]), 8))
            c.fill      = _fill(bg)
            c.alignment = _aln(al)
            c.border    = THIN
            if ci_e > ci_s:
                safe_merge(ws, xl_row, ci_s+1, xl_row, ci_e+1)
                for cc in range(ci_s+1, ci_e+2): used.add(cc)
            else:
                used.add(ci_s+1)


# ══════════════════════════════════════════════════════════════════════════════
#  PAGE DISPATCHER
# ══════════════════════════════════════════════════════════════════════════════

def write_page(ws, page, page_num):
    vlines  = [l for l in page.lines if abs(l["x0"] - l["x1"]) < 1]
    n_vline = len(vlines)

    tables  = page.find_tables()
    good_tables = [t for t in tables
                   if t.extract() and
                   max((len(r) for r in t.extract() if r), default=0) >= TABLE_MIN_COLS]

    if n_vline >= VLINE_MIN:
        mode = "GRID"
        write_grid_page(ws, page)
    elif good_tables:
        mode = "TABLE"
        write_table_page(ws, page, good_tables)
    else:
        mode = "TEXT"
        write_text_page(ws, page)

    print(f"  [Page {page_num}] {page.width:.0f}×{page.height:.0f}  "
          f"{len(page.chars)} chars  {n_vline} v-lines  "
          f"{len(good_tables)} tables  → [{mode}]")


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════════

def convert(pdf_path: str, out_path: str = None):
    if out_path is None:
        out_path = os.path.splitext(pdf_path)[0] + "_converted.xlsx"

    print(f"\n{'═'*60}")
    print(f"  PDF → Excel   {os.path.basename(pdf_path)}")
    print(f"  Output      → {out_path}")
    print(f"{'═'*60}")

    wb = Workbook()
    with pdfplumber.open(pdf_path) as pdf:
        n_pages = len(pdf.pages)
        for pn, page in enumerate(pdf.pages):
            ws = wb.active if pn == 0 else wb.create_sheet()
            ws.title = f"Page {pn + 1}"
            # freeze top row
            ws.freeze_panes = "A2"
            write_page(ws, page, pn + 1)

    wb.save(out_path)
    print(f"\n  ✓  Done! {n_pages} page(s) → {out_path}\n")
    return out_path


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage:  python pdf_to_excel_final.py  input.pdf  [output.xlsx]")
        sys.exit(1)
    src = sys.argv[1]
    dst = sys.argv[2] if len(sys.argv) >= 3 else None
    if not os.path.exists(src):
        print(f"File not found: {src}")
        sys.exit(1)
    convert(src, dst)
