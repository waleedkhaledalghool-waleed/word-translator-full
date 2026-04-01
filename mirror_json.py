"""
mirror_json.py
──────────────
Mirrors a pipeline JSON document's layout when translating between
languages with opposite text directions (LTR ↔ RTL).

Use cases:
  English → Arabic   (LTR source → RTL target)  : mirror ON
  Arabic  → English  (RTL source → LTR target)  : mirror ON
  English → French   (LTR → LTR)                : no-op passthrough
  Arabic  → Farsi    (RTL → RTL)                : no-op passthrough

What "mirroring" does:
  • Page margins: left ↔ right swapped
  • Paragraph alignment: left ↔ right swapped (center/justify untouched)
  • Paragraph indentation: left ↔ right swapped
  • Paragraph bidi flag: flipped
  • Run rtl flag: flipped
  • Tab stop types: alignment flipped (left↔right); positions kept as-is
    because Word re-measures bidi-paragraph tab positions from the right margin
  • Text box horizontal position: mirrored around page text width
  • Table: column order reversed per row

Output:
  mirrored_json/<basename>.json

Usage
─────
  python mirror_json.py <file.json> [out_dir] [src_lang] [tgt_lang]

  file.json   — full pipeline JSON (to_json/<n>.json)
  out_dir     — output folder (default: mirrored_json/)
  src_lang    — source language code or "auto"  (default: en)
  tgt_lang    — target language code            (default: ar)
"""

import json
import os
import sys
from copy import deepcopy


# ── RTL language sets ─────────────────────────────────────────────────────────

RTL_LANGS = {
    "ar",   # Arabic
    "he",   # Hebrew
    "fa",   # Persian/Farsi
    "ur",   # Urdu
    "yi",   # Yiddish
    "dv",   # Divehi/Maldivian
    "ps",   # Pashto
    "ku",   # Kurdish (Sorani)
    "ug",   # Uyghur
    "sd",   # Sindhi
    "ckb",  # Central Kurdish
}


def _is_rtl(lang: str) -> bool:
    """Return True if *lang* is a right-to-left language code."""
    if not lang:
        return False
    return lang.strip().lower() in RTL_LANGS


def needs_mirror(src_lang: str, tgt_lang: str) -> bool:
    """
    Return True when source and target have opposite text directions.
    'auto' source is treated as LTR (most common case for auto-detection).
    """
    src = src_lang.strip().lower() if src_lang else "auto"
    if src == "auto":
        src_rtl = False   # assume LTR when auto
    else:
        src_rtl = _is_rtl(src)
    tgt_rtl = _is_rtl(tgt_lang)
    return src_rtl != tgt_rtl


# ── Alignment flip ────────────────────────────────────────────────────────────

_ALIGN_FLIP = {
    "left":  "right",
    "right": "left",
    # center, justify, both, distribute → unchanged
}

def _flip_align(align: str | None) -> str | None:
    if align is None:
        return None
    return _ALIGN_FLIP.get(align, align)


# ── Core mirror logic ─────────────────────────────────────────────────────────

def _mirror_page(page: dict) -> dict:
    """
    Swap left/right margins on the page-level metadata dict.
    Also compute text_width (used for tab-stop/textbox mirroring).
    """
    p = dict(page)
    ml = p.get("margin_left",  1800)
    mr = p.get("margin_right", 1800)
    p["margin_left"]  = mr
    p["margin_right"] = ml
    return p


def _text_width(page: dict) -> int:
    """Return usable text width in twips."""
    w  = page.get("width",  12240)
    ml = page.get("margin_left",  1800)
    mr = page.get("margin_right", 1800)
    return max(w - ml - mr, 1)


def _mirror_tab_stops(tab_stops: list, text_width: int) -> list:
    """
    Mirror tab stops when flipping a paragraph from LTR → RTL (or vice versa).

    KEY INSIGHT about Word's bidi tab stop model
    ─────────────────────────────────────────────
    In a bidi (RTL) paragraph Word automatically mirrors the tab grid — it
    measures positions from the RIGHT margin instead of the left. This means:

      • An LTR paragraph with a left-aligned tab at pos=4683 (from left)
        places the cursor 4683 twips from the LEFT edge.

      • After flipping to RTL/bidi, that same pos=4683 is measured 4683 twips
        from the RIGHT edge — the correct mirrored physical column.

    So TWO things must NOT change:
      1. The position value — Word re-measures it from the other side.
      2. The alignment type — "left" in bidi means "anchor to the logical
         start side of this stop" which is visually the right side. Flipping
         it to "right" would double-invert and produce wrong column anchoring.

    Result: tab stops are passed through completely unchanged.
    """
    # Return a shallow copy of each stop with no modifications.
    return [dict(tab) for tab in tab_stops]


def _mirror_paragraph(para: dict, text_width: int) -> dict:
    """Return a new paragraph dict with direction-related fields mirrored."""
    p = dict(para)

    # Alignment
    if "align" in p:
        p["align"] = _flip_align(p["align"])

    # Indentation: swap left ↔ right.
    #
    # w:ind left/right are PHYSICAL page distances in OOXML — they always mean
    # "distance from the left/right page margin" regardless of bidi direction.
    # Swapping them mirrors the physical indent, which is exactly what we want:
    # a clearance gap that was on the right (e.g. for a right-side photo box)
    # moves to the left (for the now-left photo box).
    #
    # HOWEVER: hanging/first-line indent is relative to the paragraph START SIDE.
    # In a bidi paragraph the start side is the right, so indent_hanging stays
    # attached to the right — we leave it as-is (no swap needed).
    il = p.pop("indent_left",  None)
    ir = p.pop("indent_right", None)
    if il is not None: p["indent_right"] = il
    if ir is not None: p["indent_left"]  = ir

    # Bidi flag: flip
    if "bidi" in p:
        p["bidi"] = not p["bidi"]
    # If no bidi flag was present, after mirror a formerly-LTR doc going RTL
    # needs bidi=True on every paragraph. We set it here.
    # (We only do this if there's a "runs" key — i.e. it's a real content para)
    if "bidi" not in p:
        p["bidi"] = True   # will be set False again on RTL→LTR by the not above

    # Tab stops
    if "tab_stops" in p:
        p["tab_stops"] = _mirror_tab_stops(p["tab_stops"], text_width)
    # For bidi paragraphs with no explicit tab stops, Word falls back to its
    # default 720-twip grid measured from the RIGHT — which may not match the
    # original layout. Add no implicit stops here; if the original had none,
    # the bidi default grid will be used (which is usually fine).

    # Runs: flip rtl flag
    if "runs" in p:
        new_runs = []
        for run in p["runs"]:
            r = dict(run)
            # flip rtl: absent means False
            r["rtl"] = not r.get("rtl", False)
            new_runs.append(r)
        p["runs"] = new_runs

    return p


def _mirror_table(table: dict, text_width: int) -> dict:
    """Mirror a table: reverse column order per row."""
    t = dict(table)

    # Mirror table-level alignment
    if "align" in t:
        t["align"] = _flip_align(t["align"])

    if "rows" not in t:
        return t

    new_rows = []
    for row in t["rows"]:
        new_row = dict(row)
        if "cells" in row:
            new_row["cells"] = list(reversed([
                _mirror_cell(cell, text_width) for cell in row["cells"]
            ]))
        new_rows.append(new_row)
    t["rows"] = new_rows
    return t


def _mirror_cell(cell: dict, text_width: int) -> dict:
    """Mirror cell contents (paragraph alignment/direction)."""
    c = dict(cell)
    if "paragraphs" in c:
        c["paragraphs"] = [
            _mirror_element(p, text_width) for p in c["paragraphs"]
        ]
    return c


def _mirror_textbox(tb: dict, text_width: int, page_width: int,
                    margin_left: int, margin_right: int) -> dict:
    """
    Mirror a textbox/floating element horizontally.

    The extractor (word_to_json.py) stores textbox position as:
      pos_x        — horizontal offset in twips from the anchor reference's LEFT edge
      pos_h_anchor — "page"   : pos_x is from the physical page left edge
                     "margin" : pos_x is from the left margin (= text area left edge)
                     "column" : same as margin for single-column docs

    Mirroring formula — move the box so its RIGHT edge is the same distance
    from the RIGHT reference boundary as it was from the LEFT:

      anchor == "page":
        new_pos_x = page_width  - pos_x - box_width

      anchor == "margin" / "column":
        new_pos_x = text_width  - pos_x - box_width

    This is exactly like CSS `right: X` ↔ `left: X` where X is the gap
    between the box edge and the reference boundary — nothing goes off-page.
    """
    t = dict(tb)

    box_w  = t.get("width", 0)
    pos_x  = t.get("pos_x", 0)
    anchor = t.get("pos_h_anchor", "page")

    if anchor == "page":
        new_pos_x = page_width - pos_x - box_w
    elif anchor in ("margin", "column"):
        new_pos_x = text_width - pos_x - box_w
    else:
        # "char" or any unknown anchor: can't reliably mirror, leave as-is
        new_pos_x = pos_x

    t["pos_x"] = max(new_pos_x, 0)

    # Mirror paragraphs inside the textbox (alignment + run direction)
    if "paragraphs" in t:
        t["paragraphs"] = [
            _mirror_element(p, text_width) for p in t["paragraphs"]
        ]

    return t


def _mirror_element(el: dict, text_width: int,
                    page_width: int = 0,
                    margin_left: int = 0, margin_right: int = 0) -> dict:
    """Dispatch mirroring to the right handler based on element type."""
    etype = el.get("type", "")

    if etype in ("paragraph", "empty_paragraph"):
        return _mirror_paragraph(el, text_width)

    elif etype == "table":
        return _mirror_table(el, text_width)

    elif etype == "textbox":
        return _mirror_textbox(el, text_width, page_width, margin_left, margin_right)

    else:
        # image, unknown — return as-is
        return el


# ── Vertical line-height calculation ─────────────────────────────────────────

def _para_line_height(el: dict, bidi: bool) -> int:
    """
    Estimate the rendered line height of a paragraph element in twips.

    Word uses the COMPLEX SCRIPT font size (size_cs) for line-height
    calculations when a paragraph has bidi=True.  This is why mirroring
    paragraphs that have size_cs > size makes them taller, shifting
    everything below them downward.

    We replicate Word's rule:
      effective_size = size_cs if bidi else size (in points)
      auto/multiple  → effective_size_twips * line_spacing / 240
      atLeast        → max(effective_size_twips, line_spacing)
      exact          → line_spacing  (fixed, ignores font)
    """
    if el.get("type") not in ("paragraph", "empty_paragraph"):
        return 0

    # Collect effective font size (pt → twips)
    # Priority: explicit run size → pmark_fmt size → default 12pt
    size_pt    = None
    size_cs_pt = None

    pmark = el.get("pmark_fmt", {})
    if pmark.get("size"):    size_pt    = float(pmark["size"])
    if pmark.get("size_cs"): size_cs_pt = float(pmark["size_cs"])

    # Also check first run
    for run in el.get("runs", []):
        if size_pt    is None and run.get("size"):    size_pt    = float(run["size"])
        if size_cs_pt is None and run.get("size_cs"): size_cs_pt = float(run["size_cs"])
        break  # first run is enough

    if size_pt    is None: size_pt    = 12.0
    if size_cs_pt is None: size_cs_pt = size_pt

    effective_pt    = size_cs_pt if bidi else size_pt
    effective_twips = effective_pt * 20

    ls      = el.get("line_spacing", 240)
    ls_rule = el.get("line_spacing_rule", "auto")

    if ls_rule == "exact":
        line_h = ls
    elif ls_rule == "atLeast":
        line_h = max(effective_twips, ls)
    else:  # auto / multiple  (ls is in 240ths: 240=single, 480=double)
        line_h = effective_twips * ls / 240

    return int(line_h
               + el.get("space_before", 0)
               + el.get("space_after",  0))


def _calc_textbox_y_delta(elements: list) -> int:
    """
    Calculate how much further down (in twips) textboxes should move after
    mirroring, due to paragraphs above them becoming taller.

    For each textbox found in elements, we:
      1. Sum the line heights of all non-textbox elements ABOVE it using the
         ORIGINAL (LTR, bidi=False) heights.
      2. Sum the same elements using MIRRORED (bidi=True) heights.
      3. The difference is how much the textbox's anchor point has shifted.

    Returns the delta for the FIRST textbox found (assumes all textboxes on
    the same page are affected by the same cumulative shift).
    """
    cumulative_orig   = 0
    cumulative_mirror = 0

    for el in elements:
        if el.get("type") == "textbox":
            # Return the delta accumulated up to (but not including) this box
            return cumulative_mirror - cumulative_orig

        orig_h   = _para_line_height(el, bidi=False)
        mirror_h = _para_line_height(el, bidi=True)
        cumulative_orig   += orig_h
        cumulative_mirror += mirror_h

    return 0  # no textbox found


# ── Public API ────────────────────────────────────────────────────────────────

def mirror_document(
    json_path : str,
    src_lang  : str = "en",
    tgt_lang  : str = "ar",
    out_dir   : str = "mirrored_json",
    force     : bool = False,
) -> tuple[str, bool]:
    """
    Mirror a pipeline JSON for direction change between src_lang and tgt_lang.

    Parameters
    ----------
    json_path : str   Path to the full pipeline JSON (to_json/<n>.json)
    src_lang  : str   Source language code or "auto"
    tgt_lang  : str   Target language code
    out_dir   : str   Output folder (created if absent)
    force     : bool  If True, mirror even when not strictly needed

    Returns
    -------
    (out_path, mirrored)   — out_path is the saved file, mirrored is True/False
    """
    mirror = force or needs_mirror(src_lang, tgt_lang)

    os.makedirs(out_dir, exist_ok=True)
    basename = os.path.splitext(os.path.basename(json_path))[0]
    out_path = os.path.join(out_dir, basename + ".json")

    with open(json_path, "r", encoding="utf-8") as fh:
        data = json.load(fh)

    if not mirror:
        # Passthrough: just copy the file unchanged
        with open(out_path, "w", encoding="utf-8") as fh:
            json.dump(data, fh, ensure_ascii=False, indent=2)
        _print_report(json_path, out_path, src_lang, tgt_lang, mirrored=False)
        return out_path, False

    # ── Deep copy so we don't mutate the loaded data ──────────────────────────
    doc = deepcopy(data)

    # ── Page / section metadata ───────────────────────────────────────────────
    page_orig = doc.get("meta", {}).get("page", {})
    page_new  = _mirror_page(page_orig)
    doc.setdefault("meta", {})["page"] = page_new

    tw   = _text_width(page_new)          # text width after margin swap
    pw   = page_new.get("width", 12240)
    ml   = page_new.get("margin_left",  1800)
    mr   = page_new.get("margin_right", 1800)

    # ── Textbox Y correction ──────────────────────────────────────────────────
    # Paragraphs whose size_cs > size grow taller when bidi=True is applied.
    # This pushes the rendered position of later paragraphs downward, but
    # textbox pos_y (absolute from page top) stays fixed — making boxes
    # appear too high relative to the text rows they should sit beside.
    # We calculate the cumulative height delta and add it to each textbox's
    # pos_y so the box tracks with its surrounding text.
    elements_orig = doc.get("elements", [])
    y_delta = _calc_textbox_y_delta(elements_orig)

    # ── Elements ──────────────────────────────────────────────────────────────
    mirrored_elements = []
    for el in elements_orig:
        mel = _mirror_element(el, tw, pw, ml, mr)
        # Apply Y correction to page/margin-anchored textboxes
        if mel.get("type") == "textbox" and y_delta != 0:
            v_anchor = mel.get("pos_v_anchor", "page")
            if v_anchor in ("page", "margin"):
                mel = dict(mel)
                mel["pos_y"] = mel.get("pos_y", 0) + y_delta
        mirrored_elements.append(mel)
    doc["elements"] = mirrored_elements

    # ── Store mirror metadata ─────────────────────────────────────────────────
    doc.setdefault("meta", {})["mirror"] = {
        "src_lang": src_lang,
        "tgt_lang": tgt_lang,
        "applied":  True,
        "y_delta":  y_delta,
    }

    with open(out_path, "w", encoding="utf-8") as fh:
        json.dump(doc, fh, ensure_ascii=False, indent=2)

    _print_report(json_path, out_path, src_lang, tgt_lang, mirrored=True)
    return out_path, True


# ── Terminal report ───────────────────────────────────────────────────────────

_RESET  = "\033[0m"
_BOLD   = "\033[1m"
_CYAN   = "\033[96m"
_GREEN  = "\033[92m"
_YELLOW = "\033[93m"
_MUTED  = "\033[90m"


def _c(text, code):
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.kernel32.SetConsoleMode(
                ctypes.windll.kernel32.GetStdHandle(-11), 7)
        except Exception:
            return text
    return f"{code}{text}{_RESET}"


def _print_report(src, dst, src_lang, tgt_lang, mirrored):
    sep = _c("─" * 70, _MUTED)
    print()
    print(sep)
    print(_c("  mirror_json  ·  report", _BOLD))
    print(sep)
    print(f"  source  : {_c(src, _CYAN)}")
    print(f"  output  : {_c(dst, _GREEN)}")
    print(f"  langs   : {src_lang}  →  {tgt_lang}")
    if mirrored:
        print(_c("  ✓ Document mirrored  (LTR ↔ RTL direction flip applied)", _GREEN))
    else:
        dir_src = "RTL" if _is_rtl(src_lang) else "LTR"
        dir_tgt = "RTL" if _is_rtl(tgt_lang) else "LTR"
        print(_c(f"  – Passthrough  ({dir_src} → {dir_tgt}, no flip needed)", _MUTED))
    print(sep)
    print()


# ── CLI ───────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(0)

    json_arg    = sys.argv[1]
    out_dir_arg = sys.argv[2] if len(sys.argv) > 2 else "mirrored_json"
    src_arg     = sys.argv[3] if len(sys.argv) > 3 else "en"
    tgt_arg     = sys.argv[4] if len(sys.argv) > 4 else "ar"

    try:
        path, did_mirror = mirror_document(json_arg, src_arg, tgt_arg, out_dir_arg)
        print(f"{'Mirrored' if did_mirror else 'Passthrough'} → {path}")
        sys.exit(0)
    except (FileNotFoundError, ValueError) as exc:
        print(f"\n  ✗ Error: {exc}\n", file=sys.stderr)
        sys.exit(1)