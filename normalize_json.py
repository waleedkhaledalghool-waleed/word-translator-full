"""
normalize_json.py
─────────────────
Post-translation cleanup step.

Normalizes "field paragraphs" (label : value) in a translated pipeline JSON
so all colons align to the same column in the final Word document.

THE PROBLEM
───────────
After translation, field paragraphs look like:

    "الاسم  \t\t\t : \t أحمد سعيد..."       (short label, 3 tabs needed)
    "تاريخ الانتهاء \t : \t 2026/07/16"      (long label, 1 tab needed)
    "مكان العمل \t\t : \t منطقة الرياض"      (medium label, 2 tabs needed)

The number of tabs was chosen for Latin text widths. Arabic labels have
different widths, so the colon column zigzags visually.

THE FIX
───────
1. Detect field paragraphs — any paragraph whose runs contain a structural
   colon separator ( ":", "\t:", "\t: ", ": " etc. in a non-word run ).

2. Normalize every such paragraph to the structure:
     label_run  |  tab_run  |  " : "  |  tab_run  |  value_run

3. Apply IDENTICAL tab stops to every normalized paragraph:
     - One RIGHT tab at COLON_POS  → colon always snaps to same column
     - One LEFT  tab at VALUE_POS  → value always starts at same column

   In a bidi (RTL) document, "right" and "left" refer to the visual positions
   from the left edge of the text area (same as LTR — Word handles the mirroring).

4. The paragraph's existing spacing, font, bidi, style are untouched.
   Only the runs[] list and tab_stops[] are replaced.

USAGE
─────
  python normalize_json.py <translated.json> [out_dir]

  Input:  translated pipeline JSON (output of merge_compact / reinsert_compact)
  Output: normalized JSON in out_dir/ (default: normalized_json/)

  Or call normalize_document(data) directly on a loaded dict.

TAB STOP TUNING
───────────────
  COLON_POS : twips from left margin where " : " right-aligns.
              For RTL documents, increase this (closer to right = closer to labels).
  VALUE_POS : twips from left margin where value text starts.
              For RTL documents, this is where values begin (left of colon).

  For the Saudi Resident ID card (text width ≈ 9418 twips, RTL):
    COLON_POS = 6780  → colon at ~72% of text width (close to right-side labels)
    VALUE_POS = 4900  → value starts at ~52% (left of colon, toward center-left)
"""

import json
import os
import unicodedata

# ── Tab stop positions (twips) ────────────────────────────────────────────────
# Tune these for your document. 1 inch = 1440 twips.
COLON_POS = 2300   # right-tab: colon aligns here — for RTL docs, higher = closer to right-side labels
VALUE_POS = 2600   # left-tab:  value text starts here — for RTL docs, this is left of the colon

# ── Helpers ───────────────────────────────────────────────────────────────────

_STRUCTURAL = frozenset(
    "\t\n\r\u00a0\u200f\u200e"
    "\u202a\u202b\u202c\u202d\u202e\ufeff"
    "\u2192\u2190\u2194\u2191\u2193"
    "\u25ba\u25c4\u25b6\u25c0"
)

def _has_letters(text: str) -> bool:
    return any(unicodedata.category(ch).startswith("L") for ch in text)

def _is_pure_structural(text: str) -> bool:
    """True if every character is a structural/whitespace char."""
    return bool(text) and all(ch in _STRUCTURAL for ch in text)

def _has_content(text: str) -> bool:
    """True if text has visible non-structural content (letters, digits, punctuation)."""
    return bool(text.strip()) and not _is_pure_structural(text)


# ── Core normalizer ───────────────────────────────────────────────────────────

def _normalize_field_para(para: dict) -> bool:
    """
    If para is a field paragraph (has a structural colon separator),
    rewrite its runs[] and tab_stops[] in-place for uniform alignment.

    Returns True if the paragraph was modified.
    """
    runs = para.get("runs")
    if not runs or not isinstance(runs, list):
        return False

    # ── Find the colon separator run ─────────────────────────────────────────
    # It's a run whose text contains ":" but has no letters
    colon_idx = None
    for i, r in enumerate(runs):
        t = r.get("text", "")
        if ":" in t and not _has_letters(t):
            colon_idx = i
            break

    if colon_idx is None:
        return False   # not a field paragraph

    # ── Collect label runs (before colon) ─────────────────────────────────────
    label_runs = [r for r in runs[:colon_idx] if _has_letters(r.get("text", ""))]
    if not label_runs:
        return False

    # Smart-join: preserve internal spaces, trim boundary spaces
    parts = [r["text"] for r in label_runs]
    label_text = parts[0]
    for p in parts[1:]:
        if not label_text.endswith(" ") and not p.startswith(" "):
            label_text += " "
        label_text += p
    label_text = label_text.strip()

    # Inherit formatting from first label run
    label_fmt = {k: v for k, v in label_runs[0].items() if k != "text"}

    # ── Collect value runs (after colon) ─────────────────────────────────────
    # Include any run with visible content: letters, numbers, dates, etc.
    value_runs = [r for r in runs[colon_idx + 1:] if _has_content(r.get("text", ""))]

    if value_runs:
        # Join value parts preserving internal structure
        value_parts = [r["text"] for r in value_runs]
        value_text  = value_parts[0]
        for p in value_parts[1:]:
            if not value_text.endswith(" ") and not p.startswith(" "):
                value_text += " "
            value_text += p
        value_text = value_text.strip()
        value_fmt  = {k: v for k, v in value_runs[0].items() if k != "text"}
    else:
        value_text = ""
        value_fmt  = label_fmt

    # ── Build normalized runs ─────────────────────────────────────────────────
    # Structure: label | \t | " : " | \t | value
    new_runs = [
        {"text": label_text, **label_fmt},
        {"text": "\t",       **label_fmt},
        {"text": " : ",      **label_fmt},
        {"text": "\t",       **label_fmt},
        {"text": value_text, **value_fmt},
    ]

    # ── Apply uniform tab stops ───────────────────────────────────────────────
    para["tab_stops"] = [
        {"pos": COLON_POS, "type": "right", "leader": "none"},
        {"pos": VALUE_POS, "type": "left",  "leader": "none"},
    ]

    para["runs"] = new_runs
    para["text"] = "".join(r["text"] for r in new_runs)

    return True


# ── Document walker ───────────────────────────────────────────────────────────

def normalize_document(data: dict) -> int:
    """
    Walk all paragraphs in a pipeline JSON document and normalize field paragraphs.
    Skips paragraphs inside table cells — tables manage their own layout.
    Returns the number of paragraphs modified.
    """
    count = [0]

    def _walk(obj, in_table: bool = False):
        if isinstance(obj, dict):
            node_type = obj.get("type")

            # Once inside a table, all descendants are in_table
            if node_type == "table":
                for v in obj.values():
                    _walk(v, in_table=True)
                return

            if node_type == "paragraph":
                # Only normalize top-level (non-table) paragraphs
                if not in_table:
                    if _normalize_field_para(obj):
                        count[0] += 1
                return  # don't recurse into runs

            for v in obj.values():
                _walk(v, in_table=in_table)

        elif isinstance(obj, list):
            for item in obj:
                _walk(item, in_table=in_table)

    _walk(data)
    return count[0]


# ── File-level entry point ─────────────────────────────────────────────────────

def normalize_file(
    json_path: str,
    out_dir:   str = "normalized_json",
) -> str:
    """
    Normalize field paragraphs in a translated pipeline JSON and save to out_dir.

    Parameters
    ----------
    json_path : str   Path to the translated JSON
    out_dir   : str   Output folder (created if absent)

    Returns the output path.
    """
    os.makedirs(out_dir, exist_ok=True)
    basename = os.path.basename(json_path)
    out_path = os.path.join(out_dir, basename)

    with open(json_path, "r", encoding="utf-8") as fh:
        data = json.load(fh)

    n = normalize_document(data)

    with open(out_path, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)

    print(
        f"[normalize_json] {n} field paragraphs normalized  "
        f"(colon @ {COLON_POS} twips, value @ {VALUE_POS} twips)  "
        f"→ {out_path}"
    )
    return out_path


# ── CLI ───────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python normalize_json.py <translated.json> [out_dir]")
        print()
        print("Normalizes field paragraphs (label : value) for uniform colon alignment.")
        print(f"Tab stops: colon right-tab @ {COLON_POS} twips, value left-tab @ {VALUE_POS} twips")
        sys.exit(1)

    _json_path = sys.argv[1]
    _out_dir   = sys.argv[2] if len(sys.argv) > 2 else "normalized_json"

    result = normalize_file(_json_path, _out_dir)
    print(f"Done → {result}")