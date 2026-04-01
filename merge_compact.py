"""
merge_compact.py
────────────────
Takes an edited compact JSON  (compact/<n>.compact.json)
and the original full JSON    (to_json/<n>.json)
and produces a merged JSON    (merged_json/<n>.json)

Only the text values that were actually changed are reinserted.
Everything else — every font, size, spacing, table, border, margin —
is carried over byte-for-byte from the original.

Terminal output
───────────────
Prints a diff of ONLY the run texts that changed:

  [  3]  "KINGDOM OF SAUDI ARABIA"
       → "المملكة العربية السعودية"

  [ 12]  "Name:"
       → "الاسم:"

No output for unchanged texts — only real edits are shown.

Usage
─────
  python merge_compact.py <compact.json> [full.json] [out_dir]

  compact.json  — edited compact file (required)
  full.json     — original full JSON  (auto-discovered from to_json/ if omitted)
  out_dir       — output folder       (default: merged_json/)
"""

import json
import os
import sys
from copy import deepcopy


# ────────────────────────────────────────────────────────────────────────────
# Internal helpers (shared with compact_json.py logic)
# ────────────────────────────────────────────────────────────────────────────

def _collect_run_refs(obj: object, out: list) -> None:
    """
    Recursively walk *obj* in document order and append every run-dict
    that contains a "text" key.  Paragraph-level "text" fields are
    intentionally skipped — they are regenerated after reinsertion.
    """
    if isinstance(obj, dict):
        runs = obj.get("runs")
        if isinstance(runs, list):
            for run in runs:
                if isinstance(run, dict) and "text" in run:
                    out.append(run)
        for key, value in obj.items():
            if key == "runs":
                continue
            _collect_run_refs(value, out)
    elif isinstance(obj, list):
        for item in obj:
            _collect_run_refs(item, out)


def _regen_paragraph_texts(obj: object) -> None:
    """Rebuild every paragraph.text as the joined text of its runs."""
    if isinstance(obj, dict):
        if "runs" in obj and "text" in obj and isinstance(obj["runs"], list):
            obj["text"] = "".join(r.get("text", "") for r in obj["runs"])
        for v in obj.values():
            _regen_paragraph_texts(v)
    elif isinstance(obj, list):
        for item in obj:
            _regen_paragraph_texts(item)


# ────────────────────────────────────────────────────────────────────────────
# ANSI colours for terminal output
# ────────────────────────────────────────────────────────────────────────────

_RESET  = "\033[0m"
_RED    = "\033[91m"
_GREEN  = "\033[92m"
_CYAN   = "\033[96m"
_YELLOW = "\033[93m"
_BOLD   = "\033[1m"
_MUTED  = "\033[90m"

def _c(text, code):
    """Wrap text in ANSI colour (skip on Windows without ANSI support)."""
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.kernel32.SetConsoleMode(
                ctypes.windll.kernel32.GetStdHandle(-11), 7)
        except Exception:
            return text
    return f"{code}{text}{_RESET}"


# ────────────────────────────────────────────────────────────────────────────
# Core function
# ────────────────────────────────────────────────────────────────────────────

def merge_and_diff(
    compact_path : str,
    full_json_path: str | None = None,
    out_dir      : str = "merged_json",
) -> tuple[str, int, int]:
    """
    Merge an edited compact JSON back into the full pipeline JSON,
    print a diff of changed texts, and save the merged file.

    Parameters
    ----------
    compact_path   : str   Path to the edited compact JSON
    full_json_path : str   Path to the full pipeline JSON.
                           If None, auto-discovered from to_json/.
    out_dir        : str   Output folder (created if absent)

    Returns
    -------
    (out_path, total_runs, changed_count)
    """

    # ── Resolve full JSON path ────────────────────────────────────────────────
    if full_json_path is None:
        base = os.path.basename(compact_path)          # e.g.  doc.compact.json
        stem = base.replace(".compact.json", "")       # e.g.  doc
        if not stem or stem == base:
            # fallback: strip one extension
            stem = os.path.splitext(stem or base)[0]

        # Search relative to compact file, then to_json/
        candidates = [
            os.path.join(os.path.dirname(compact_path), "..", "to_json", stem + ".json"),
            os.path.join("to_json", stem + ".json"),
            os.path.join(os.path.dirname(compact_path), stem + ".json"),
        ]
        for c in candidates:
            if os.path.exists(c):
                full_json_path = os.path.normpath(c)
                break

        if full_json_path is None:
            raise FileNotFoundError(
                f"Cannot find full JSON for '{compact_path}'.\n"
                f"Tried: {candidates}\n"
                f"Pass the path explicitly as the second argument."
            )

    # ── Load files ────────────────────────────────────────────────────────────
    with open(full_json_path, "r", encoding="utf-8") as fh:
        data = json.load(fh)

    with open(compact_path, "r", encoding="utf-8") as fh:
        compact = json.load(fh)

    new_texts : list[str] = compact["texts"]

    # ── Collect live run references from the cloned data tree ─────────────────
    merged_data = deepcopy(data)
    run_refs: list[dict] = []
    _collect_run_refs(merged_data, run_refs)

    total_runs = len(run_refs)

    if total_runs != len(new_texts):
        raise ValueError(
            f"Run count mismatch!\n"
            f"  Full JSON has {total_runs} runs.\n"
            f"  Compact JSON has {len(new_texts)} entries.\n"
            f"Make sure the compact file was generated from this exact full JSON\n"
            f"and that you have not added or removed any entries."
        )

    # ── Collect original texts for diffing ────────────────────────────────────
    orig_refs: list[dict] = []
    _collect_run_refs(data, orig_refs)          # from the unmodified copy
    orig_texts = [r["text"] for r in orig_refs]

    # ── Apply new texts & collect diffs ───────────────────────────────────────
    diffs: list[tuple[int, str, str]] = []      # (index, old, new)

    for i, (run, new_text) in enumerate(zip(run_refs, new_texts)):
        old_text = orig_texts[i]
        run["text"] = new_text
        if new_text != old_text:
            diffs.append((i, old_text, new_text))

    # Regenerate paragraph.text fields
    _regen_paragraph_texts(merged_data)

    # ── Save merged JSON ──────────────────────────────────────────────────────
    os.makedirs(out_dir, exist_ok=True)
    stem    = os.path.basename(full_json_path)
    out_path = os.path.join(out_dir, stem)

    with open(out_path, "w", encoding="utf-8") as fh:
        json.dump(merged_data, fh, ensure_ascii=False, indent=2)

    # ── Terminal output ───────────────────────────────────────────────────────
    _print_diff_report(
        compact_path, full_json_path, out_path,
        total_runs, diffs,
    )

    return out_path, total_runs, len(diffs)


def _print_diff_report(
    compact_path  : str,
    full_path     : str,
    out_path      : str,
    total_runs    : int,
    diffs         : list[tuple[int, str, str]],
) -> None:
    """Pretty-print the diff report to stdout."""

    sep = _c("─" * 70, _MUTED)

    print()
    print(sep)
    print(_c("  merge_compact  ·  diff report", _BOLD))
    print(sep)
    print(f"  compact  : {_c(compact_path,  _CYAN)}")
    print(f"  source   : {_c(full_path,     _MUTED)}")
    print(f"  output   : {_c(out_path,      _GREEN)}")
    print(f"  total runs : {total_runs}")
    print(f"  changed    : {_c(str(len(diffs)), _YELLOW if diffs else _MUTED)}")
    print(sep)

    if not diffs:
        print()
        print(_c("  ✓ No differences — compact texts are identical to the original.", _MUTED))
        print()
        return

    for idx, old, new in diffs:
        idx_str  = _c(f"[{idx:>4}]", _MUTED)
        old_disp = _c(repr(old), _RED)
        new_disp = _c(repr(new), _GREEN)
        print()
        print(f"  {idx_str}  {old_disp}")
        print(f"  {' ':>6}→ {new_disp}")

    print()
    print(sep)
    print(_c(f"  ✓ Saved → {out_path}  ({len(diffs)} change{'s' if len(diffs) != 1 else ''})", _GREEN))
    print(sep)
    print()


# ────────────────────────────────────────────────────────────────────────────
# Batch helper  (used by pipeline_ui.py)
# ────────────────────────────────────────────────────────────────────────────

def merge_by_name(
    name          : str,
    dir_compact   : str = "compact",
    dir_json      : str = "to_json",
    out_dir       : str = "merged_json",
) -> tuple[str, int, int]:
    """
    Convenience wrapper that resolves paths from a document stem name.

    Parameters
    ----------
    name   : str   Original .docx filename  (e.g. "doc.docx")  or bare stem
    """
    stem         = os.path.splitext(os.path.basename(name))[0]
    compact_path = os.path.join(dir_compact, stem + ".compact.json")
    full_path    = os.path.join(dir_json,    stem + ".json")

    if not os.path.exists(compact_path):
        raise FileNotFoundError(f"Compact file not found: {compact_path}")
    if not os.path.exists(full_path):
        raise FileNotFoundError(f"Full JSON not found: {full_path}")

    return merge_and_diff(compact_path, full_path, out_dir)


# ────────────────────────────────────────────────────────────────────────────
# CLI
# ────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(0)

    compact_arg  = sys.argv[1]
    full_arg     = sys.argv[2] if len(sys.argv) > 2 else None
    out_dir_arg  = sys.argv[3] if len(sys.argv) > 3 else "merged_json"

    try:
        path, total, changed = merge_and_diff(compact_arg, full_arg, out_dir_arg)
        sys.exit(0)
    except (FileNotFoundError, ValueError) as exc:
        print(_c(f"\n  ✗ Error: {exc}\n", "\033[91m"), file=sys.stderr)
        sys.exit(1)