"""
compact_json.py
───────────────
Extracts translatable run-text from a pipeline JSON (to_json/*.json)
into a compact format ready for AI translation, then reinserts
translated texts back into the full pipeline JSON.

KEY FEATURES
────────────

1. PARAGRAPH BOUNDARIES  (para_breaks)
   _collect_run_refs() now records which flat index each paragraph starts at.
   This list (para_breaks) is saved in the compact JSON and used by
   translate_compact.py to prevent word-groups from crossing paragraph
   boundaries — fixing the "words merged across lines" bug.

2. STRUCTURAL PRESERVATION  (skip + suffix migration)
   Runs containing only tabs/arrows/numbers are marked skip=True.
   Their original text is stored verbatim — they are never sent to the AI.
   On reinsertion, any structural suffix attached to a translated run
   is migrated to the following structural run so all tab characters share
   the same font metrics (preventing wrong tab widths in Arabic docs).

Compact format  →  compact/<basename>.compact.json
──────────────────────────────────────────────────
{
  "source":      "filename.json",
  "count":       42,
  "texts":       ["Name", "AHMED SAID...", "\t→\t→", ...],
  "prefixes":    ["", "", ...],
  "suffixes":    ["\t→", "", ...],
  "skip":        [false, false, true, ...],
  "para_breaks": [0, 1, 3, 8, ...]   ← NEW: flat indices where paragraphs start
}

For skip=True runs: texts[i] = original run text (NOT empty).
This ensures downstream merge tools restore structural runs correctly.
"""

import json
import os
import unicodedata


# ── Structural character set ──────────────────────────────────────────────────

_STRUCTURAL: frozenset = frozenset(
    "\t\n\r"
    "\u00a0"                    # non-breaking space
    "\u200f\u200e"              # RTL / LTR marks
    "\u202a\u202b\u202c"        # embedding / override marks
    "\u202d\u202e"
    "\ufeff"                    # BOM
    "\u2192\u2190\u2194"        # → ← ↔
    "\u2191\u2193"              # ↑ ↓
    "\u25ba\u25c4\u25b6\u25c0"  # ► ◄ ▶ ◀
)


def _has_letters(text: str) -> bool:
    return any(unicodedata.category(ch).startswith("L") for ch in text)


def _is_structural(ch: str) -> bool:
    return ch in _STRUCTURAL


def _split_run(text: str) -> tuple[str, str, str, bool]:
    """
    Split a run text into (prefix, core, suffix, skip).

    skip=True  → no letters found; caller stores original text in texts[i].
    skip=False → core is the translatable span bounded by structural chars.

    Examples
    --------
    "Name:"                   → ("", "Name:",    "",              False)
    "Name: \\t→\\t→\\t:"      → ("", "Name: ",   "\\t→\\t→\\t:", False)
    "(Barcode)\\t→"           → ("", "(Barcode)", "\\t→",         False)
    " → \\t :. → \\t"         → ("", "",          "",             True)
    "\\t"                     → ("", "",          "",             True)
    "2437154343"              → ("", "",          "",             True)  no letters
    "16/07/2026 A.D."         → ("", "16/07/2026 A.D.", "",      False)
    """
    if not _has_letters(text):
        return ("", "", "", True)

    first = next(i for i, ch in enumerate(text) if unicodedata.category(ch).startswith("L"))
    last  = max(i for i, ch in enumerate(text) if unicodedata.category(ch).startswith("L"))

    start = first
    while start > 0 and not _is_structural(text[start - 1]):
        start -= 1

    end = last
    while end < len(text) - 1 and not _is_structural(text[end + 1]):
        end += 1

    return (text[:start], text[start:end + 1], text[end + 1:], False)


# ── Internal walker ───────────────────────────────────────────────────────────

def _collect_run_refs(
    obj: object,
    out: list,
    para_breaks: list[int],
) -> None:
    """
    Recursively walk obj in document order and append every run-dict
    that contains a "text" key to out.

    Whenever a paragraph (dict with both "runs" and "type"=="paragraph"
    or "type"=="empty_paragraph") is entered, record the current flat
    index in para_breaks — this marks where that paragraph's runs start.

    Textbox paragraphs are also tracked so the AI never merges their
    content with the surrounding body text.
    """
    if isinstance(obj, dict):
        runs = obj.get("runs")
        obj_type = obj.get("type", "")

        # Paragraph or textbox paragraph — record a boundary
        if obj_type in ("paragraph", "empty_paragraph") and "runs" in obj:
            para_breaks.append(len(out))

        if isinstance(runs, list):
            for run in runs:
                if isinstance(run, dict) and "text" in run:
                    out.append(run)

        for key, value in obj.items():
            if key == "runs":
                continue
            _collect_run_refs(value, out, para_breaks)

    elif isinstance(obj, list):
        for item in obj:
            _collect_run_refs(item, out, para_breaks)


# ── Public API ────────────────────────────────────────────────────────────────

def extract_compact(json_path: str, out_dir: str = "compact") -> tuple[str, int]:
    """
    Read a pipeline JSON, split run texts, and save a compact JSON.

    The compact JSON includes para_breaks: a list of flat run indices
    where each paragraph starts. translate_compact uses this to prevent
    word groups from spanning paragraph boundaries.

    Returns (out_path, total_run_count).
    """
    os.makedirs(out_dir, exist_ok=True)
    basename = os.path.splitext(os.path.basename(json_path))[0]
    out_path = os.path.join(out_dir, basename + ".compact.json")

    with open(json_path, "r", encoding="utf-8") as fh:
        data = json.load(fh)

    run_refs:    list[dict] = []
    para_breaks: list[int]  = []
    _collect_run_refs(data, run_refs, para_breaks)

    # Deduplicate and sort para_breaks (the walker may visit a paragraph node
    # before its runs are appended, so values are correct but may repeat if
    # the same index is recorded twice for empty paras)
    para_breaks = sorted(set(para_breaks))

    texts:    list[str]  = []
    prefixes: list[str]  = []
    suffixes: list[str]  = []
    skip:     list[bool] = []

    for r in run_refs:
        original = r["text"]
        prefix, core, suffix, is_skip = _split_run(original)

        if is_skip:
            texts.append(original)   # original text — not empty string
            prefixes.append("")
            suffixes.append("")
        else:
            texts.append(core)
            prefixes.append(prefix)
            suffixes.append(suffix)

        skip.append(is_skip)

    n_translatable = sum(1 for s in skip if not s)

    compact = {
        "source":      os.path.basename(json_path),
        "count":       len(run_refs),
        "texts":       texts,
        "prefixes":    prefixes,
        "suffixes":    suffixes,
        "skip":        skip,
        "para_breaks": para_breaks,
    }

    with open(out_path, "w", encoding="utf-8") as fh:
        json.dump(compact, fh, ensure_ascii=False, indent=2)

    print(
        f"[compact_json] {len(run_refs)} runs  "
        f"({n_translatable} translatable, "
        f"{len(run_refs) - n_translatable} structural/numeric skipped)  "
        f"{len(para_breaks)} paragraphs"
        f"  →  {out_path}"
    )
    return out_path, len(run_refs)


def reinsert_compact(
    json_path:    str,
    compact_path: str,
    out_dir:      str = "translated_json",
) -> str:
    """
    Merge a (translated) compact JSON back into the full pipeline JSON.

    Reinsertion rules
    -----------------
    skip=True  → run["text"] = texts[i]  (original structural text, verbatim)

    skip=False → normally: run["text"] = prefix + translated_core + suffix
                 BUT if suffix is non-empty AND the next run is skip=True,
                 the suffix is MIGRATED to the start of the next run instead:

                   run[i]["text"]   = prefix + translated_core      (no suffix)
                   run[i+1]["text"] = suffix + run[i+1]["text"]     (suffix prepended)

                 This ensures all tab characters use the structural run's font
                 metrics, preventing wrong tab widths after Arabic text.

    paragraph["text"] is regenerated from the joined run texts.

    Returns the path to the merged JSON file.
    """
    os.makedirs(out_dir, exist_ok=True)
    basename = os.path.splitext(os.path.basename(json_path))[0]
    out_path = os.path.join(out_dir, basename + ".json")

    with open(json_path, "r", encoding="utf-8") as fh:
        data = json.load(fh)
    with open(compact_path, "r", encoding="utf-8") as fh:
        compact = json.load(fh)

    translated: list[str]  = compact["texts"]
    prefixes:   list[str]  = compact.get("prefixes", [""] * len(translated))
    suffixes:   list[str]  = compact.get("suffixes", [""] * len(translated))
    skip:       list[bool] = compact.get("skip",     [False] * len(translated))

    # Legacy support
    if "pure_ws" in compact and "skip" not in compact:
        skip = compact["pure_ws"]

    run_refs:    list[dict] = []
    para_breaks: list[int]  = []
    _collect_run_refs(data, run_refs, para_breaks)

    if len(run_refs) != len(translated):
        raise ValueError(
            f"Run count mismatch: JSON has {len(run_refs)} runs, "
            f"compact has {len(translated)} entries.\n"
            f"Make sure you used the compact file generated from this exact JSON."
        )

    n = len(run_refs)

    # First pass: write all texts (structural runs get their original text back)
    for i, (run, stored_text) in enumerate(zip(run_refs, translated)):
        if skip[i]:
            run["text"] = stored_text   # original structural text
        else:
            run["text"] = prefixes[i] + stored_text + suffixes[i]

    # Second pass: migrate suffixes to the next structural run
    for i in range(n - 1):
        if skip[i]:
            continue
        suffix = suffixes[i]
        if not suffix:
            continue
        if skip[i + 1]:
            run_refs[i]["text"]     = prefixes[i] + translated[i]
            run_refs[i + 1]["text"] = suffix + run_refs[i + 1]["text"]

    # Regenerate paragraph.text from runs
    def _regen(obj: object) -> None:
        if isinstance(obj, dict):
            if "runs" in obj and "text" in obj and isinstance(obj["runs"], list):
                obj["text"] = "".join(r.get("text", "") for r in obj["runs"])
            for v in obj.values():
                _regen(v)
        elif isinstance(obj, list):
            for item in obj:
                _regen(item)

    _regen(data)

    with open(out_path, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)

    return out_path


def count_compact_chars(compact_path: str) -> int | None:
    """Return total character count of translatable (non-skip) texts."""
    try:
        with open(compact_path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        skip = data.get("skip", data.get("pure_ws", [False] * len(data.get("texts", []))))
        return sum(len(t) for t, s in zip(data.get("texts", []), skip) if not s)
    except Exception:
        return None


# ── CLI ───────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage:")
        print("  Extract:  python compact_json.py extract <file.json> [out_dir]")
        print("  Reinsert: python compact_json.py reinsert <file.json> <compact.json> [out_dir]")
        sys.exit(1)

    cmd = sys.argv[1].lower()

    if cmd == "extract":
        src     = sys.argv[2]
        out_dir = sys.argv[3] if len(sys.argv) > 3 else "compact"
        path, n = extract_compact(src, out_dir)
        print(f"Extracted {n} run texts → {path}")

    elif cmd == "reinsert":
        if len(sys.argv) < 4:
            print("reinsert needs <file.json> <compact.json> [out_dir]")
            sys.exit(1)
        src         = sys.argv[2]
        compact_src = sys.argv[3]
        out_dir     = sys.argv[4] if len(sys.argv) > 4 else "translated_json"
        path = reinsert_compact(src, compact_src, out_dir)
        print(f"Reinserted → {path}")

    else:
        print(f"Unknown command: {cmd}")
        sys.exit(1)