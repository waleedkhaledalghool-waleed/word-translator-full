"""
translate_compact.py
────────────────────
Translates the "texts" list inside a compact JSON file in-place,
using the Google Gemini API.

PARAGRAPH-AWARE GROUP MERGING
──────────────────────────────
Word processors (Word) often split a single visible word across multiple
runs due to internal formatting. For example:

    "Place of " | "I" | "ssue"  →  should be "Place of Issue"
    "Employer " | "Identity" | " No. "  →  "Employer Identity No."

The fix: consecutive word-containing runs are merged into one translation
unit — but ONLY within the same paragraph. The compact JSON's "para_breaks"
field records the flat run index where each paragraph starts, and group
building flushes at every paragraph boundary.

This means:
  - The AI always sees complete phrases, never isolated word fragments
  - Groups never cross paragraph or textbox boundaries
  - One translated string is written back to run[0] of each group;
    remaining runs in the group are cleared to "" so Word concatenates
    them into a single correct word

STRUCTURAL runs (skip=True: tabs, arrows, pure numbers) are never
grouped or sent to the translator — they pass through untouched.

SESSION-BASED TRANSLATION
──────────────────────────
One chat session per document. The model sees all previous batches
in its history and maintains consistent terminology automatically.

Requires: pip install google-genai python-dotenv
"""

import json
import os
import shutil
import time
import unicodedata
from datetime import datetime

from dotenv import load_dotenv
load_dotenv()

from google import genai

# ── Constants ─────────────────────────────────────────────────────────────────

_CHARS_PER_PAGE    = 1_500
_INTER_BATCH_DELAY = 0.5
_MAX_GLOSSARY      = 12

_LANG_NAMES: dict[str, str] = {
    "en": "English", "ar": "Arabic",  "fr": "French",  "de": "German",
    "es": "Spanish", "it": "Italian", "ru": "Russian", "zh": "Chinese",
    "ja": "Japanese","tr": "Turkish", "fa": "Persian", "ur": "Urdu",
    "he": "Hebrew",  "auto": "auto-detected",
}

def _lang_name(code: str) -> str:
    return _LANG_NAMES.get(code.lower(), code)

def _has_letters(text: str) -> bool:
    return any(unicodedata.category(ch).startswith("L") for ch in text)


# ── Group builder ─────────────────────────────────────────────────────────────

def _build_groups(
    texts:       list[str],
    skip:        list[bool],
    para_breaks: list[int],
) -> list[list[int]]:
    """
    Group consecutive non-skip, word-containing run indices together,
    respecting paragraph boundaries.

    A group is flushed (closed) whenever:
      - A paragraph boundary (index in para_breaks) is reached
      - A skip=True run is encountered (structural: tab, arrow, etc.)
      - A non-skip run with NO letters is encountered

    Returns a list of groups, each group being a list of indices into texts[].
    """
    break_set = set(para_breaks)
    groups: list[list[int]] = []
    current_group: list[int] = []

    for i, (text, is_skip) in enumerate(zip(texts, skip)):
        # Paragraph boundary — always flush before processing this run
        if i in break_set and i > 0:
            if current_group:
                groups.append(current_group)
                current_group = []

        if is_skip:
            # Structural run — flush and skip
            if current_group:
                groups.append(current_group)
                current_group = []
        elif _has_letters(text):
            # Word content — add to current group
            current_group.append(i)
        else:
            # Non-skip but no letters (e.g. pure space or lone punctuation)
            if current_group:
                groups.append(current_group)
                current_group = []

    if current_group:
        groups.append(current_group)

    return groups


# ── Core function ─────────────────────────────────────────────────────────────

def translate_compact(
    compact_path : str,
    src_lang     : str        = "en",
    tgt_lang     : str        = "ar",
    batch_size   : int        = 40,
    model        : str        = "gemini-2.5-flash-lite",
    demo_pages   : int | None = None,
    doc_type     : str | None = None,
) -> None:
    """
    Translate texts in a compact JSON file in-place.

    Adjacent word-content runs within the same paragraph are merged into
    single translation units so the AI always sees complete phrases.
    Groups never cross paragraph or textbox boundaries.
    """

    with open(compact_path, "r", encoding="utf-8") as fh:
        compact = json.load(fh)

    texts:       list[str]  = compact["texts"]
    skip:        list[bool] = compact.get("skip", compact.get("pure_ws", [False] * len(texts)))
    para_breaks: list[int]  = compact.get("para_breaks", [])
    total = len(texts)

    if total == 0:
        print("[translate_compact] No texts to translate.")
        return

    backup_path = compact_path.replace(".compact.json", ".compact.orig.json")
    if not os.path.exists(backup_path):
        shutil.copy2(compact_path, backup_path)
        print(f"[translate_compact] Backup → {os.path.basename(backup_path)}")

    # ── Build word groups (paragraph-boundary-aware) ──────────────────────────
    groups = _build_groups(texts, skip, para_breaks)

    # Each group → one translation unit = joined text of all runs in group
    group_texts: list[str] = ["".join(texts[i] for i in g) for g in groups]

    n_structural = sum(1 for s in skip if s)
    print(
        f"[translate_compact] {total} runs → {len(groups)} word groups  "
        f"({n_structural} structural, {len(para_breaks)} paragraphs)"
    )

    # ── Demo mode cutoff ──────────────────────────────────────────────────────
    cutoff_count = len(groups)
    if demo_pages is not None and demo_pages > 0:
        char_limit = demo_pages * _CHARS_PER_PAGE
        cumulative = 0
        for k, gt in enumerate(group_texts):
            cumulative += len(gt)
            if cumulative >= char_limit:
                cutoff_count = k + 1
                break
        print(f"[translate_compact] Demo mode: first {cutoff_count} groups")

    groups_to_translate = groups[:cutoff_count]
    texts_to_translate  = group_texts[:cutoff_count]
    n = len(groups_to_translate)

    if n == 0:
        print("[translate_compact] Nothing to translate.")
        return

    # ── Setup ─────────────────────────────────────────────────────────────────
    client    = genai.Client()
    src_name  = _lang_name(src_lang)
    tgt_name  = _lang_name(tgt_lang)
    n_batches = (n + batch_size - 1) // batch_size

    print(
        f"[translate_compact] {src_name} → {tgt_name}  |  "
        f"model={model}  batch={batch_size}  groups={n}  batches={n_batches}"
    )

    # ── Open chat session ─────────────────────────────────────────────────────
    if doc_type:
        doc_line = (
            f"You are translating a {doc_type}. "
            f"Use formal, official language and standard government terminology."
        )
    else:
        doc_line = (
            "You are translating an official document. "
            "Use formal language and consistent terminology."
        )

    system_prompt = (
        f"{doc_line}\n\n"
        f"You will receive the document's text in batches of numbered strings.\n"
        f"Translate from {src_name} to {tgt_name}.\n\n"
        f"Rules (apply to every batch):\n"
        f"- Return EXACTLY the same number of lines as given, numbered the same way.\n"
        f"- Format: '1. translated text', '2. translated text', etc.\n"
        f"- Preserve empty strings as empty lines.\n"
        f"- Do NOT add explanations, notes, or extra lines.\n"
        f"- Keep numbers, codes, dates, and IDs unchanged.\n"
        f"- Each string is a complete phrase or label from a Word document.\n"
        f"  Translate it as a whole — you are seeing the full word context.\n"
        f"- Keep punctuation that is part of labels (colons, periods) unchanged.\n"
        f"- Preserve any leading or trailing spaces exactly as given.\n"
        f"- Be consistent: use the same translation for the same term throughout.\n"
        f"- You will remember all previous batches in this session.\n\n"
        f"Confirm you understand by replying 'Ready.'"
    )

    chat = client.chats.create(model=model)
    try:
        ack = chat.send_message(system_prompt)
        print(f"[translate_compact] Session opened — {ack.text.strip()[:60]}")
    except Exception as e:
        print(f"[translate_compact] Session setup failed: {e} — falling back to stateless")
        chat = None

    # ── Batch loop ────────────────────────────────────────────────────────────
    translated_groups: list[str] = []
    glossary: dict[str, str] = {}
    t_start = time.time()

    for batch_num in range(1, n_batches + 1):
        start = (batch_num - 1) * batch_size
        end   = min(start + batch_size, n)
        batch = texts_to_translate[start:end]

        batch_text_joined = " ".join(batch).lower()
        relevant_glossary = {
            src: tgt for src, tgt in glossary.items()
            if src.lower() in batch_text_joined
        }
        if len(relevant_glossary) > _MAX_GLOSSARY:
            relevant_glossary = dict(
                sorted(relevant_glossary.items(), key=lambda x: len(x[0]))[:_MAX_GLOSSARY]
            )

        result = _translate_batch(
            chat              = chat,
            client            = client,
            batch             = batch,
            src_name          = src_name,
            tgt_name          = tgt_name,
            model             = model,
            batch_num         = batch_num,
            total_batches     = n_batches,
            doc_type          = doc_type,
            relevant_glossary = relevant_glossary,
        )
        translated_groups.extend(result)

        for src_item, tgt_item in zip(batch, result):
            if src_item.strip() and tgt_item.strip() and src_item.strip() != tgt_item.strip():
                glossary[src_item.strip()] = tgt_item.strip()

        elapsed = time.time() - t_start
        gloss_note = f"  gloss={len(relevant_glossary)}" if relevant_glossary else ""
        print(
            f"  Batch {batch_num}/{n_batches}  "
            f"({batch_num / n_batches * 100:.0f}%)  "
            f"{len(batch)} groups{gloss_note}  {elapsed:.1f}s"
        )

        if batch_num < n_batches:
            time.sleep(_INTER_BATCH_DELAY)

    # ── Expand group translations back to per-run texts ───────────────────────
    # For each translated group:
    #   - Write full translation into run[0] of the group
    #   - Clear all other runs in the group to ""
    # Structural runs (skip=True) stay unchanged throughout.

    new_texts = list(texts)

    for group, translation in zip(groups_to_translate, translated_groups):
        if not group:
            continue
        new_texts[group[0]] = translation
        for idx in group[1:]:
            new_texts[idx] = ""

    if len(new_texts) != total:
        raise RuntimeError(f"Length mismatch: expected {total}, got {len(new_texts)}")

    compact["texts"]      = new_texts
    compact["translated"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    compact["src_lang"]   = src_lang
    compact["tgt_lang"]   = tgt_lang
    compact["demo_pages"] = demo_pages

    with open(compact_path, "w", encoding="utf-8") as fh:
        json.dump(compact, fh, ensure_ascii=False, indent=2)

    total_elapsed = time.time() - t_start
    print(
        f"[translate_compact] Done — {n} groups in {total_elapsed:.1f}s  "
        f"→ {os.path.basename(compact_path)}"
    )


# ── Batch helper ──────────────────────────────────────────────────────────────

def _translate_batch(
    chat              : object,
    client            : object,
    batch             : list[str],
    src_name          : str,
    tgt_name          : str,
    model             : str,
    batch_num         : int,
    total_batches     : int,
    doc_type          : str | None,
    relevant_glossary : dict[str, str],
) -> list[str]:
    n = len(batch)

    numbered_input = "\n".join(
        f"{i + 1}. {text}" if text.strip() else f"{i + 1}. "
        for i, text in enumerate(batch)
    )

    glossary_block = ""
    if relevant_glossary:
        lines = [f"  {src}  →  {tgt}" for src, tgt in relevant_glossary.items()]
        glossary_block = (
            "Reminder — use these established translations:\n"
            + "\n".join(lines)
            + "\n\n"
        )

    if chat is not None:
        message = (
            f"{glossary_block}"
            f"Batch {batch_num}/{total_batches} — translate these {n} items:\n\n"
            f"{numbered_input}"
        )
    else:
        if doc_type:
            doc_line = f"You are translating a {doc_type}. Use formal, official language."
        else:
            doc_line = "You are translating an official document. Use formal language."

        message = (
            f"{doc_line}\n\n"
            f"Translate the following {n} strings from {src_name} to {tgt_name}.\n\n"
            f"{glossary_block}"
            f"Rules:\n"
            f"- Return EXACTLY {n} lines numbered 1 to {n}.\n"
            f"- Format: '1. translated text', '2. translated text', etc.\n"
            f"- Keep numbers, codes, dates, and IDs unchanged.\n"
            f"- Preserve any leading or trailing spaces exactly.\n"
            f"- Keep label punctuation (colons, periods) unchanged.\n\n"
            f"Input:\n{numbered_input}\n\nOutput:"
        )

    for attempt in range(1, 4):
        try:
            if chat is not None:
                response = chat.send_message(message)
            else:
                response = client.models.generate_content(model=model, contents=message)
            return _parse_response(response.text.strip(), n, batch)
        except Exception as exc:
            print(f"  [batch {batch_num}/{total_batches}] attempt {attempt} failed: {exc}")
            if attempt < 3:
                time.sleep(2 ** attempt)
            else:
                print(f"  [batch {batch_num}/{total_batches}] returning originals as fallback")
                return list(batch)


def _parse_response(raw: str, expected: int, originals: list[str]) -> list[str]:
    result = [""] * expected

    for line in raw.splitlines():
        line = line.strip()
        if not line:
            continue
        dot = line.find(".")
        if dot < 1:
            continue
        num_str = line[:dot].strip()
        if not num_str.isdigit():
            continue
        idx = int(num_str) - 1
        if 0 <= idx < expected:
            result[idx] = line[dot + 1:].lstrip(" ")

    for i in range(expected):
        if result[i] == "" and originals[i].strip() != "":
            result[i] = originals[i]

    return result


# ── CLI ───────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python translate_compact.py <file.compact.json> [src] [tgt] [batch] [model] [doc_type]")
        sys.exit(1)

    _path     = sys.argv[1]
    _src      = sys.argv[2] if len(sys.argv) > 2 else "en"
    _tgt      = sys.argv[3] if len(sys.argv) > 3 else "ar"
    _batch    = int(sys.argv[4]) if len(sys.argv) > 4 else 40
    _model    = sys.argv[5] if len(sys.argv) > 5 else "gemini-2.5-flash-lite"
    _doc_type = sys.argv[6] if len(sys.argv) > 6 else None

    translate_compact(_path, _src, _tgt, _batch, _model, doc_type=_doc_type)