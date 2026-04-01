"""
pipeline_ui.py
--------------
Main workflow UI for the Word-fidelity pipeline.

Folders (auto-created on first run):
  Originals/      <- put original .docx files here
  to_json/        <- full intermediate JSON per doc
  mirrored_json/  <- direction-mirrored JSON (when LTR↔RTL switch detected)
  compact/        <- compact text-only JSON for translation
  merged_json/    <- translated text merged back into full JSON
  normalized_json/<- field paragraphs normalized for uniform colon alignment
  to_word/        <- final reconstructed .docx

Pipeline steps:
  1  Extract    Originals/<n>.docx     ->  to_json/<n>.json
  2  Mirror     to_json/<n>.json       ->  mirrored_json/<n>.json  (flips layout if LTR↔RTL)
     OR skip mirror and go straight to compact if same direction
  3  Compact    [mirrored_]json/<n>.json  ->  compact/<n>.compact.json
  4  Translate  compact/<n>.compact.json  (in-place, API call)
  5  Merge      compact + json          ->  merged_json/<n>.json
  5b Normalize  merged_json/<n>.json   ->  normalized_json/<n>.json
  6  Rebuild    normalized_json/<n>.json ->  to_word/<n>.docx

Demo mode (toggle):  only translates paragraphs from the first 2 pages worth of
  content — useful for a quick sanity-check without burning API tokens.

Requires:  pip install python-docx lxml google-genai python-dotenv
"""

import tkinter as tk
from tkinter import font as tkfont, messagebox, simpledialog
import os
import sys
import json
import threading
import subprocess
import time
from datetime import datetime
from normalize_json import normalize_document

# ---------- Palette -----------------------------------------------------------
BG      = "#0f1117"
PANEL   = "#181c27"
BORDER  = "#2a2f3d"
ACCENT  = "#4f8ef7"
ACCENT2 = "#7c3aed"
TEXT    = "#e8eaf0"
MUTED   = "#6b7280"
SUCCESS = "#22c55e"
WARN    = "#f59e0b"
DANGER  = "#ef4444"
CODE_BG = "#12151e"
CODE_FG = "#a8d8a0"
HEADER  = "#1e2235"

ROW_ODD  = "#14182a"
ROW_EVEN = "#181c27"
ROW_SEL  = "#1e2d4a"

COMPACT_COL  = "#06b6d4"   # cyan
XLATE_COL    = "#f472b6"   # pink
MERGE_COL    = "#a78bfa"   # violet
MIRROR_COL   = "#fb923c"   # orange
DEMO_COL     = "#facc15"   # yellow
NORM_COL     = "#34d399"   # emerald green

# ---------- Folder paths ------------------------------------------------------
BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
DIR_ORIG    = os.path.join(BASE_DIR, "Originals")
DIR_JSON    = os.path.join(BASE_DIR, "to_json")
DIR_COMPACT = os.path.join(BASE_DIR, "compact")
DIR_MERGED  = os.path.join(BASE_DIR, "merged_json")
DIR_MIRROR  = os.path.join(BASE_DIR, "mirrored_json")
DIR_NORM    = os.path.join(BASE_DIR, "normalized_json")
DIR_WORD    = os.path.join(BASE_DIR, "to_word")
DIR_RAW     = os.path.join(BASE_DIR, "to_word_raw")   # rebuilt from merged (no normalize)
INDEX_PATH  = os.path.join(DIR_JSON, "_index.json")

for d in (DIR_ORIG, DIR_JSON, DIR_COMPACT, DIR_MERGED, DIR_MIRROR, DIR_NORM, DIR_WORD, DIR_RAW):
    os.makedirs(d, exist_ok=True)

# ---------- Index helpers -----------------------------------------------------
def _load_index():
    if os.path.exists(INDEX_PATH):
        try:
            with open(INDEX_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def _save_index(idx):
    with open(INDEX_PATH, "w", encoding="utf-8") as f:
        json.dump(idx, f, indent=2, ensure_ascii=False)

def _now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def _stem(path):
    return os.path.splitext(os.path.basename(path))[0]

# ---------- Char-count helpers ------------------------------------------------
def _count_docx_chars(docx_path):
    try:
        import zipfile
        from lxml import etree
        NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        with zipfile.ZipFile(docx_path, "r") as zf:
            doc_xml = zf.read("word/document.xml")
        tree = etree.fromstring(doc_xml)
        return sum(len(t.text or "") for t in tree.iter(f"{{{NS_W}}}t"))
    except Exception:
        return None

def _count_json_chars(json_path):
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        total = 0
        def _walk(obj):
            nonlocal total
            if isinstance(obj, dict):
                if "text" in obj and isinstance(obj["text"], str):
                    total += len(obj["text"])
                for v in obj.values():
                    _walk(v)
            elif isinstance(obj, list):
                for item in obj:
                    _walk(item)
        _walk(data)
        return total
    except Exception:
        return None

def _count_compact_chars(compact_path):
    try:
        with open(compact_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return sum(len(t) for t in data.get("texts", []))
    except Exception:
        return None

def _fmt_chars(n):
    if n is None: return "?"
    if n >= 1000: return f"{n:,} chars"
    return f"{n} chars"

# ---------- File path helpers -------------------------------------------------
def _json_path(name):    return os.path.join(DIR_JSON,    _stem(name) + ".json")
def _compact_path(name): return os.path.join(DIR_COMPACT, _stem(name) + ".compact.json")
def _merged_path(name):  return os.path.join(DIR_MERGED,  _stem(name) + ".json")
def _mirror_path(name):  return os.path.join(DIR_MIRROR,  _stem(name) + ".json")
def _norm_path(name):    return os.path.join(DIR_NORM,    _stem(name) + ".json")
def _out_path(name):     return os.path.join(DIR_WORD,    _stem(name) + ".docx")
def _raw_path(name):     return os.path.join(DIR_RAW,     _stem(name) + ".docx")  # rebuilt without normalize
def _orig_path(name):    return os.path.join(DIR_ORIG,    name)

# ---------- File discovery ----------------------------------------------------
def _scan_originals():
    try:
        return sorted(
            f for f in os.listdir(DIR_ORIG)
            if f.lower().endswith(".docx") and not f.startswith("~")
        )
    except Exception:
        return []

def _status_for(name, idx):
    entry = idx.get(name, {})
    s1 = os.path.exists(_json_path(name))    and bool(entry.get("extracted"))
    s2 = os.path.exists(_mirror_path(name))  and bool(entry.get("mirrored"))
    s3 = os.path.exists(_compact_path(name)) and bool(entry.get("compacted"))
    s4 = bool(entry.get("translated"))
    s5 = os.path.exists(_merged_path(name))  and bool(entry.get("merged"))
    s5b= os.path.exists(_norm_path(name))    and bool(entry.get("normalized"))
    s6 = os.path.exists(_out_path(name))     and bool(entry.get("rebuilt"))
    return s1, s2, s3, s4, s5, s5b, s6

def _icon_for(flags):
    s1, s2, s3, s4, s5, s5b, s6 = flags
    if s6:  return ("*",    SUCCESS)
    if s5b: return ("5b/6", NORM_COL)
    if s5:  return ("5/6",  MERGE_COL)
    if s4:  return ("4/6",  XLATE_COL)
    if s3:  return ("3/6",  COMPACT_COL)
    if s2:  return ("2/6",  MIRROR_COL)
    if s1:  return ("1/6",  WARN)
    return ("o", MUTED)


# =============================================================================
class PipelineApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DocX Pipeline")
        self.geometry("1280x820")
        self.minsize(960, 640)
        self.configure(bg=BG)
        self._index      = _load_index()
        self._selected   = None
        self._job_thread = None
        self._xlate_src      = "en"
        self._xlate_tgt      = "ar"
        self._xlate_model    = "gemini-2.5-flash-lite"
        self._xlate_batch    = 40
        self._xlate_doc_type = ""
        self._build_fonts()
        self._build_ui()
        self._refresh_list()
        self._status("Ready - add .docx files to Originals/ and press Refresh")

    def _build_fonts(self):
        self.f_title = tkfont.Font(family="Segoe UI", size=11, weight="bold")
        self.f_bold  = tkfont.Font(family="Segoe UI", size=9,  weight="bold")
        self.f_label = tkfont.Font(family="Segoe UI", size=9)
        self.f_small = tkfont.Font(family="Segoe UI", size=8)
        self.f_mono  = tkfont.Font(family="Consolas",  size=9)
        self.f_btn   = tkfont.Font(family="Segoe UI",  size=9,  weight="bold")
        self.f_hdr   = tkfont.Font(family="Segoe UI",  size=8,  weight="bold")

    # ======================================================================
    #  UI BUILD
    # ======================================================================
    def _build_ui(self):
        # Top bar
        top = tk.Frame(self, bg=HEADER, height=52)
        top.pack(fill="x")
        top.pack_propagate(False)

        tl = tk.Frame(top, bg=HEADER)
        tl.pack(side="left", padx=16, fill="y")
        tk.Label(tl, text="*", fg=ACCENT,  bg=HEADER, font=self.f_title).pack(side="left")
        tk.Label(tl, text="*", fg=ACCENT2, bg=HEADER, font=self.f_title).pack(side="left", padx=(2,8))
        tk.Label(tl, text="DocX Pipeline", fg=TEXT, bg=HEADER, font=self.f_title).pack(side="left")

        tr = tk.Frame(top, bg=HEADER)
        tr.pack(side="right", padx=10, pady=10)
        self._btn(tr, "Refresh",     self._refresh_list,    bg=PANEL,     fg=TEXT        ).pack(side="left", padx=(0,4))
        self._btn(tr, "Originals",   self._open_originals,  bg=PANEL,     fg="#818cf8"   ).pack(side="left", padx=(0,4))
        self._btn(tr, "Output",      self._open_output,     bg=PANEL,     fg="#818cf8"   ).pack(side="left", padx=(0,12))
        self._btn(tr, "1 Extract",   self._run_extract,     bg="#1a2a1a", fg=SUCCESS     ).pack(side="left", padx=(0,3))
        self._btn(tr, "2 Mirror",    self._run_mirror,      bg="#1a1408", fg=MIRROR_COL  ).pack(side="left", padx=(0,3))
        self._btn(tr, "3 Compact",   self._run_compact,     bg="#0d1e22", fg=COMPACT_COL ).pack(side="left", padx=(0,3))
        self._btn(tr, "4 Translate", self._run_translate,   bg="#22101a", fg=XLATE_COL   ).pack(side="left", padx=(0,3))
        self._btn(tr, "5 Merge",     self._run_merge,       bg="#1a1228", fg=MERGE_COL   ).pack(side="left", padx=(0,3))
        self._btn(tr, "5b Normalize",self._run_normalize,   bg="#0d1f18", fg=NORM_COL    ).pack(side="left", padx=(0,3))
        self._btn(tr, "6 Rebuild",   self._run_rebuild,     bg="#1a1a2e", fg=ACCENT      ).pack(side="left", padx=(0,3))
        self._btn(tr, "6r Raw",      self._run_rebuild_raw, bg="#12121e", fg="#6366f1"   ).pack(side="left", padx=(0,3))
        self._btn(tr, "Full Run",    self._run_full,        bg="#251a10", fg=WARN        ).pack(side="left", padx=(0,3))
        self._btn(tr, "Run ALL",     self._run_all,         bg="#2a1010", fg=DANGER      ).pack(side="left")

        # Main pane
        pane = tk.PanedWindow(self, orient="horizontal",
                              bg=BORDER, sashwidth=5, sashrelief="flat", handlesize=0)
        pane.pack(fill="both", expand=True, padx=10, pady=(8,0))

        # Left: file list
        left = tk.Frame(pane, bg=PANEL)
        pane.add(left, minsize=340, stretch="never")

        hdr = tk.Frame(left, bg=HEADER, height=28)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text=" *",        fg=MUTED, bg=HEADER, font=self.f_hdr, width=4,  anchor="w").pack(side="left")
        tk.Label(hdr, text="File",      fg=TEXT,  bg=HEADER, font=self.f_hdr,            anchor="w").pack(side="left", fill="x", expand=True)
        tk.Label(hdr, text="Extracted", fg=MUTED, bg=HEADER, font=self.f_hdr, width=17, anchor="e").pack(side="right", padx=(0,6))

        scroll_f = tk.Frame(left, bg=PANEL)
        scroll_f.pack(fill="both", expand=True)
        vsb = tk.Scrollbar(scroll_f, orient="vertical", bg=BORDER,
                           troughcolor=PANEL, width=9, bd=0, relief="flat")
        vsb.pack(side="right", fill="y")
        self.file_canvas = tk.Canvas(scroll_f, bg=PANEL, yscrollcommand=vsb.set,
                                     highlightthickness=0, bd=0)
        self.file_canvas.pack(side="left", fill="both", expand=True)
        vsb.config(command=self.file_canvas.yview)
        self.file_frame = tk.Frame(self.file_canvas, bg=PANEL)
        self._cw = self.file_canvas.create_window((0,0), window=self.file_frame, anchor="nw")
        self.file_frame.bind("<Configure>",
            lambda e: self.file_canvas.configure(scrollregion=self.file_canvas.bbox("all")))
        self.file_canvas.bind("<Configure>",
            lambda e: self.file_canvas.itemconfig(self._cw, width=e.width))
        self.file_frame.bind("<MouseWheel>",
            lambda e: self.file_canvas.yview_scroll(-1*(e.delta//120), "units"))

        # Right: detail + log
        right = tk.Frame(pane, bg=BG)
        pane.add(right, minsize=460, stretch="always")

        self.detail_frame = tk.Frame(right, bg=PANEL,
                                     highlightthickness=1, highlightbackground=BORDER)
        self.detail_frame.pack(fill="x", pady=(0,8))
        self._build_detail_card()

        log_hdr = tk.Frame(right, bg=HEADER, height=26)
        log_hdr.pack(fill="x")
        log_hdr.pack_propagate(False)
        tk.Label(log_hdr, text="  Log", fg=MUTED, bg=HEADER,
                 font=self.f_hdr, anchor="w").pack(side="left", fill="y")
        self._btn(log_hdr, "Clear", self._clear_log, bg=HEADER, fg=MUTED).pack(side="right", padx=6, pady=2)

        log_body = tk.Frame(right, bg=CODE_BG)
        log_body.pack(fill="both", expand=True)
        lvsb = tk.Scrollbar(log_body, orient="vertical", bg=BORDER,
                            troughcolor=CODE_BG, width=9, bd=0, relief="flat")
        lvsb.pack(side="right", fill="y")
        self.log_txt = tk.Text(log_body, bg=CODE_BG, fg=CODE_FG, font=self.f_mono,
            state="disabled", relief="flat", bd=0, padx=12, pady=10, wrap="word",
            yscrollcommand=lvsb.set, insertbackground=ACCENT)
        lvsb.config(command=self.log_txt.yview)
        self.log_txt.pack(fill="both", expand=True)
        for tag, col in [("ok",SUCCESS),("warn",WARN),("err",DANGER),("info",ACCENT),
                         ("compact",COMPACT_COL),("xlate",XLATE_COL),
                         ("merge",MERGE_COL),("norm",NORM_COL),("muted",MUTED),("demo",DEMO_COL)]:
            self.log_txt.tag_config(tag, foreground=col)

        # Status bar
        self.status_var = tk.StringVar()
        sb = tk.Frame(self, bg=HEADER, height=26)
        sb.pack(fill="x", side="bottom")
        sb.pack_propagate(False)
        self.status_lbl = tk.Label(sb, textvariable=self.status_var,
            fg=MUTED, bg=HEADER, font=self.f_small, anchor="w")
        self.status_lbl.pack(side="left", padx=14, fill="y")
        hints = tk.Frame(sb, bg=HEADER)
        hints.pack(side="right", padx=10)
        for label, path in [("to_json",DIR_JSON),("compact",DIR_COMPACT),
                             ("merged",DIR_MERGED),("normalized",DIR_NORM),("to_word",DIR_WORD)]:
            tk.Label(hints, text=f"{label}: {path}",
                     fg="#374151", bg=HEADER, font=self.f_small).pack(side="left", padx=(0,10))

    # ----------------------------------------------------------  Detail card
    def _build_detail_card(self):
        inner = tk.Frame(self.detail_frame, bg=PANEL)
        inner.pack(fill="x", padx=14, pady=10)

        # Step flow
        step_row = tk.Frame(inner, bg=PANEL)
        step_row.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0,6))
        for label, color in [
            ("1 Extract",    SUCCESS),
            ("→", MUTED),
            ("2 Mirror",     MIRROR_COL),
            ("→", MUTED),
            ("3 Compact",    COMPACT_COL),
            ("→", MUTED),
            ("4 Translate",  XLATE_COL),
            ("→", MUTED),
            ("5 Merge",      MERGE_COL),
            ("→", MUTED),
            ("5b Normalize", NORM_COL),
            ("→", MUTED),
            ("6 Rebuild",    ACCENT),
        ]:
            tk.Label(step_row, text=label, fg=color, bg=PANEL, font=self.f_bold).pack(side="left", padx=(0,2))

        tk.Label(inner, text="Selected File", fg=MUTED, bg=PANEL,
                 font=self.f_hdr, anchor="w").grid(row=1, column=0, sticky="w")
        self.lbl_name = tk.Label(inner, text="--", fg=TEXT, bg=PANEL, font=self.f_bold, anchor="w")
        self.lbl_name.grid(row=1, column=1, sticky="w", padx=(10,0))

        # Language direction row
        xlate_row = tk.Frame(inner, bg=PANEL)
        xlate_row.grid(row=2, column=0, columnspan=2, sticky="w", pady=(4,4))
        tk.Label(xlate_row, text="Direction:", fg=MIRROR_COL, bg=PANEL, font=self.f_label).pack(side="left")
        self.var_src      = tk.StringVar(value=self._xlate_src)
        self.var_tgt      = tk.StringVar(value=self._xlate_tgt)
        self.var_model    = tk.StringVar(value=self._xlate_model)
        self.var_batch    = tk.StringVar(value=str(self._xlate_batch))
        self.var_doc_type = tk.StringVar(value=self._xlate_doc_type)

        for label, var, width, choices in [
            ("from", self.var_src,  8,  ["en","ar","auto","fr","de","es","it","ru","zh","ja","fa","he","ur"]),
            ("to",   self.var_tgt,  8,  ["ar","en","fr","de","es","it","ru","zh","ja","tr","fa","ur","he"]),
        ]:
            tk.Label(xlate_row, text=f" {label}", fg=MUTED, bg=PANEL, font=self.f_small).pack(side="left")
            om = tk.OptionMenu(xlate_row, var, *choices)
            om.configure(bg=PANEL, fg=TEXT, font=self.f_small, relief="flat",
                         highlightthickness=0, bd=0, width=width,
                         activebackground=BORDER, activeforeground=TEXT)
            om["menu"].configure(bg=PANEL, fg=TEXT, font=self.f_small)
            om.pack(side="left", padx=(0,2))

        self.lbl_mirror_will = tk.Label(xlate_row, text="", fg=MIRROR_COL, bg=PANEL, font=self.f_small)
        self.lbl_mirror_will.pack(side="left", padx=(8,0))
        self.var_src.trace_add("write", lambda *_: self._update_mirror_indicator())
        self.var_tgt.trace_add("write", lambda *_: self._update_mirror_indicator())

        tk.Label(xlate_row, text="   model", fg=MUTED, bg=PANEL, font=self.f_small).pack(side="left")
        model_om = tk.OptionMenu(xlate_row, self.var_model,
                                 "gemini-2.5-flash-lite", "gemini-2.5-flash",
                                 "gemini-2.0-flash", "gemini-1.5-flash")
        model_om.configure(bg=PANEL, fg=XLATE_COL, font=self.f_small, relief="flat",
                           highlightthickness=0, bd=0, width=20,
                           activebackground=BORDER, activeforeground=TEXT)
        model_om["menu"].configure(bg=PANEL, fg=TEXT, font=self.f_small)
        model_om.pack(side="left", padx=(0,2))

        tk.Label(xlate_row, text=" batch", fg=MUTED, bg=PANEL, font=self.f_small).pack(side="left")
        batch_entry = tk.Entry(xlate_row, textvariable=self.var_batch, width=4,
                               bg=BORDER, fg=TEXT, font=self.f_small,
                               relief="flat", insertbackground=TEXT)
        batch_entry.pack(side="left")

        tk.Label(xlate_row, text=" doc", fg=MUTED, bg=PANEL, font=self.f_small).pack(side="left")
        doc_entry = tk.Entry(xlate_row, textvariable=self.var_doc_type, width=28,
                             bg=BORDER, fg=XLATE_COL, font=self.f_small,
                             relief="flat", insertbackground=TEXT)
        doc_entry.pack(side="left")

        # Demo mode toggle
        demo_row = tk.Frame(inner, bg=PANEL)
        demo_row.grid(row=3, column=0, columnspan=2, sticky="w", pady=(2, 4))

        self.var_demo = tk.BooleanVar(value=False)
        demo_cb = tk.Checkbutton(
            demo_row,
            text="Demo mode  (translate first 2 pages only — saves API tokens)",
            variable=self.var_demo,
            bg=PANEL, fg=DEMO_COL,
            selectcolor=BORDER,
            activebackground=PANEL, activeforeground=DEMO_COL,
            font=self.f_small,
            bd=0, relief="flat",
            command=self._on_demo_toggle,
        )
        demo_cb.pack(side="left")

        self.lbl_demo_hint = tk.Label(
            demo_row, text="", fg=DEMO_COL, bg=PANEL, font=self.f_small
        )
        self.lbl_demo_hint.pack(side="left", padx=(10, 0))

        # Info rows
        rows = [
            ("Step 1 - Extracted",    "lbl_extracted"),
            ("Step 2 - Mirrored",     "lbl_mirrored"),
            ("Step 3 - Compacted",    "lbl_compacted"),
            ("Step 4 - Translated",   "lbl_translated"),
            ("Step 5 - Merged",       "lbl_merged"),
            ("Step 5b - Normalized",  "lbl_normalized"),
            ("Step 6 - Rebuilt",      "lbl_rebuilt"),
            ("Status",                "lbl_stat"),
            ("Orig chars",            "lbl_orig_chars"),
            ("JSON chars",            "lbl_json_chars"),
            ("Mirror",                "lbl_mirror_info"),
            ("Output chars",          "lbl_out_chars"),
        ]
        for i, (label, attr) in enumerate(rows, start=4):
            tk.Label(inner, text=label, fg=MUTED, bg=PANEL,
                     font=self.f_label, anchor="w").grid(row=i, column=0, sticky="w", pady=1)
            lbl = tk.Label(inner, text="--", fg=MUTED, bg=PANEL, font=self.f_label, anchor="w")
            lbl.grid(row=i, column=1, sticky="w", padx=(10,0))
            setattr(self, attr, lbl)

        # Buttons row 1 — pipeline actions
        br1 = tk.Frame(inner, bg=PANEL)
        br1.grid(row=17, column=0, columnspan=2, pady=(8,0), sticky="w")
        self._btn(br1, "1 Extract",    lambda: self._run_extract(),   bg="#1a2a1a", fg=SUCCESS    ).pack(side="left", padx=(0,3))
        self._btn(br1, "2 Mirror",     lambda: self._run_mirror(),    bg="#1a1408", fg=MIRROR_COL ).pack(side="left", padx=(0,3))
        self._btn(br1, "3 Compact",    lambda: self._run_compact(),   bg="#0d1e22", fg=COMPACT_COL).pack(side="left", padx=(0,3))
        self._btn(br1, "4 Translate",  lambda: self._run_translate(), bg="#22101a", fg=XLATE_COL  ).pack(side="left", padx=(0,3))
        self._btn(br1, "5 Merge",      lambda: self._run_merge(),     bg="#1a1228", fg=MERGE_COL  ).pack(side="left", padx=(0,3))
        self._btn(br1, "5b Normalize", lambda: self._run_normalize(), bg="#0d1f18", fg=NORM_COL   ).pack(side="left", padx=(0,3))
        self._btn(br1, "6 Rebuild",    lambda: self._run_rebuild(),   bg="#1a1a2e", fg=ACCENT     ).pack(side="left", padx=(0,3))
        self._btn(br1, "6r Raw",       lambda: self._run_rebuild_raw(),bg="#12121e",fg="#6366f1"  ).pack(side="left", padx=(0,3))
        self._btn(br1, "Full",         lambda: self._run_full(),      bg="#251a10", fg=WARN       ).pack(side="left", padx=(0,3))

        # Buttons row 2 — open / delete
        br2 = tk.Frame(inner, bg=PANEL)
        br2.grid(row=18, column=0, columnspan=2, pady=(4,0), sticky="w")
        self._btn(br2, "Open Original", lambda: self._open_original(),  bg=PANEL,     fg="#818cf8"  ).pack(side="left", padx=(0,3))
        self._btn(br2, "Open Result",   lambda: self._open_result(),    bg=PANEL,     fg="#818cf8"  ).pack(side="left", padx=(0,3))
        self._btn(br2, "Open raw/",     lambda: _open_folder(DIR_RAW),  bg=PANEL,     fg="#6366f1"  ).pack(side="left", padx=(0,3))
        self._btn(br2, "Open norm/",    lambda: _open_folder(DIR_NORM), bg=PANEL,     fg=NORM_COL  ).pack(side="left", padx=(0,3))
        self._btn(br2, "Open mirror/",  lambda: _open_folder(DIR_MIRROR), bg=PANEL,  fg=MIRROR_COL).pack(side="left", padx=(0,3))
        self._btn(br2, "Debug XML",     lambda: self._run_debug(),      bg="#1a1a2a", fg="#a78bfa" ).pack(side="left", padx=(0,3))

        # Buttons row 3 — delete
        br3 = tk.Frame(inner, bg=PANEL)
        br3.grid(row=19, column=0, columnspan=2, pady=(4,2), sticky="w")
        self._btn(br3, "Del JSON",      lambda: self._clean_json(),      bg="#2a1515", fg=DANGER).pack(side="left", padx=(0,2))
        self._btn(br3, "Del Mirror",    lambda: self._clean_mirror(),    bg="#2a1515", fg=DANGER).pack(side="left", padx=(0,2))
        self._btn(br3, "Del Compact",   lambda: self._clean_compact(),   bg="#2a1515", fg=DANGER).pack(side="left", padx=(0,2))
        self._btn(br3, "Del Merged",    lambda: self._clean_merged(),    bg="#2a1515", fg=DANGER).pack(side="left", padx=(0,2))
        self._btn(br3, "Del Normalize", lambda: self._clean_normalize(), bg="#2a1515", fg=DANGER).pack(side="left", padx=(0,2))
        self._btn(br3, "Del DOCX",      lambda: self._clean_docx(),      bg="#2a1515", fg=DANGER).pack(side="left", padx=(0,2))
        self._btn(br3, "Del All",       lambda: self._clean_all(),       bg="#3a0a0a", fg=DANGER).pack(side="left")

    def _on_demo_toggle(self):
        if self.var_demo.get():
            self.lbl_demo_hint.configure(text="⚠ Only first ~2 pages will be translated")
        else:
            self.lbl_demo_hint.configure(text="")

    def _update_mirror_indicator(self):
        try:
            from mirror_json import needs_mirror
        except ImportError:
            return
        src = self.var_src.get()
        tgt = self.var_tgt.get()
        will = needs_mirror(src, tgt)
        if will:
            self.lbl_mirror_will.configure(
                text=f"⟺  MIRROR ON  ({src} ↔ {tgt})", fg=MIRROR_COL)
        else:
            self.lbl_mirror_will.configure(
                text=f"→  passthrough  (same direction)", fg=MUTED)

    # ----------------------------------------------------------  File list
    def _refresh_list(self):
        for w in self.file_frame.winfo_children():
            w.destroy()
        files = _scan_originals()
        if not files:
            tk.Label(self.file_frame, text="No .docx files found in Originals/",
                     fg=MUTED, bg=PANEL, font=self.f_label, pady=30).pack(fill="x")
            return
        for i, name in enumerate(files):
            bg  = ROW_EVEN if i % 2 == 0 else ROW_ODD
            row = tk.Frame(self.file_frame, bg=bg, cursor="hand2")
            row.pack(fill="x")
            icon, col = _icon_for(_status_for(name, self._index))
            tk.Label(row, text=icon, fg=col, bg=bg, font=self.f_bold,
                     width=5, anchor="center").pack(side="left", pady=6)
            disp = name if len(name) <= 34 else name[:31] + "..."
            tk.Label(row, text=disp, fg=TEXT, bg=bg, font=self.f_label,
                     anchor="w").pack(side="left", fill="x", expand=True)
            ext_t = (self._index.get(name,{}).get("extracted") or "--")[:16]
            tk.Label(row, text=ext_t, fg=MUTED, bg=bg, font=self.f_small,
                     anchor="e", width=15).pack(side="right", padx=(0,8))
            for w2 in [row] + list(row.winfo_children()):
                w2.bind("<Button-1>", lambda e, n=name: self._select(n))

    def _select(self, name):
        self._selected = name
        self._update_detail()
        for w in self.file_frame.winfo_children():
            if not isinstance(w, tk.Frame): continue
            labels = [c for c in w.winfo_children() if isinstance(c, tk.Label)]
            is_sel = any(l.cget("text").replace("...","") in name or
                         name[:10] in l.cget("text") for l in labels)
            nb = ROW_SEL if is_sel else w.cget("bg")
            w.configure(bg=nb)
            for c in w.winfo_children(): c.configure(bg=nb)

    def _update_detail(self):
        name = self._selected
        if not name: return
        self.lbl_name.configure(text=name, fg=ACCENT)
        entry                      = self._index.get(name, {})
        s1,s2,s3,s4,s5,s5b,s6     = _status_for(name, self._index)

        self.lbl_extracted.configure( text=entry.get("extracted",   "Not yet"), fg=SUCCESS    if s1  else MUTED)
        self.lbl_mirrored.configure(  text=entry.get("mirrored",    "Not yet"), fg=MIRROR_COL if s2  else MUTED)
        self.lbl_compacted.configure( text=entry.get("compacted",   "Not yet"), fg=COMPACT_COL if s3 else MUTED)
        self.lbl_translated.configure(text=entry.get("translated",  "Not yet"), fg=XLATE_COL  if s4  else MUTED)
        self.lbl_merged.configure(    text=entry.get("merged",      "Not yet"), fg=MERGE_COL  if s5  else MUTED)
        self.lbl_normalized.configure(text=entry.get("normalized",  "Not yet"), fg=NORM_COL   if s5b else MUTED)
        self.lbl_rebuilt.configure(   text=entry.get("rebuilt",     "Not yet"), fg=SUCCESS    if s6  else MUTED)

        if s6:   st,sc = "Complete",   SUCCESS
        elif s5b:st,sc = "Normalized", NORM_COL
        elif s5: st,sc = "Merged",     MERGE_COL
        elif s4: st,sc = "Translated", XLATE_COL
        elif s3: st,sc = "Compacted",  COMPACT_COL
        elif s2: st,sc = "Mirrored",   MIRROR_COL
        elif s1: st,sc = "Extracted",  WARN
        else:    st,sc = "Not started",MUTED
        self.lbl_stat.configure(text=st, fg=sc)

        def _cache(key, fn, path):
            val = entry.get(key)
            if val is None and os.path.exists(path):
                val = fn(path)
                if val is not None:
                    self._index.setdefault(name, {})[key] = val
                    _save_index(self._index)
            return val

        orig_n = _cache("orig_chars", _count_docx_chars, os.path.join(DIR_ORIG, name))
        json_n = _cache("json_chars", _count_json_chars, _json_path(name))
        out_n  = _cache("out_chars",  _count_docx_chars, _out_path(name))

        def _cl(n, ref=None):
            if n is None: return "--"
            s = _fmt_chars(n)
            if ref:
                d = n-ref; sg = "+" if d>=0 else ""
                s += f"  ({sg}{d:,}  {sg}{d/ref*100:.1f}%)"
            return s

        self.lbl_orig_chars.configure(text=_cl(orig_n), fg=TEXT if orig_n else MUTED)
        self.lbl_json_chars.configure(text=_cl(json_n, orig_n), fg=TEXT if json_n else MUTED)

        mirror_meta = entry.get("mirror_applied")
        if mirror_meta is not None:
            label = "⟺ mirrored (LTR↔RTL)" if mirror_meta else "→ passthrough (same dir)"
            self.lbl_mirror_info.configure(text=label, fg=MIRROR_COL if mirror_meta else MUTED)
        else:
            self.lbl_mirror_info.configure(text="--", fg=MUTED)

        self.lbl_out_chars.configure(
            text=_cl(out_n, orig_n),
            fg=(SUCCESS if out_n and orig_n and abs(out_n-orig_n)<=orig_n*0.02
                else WARN if out_n else MUTED))

    # -----------------------------------------------------------  Logging
    def _log(self, msg, tag="info"):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_txt.configure(state="normal")
        self.log_txt.insert("end", f"[{ts}] {msg}\n", tag)
        self.log_txt.configure(state="disabled")
        self.log_txt.see("end")

    def _clear_log(self):
        self.log_txt.configure(state="normal")
        self.log_txt.delete("1.0","end")
        self.log_txt.configure(state="disabled")

    def _status(self, msg, color=MUTED):
        self.status_var.set(f"  {msg}")
        self.status_lbl.configure(fg=color)

    def _btn(self, parent, text, cmd, bg=PANEL, fg=TEXT):
        return tk.Button(parent, text=text, command=cmd, bg=bg, fg=fg,
                         font=self.f_btn, relief="flat", bd=0, padx=9, pady=5,
                         activebackground=BORDER, activeforeground=TEXT, cursor="hand2")

    def _guard(self):
        if self._job_thread and self._job_thread.is_alive():
            self._status("A job is already running...", WARN); return False
        return True

    def _need_selection(self):
        if not self._selected:
            messagebox.showwarning("No file selected", "Please click a file in the list first.")
            return False
        return True

    def _run_extract(self):
        if not self._guard() or not self._need_selection(): return
        self._start_job(self._job_extract, [self._selected])

    def _run_mirror(self):
        if not self._guard() or not self._need_selection(): return
        self._start_job(self._job_mirror, [self._selected])

    def _run_compact(self):
        if not self._guard() or not self._need_selection(): return
        self._start_job(self._job_compact, [self._selected])

    def _run_translate(self):
        if not self._guard() or not self._need_selection(): return
        self._start_job(self._job_translate, [self._selected])

    def _run_merge(self):
        if not self._guard() or not self._need_selection(): return
        self._start_job(self._job_merge, [self._selected])

    def _run_normalize(self):
        if not self._guard() or not self._need_selection(): return
        self._start_job(self._job_normalize, [self._selected])

    def _run_rebuild(self):
        if not self._guard() or not self._need_selection(): return
        self._start_job(self._job_rebuild, [self._selected])

    def _run_rebuild_raw(self):
        if not self._guard() or not self._need_selection(): return
        self._start_job(self._job_rebuild_raw, [self._selected])

    def _run_full(self):
        if not self._guard() or not self._need_selection(): return
        self._start_job(self._job_full_pipeline, [self._selected])

    def _run_all(self):
        if not self._guard(): return
        files = _scan_originals()
        if not files:
            self._log("No files in Originals/", "warn"); return
        self._start_job(self._job_run_all, [files])

    def _start_job(self, fn, args):
        self._job_thread = threading.Thread(target=fn, args=args, daemon=True)
        self._job_thread.start()

    # =====================================================================
    #  JOBS
    # =====================================================================
    def _job_extract(self, name):
        src = os.path.join(DIR_ORIG, name)
        self._log(f"-- Step 1 Extract: {name}", "info")
        self._status(f"Extracting {name}...", ACCENT)
        try:
            orig_chars = _count_docx_chars(src)
            if orig_chars: self._log(f"  Original:  {_fmt_chars(orig_chars)}", "muted")
            sys.path.insert(0, BASE_DIR)
            from word_to_json import extract_to_file
            out = extract_to_file(src, DIR_JSON)
            sz = os.path.getsize(out)
            json_chars = _count_json_chars(out)
            self._index.setdefault(name, {})
            self._index[name]["extracted"] = _now()
            self._index[name]["json_size"] = sz
            if orig_chars:  self._index[name]["orig_chars"] = orig_chars
            if json_chars:  self._index[name]["json_chars"] = json_chars
            _save_index(self._index)
            if json_chars and orig_chars:
                d = json_chars - orig_chars; sg = "+" if d >= 0 else ""
                self._log(f"  JSON text: {_fmt_chars(json_chars)}  ({sg}{d:,}  {sg}{d/orig_chars*100:.1f}%)", "ok")
            self._log(f"  OK  to_json/{os.path.basename(out)}  ({sz//1024} KB)", "ok")
            self._status(f"Extracted {name}", SUCCESS)
        except Exception as e:
            self._log(f"  FAIL {e}", "err")
            self._status(f"Extract failed: {e}", DANGER)
        finally:
            self.after(0, self._refresh_list); self.after(0, self._update_detail)

    def _job_mirror(self, name):
        json_p = _json_path(name)
        if not os.path.exists(json_p):
            self._log(f"  FAIL No JSON for {name} - run Extract first", "warn"); return
        src = self.var_src.get()
        tgt = self.var_tgt.get()
        self._log(f"-- Step 2 Mirror: {name}  [{src} → {tgt}]", "info")
        self._status(f"Mirroring {name}...", MIRROR_COL)
        try:
            sys.path.insert(0, BASE_DIR)
            from mirror_json import mirror_document, needs_mirror
            will_mirror = needs_mirror(src, tgt)
            self._log(
                f"  {'⟺ Mirroring layout (LTR↔RTL direction flip)' if will_mirror else '→ Passthrough (same text direction, no flip)'}",
                "info" if will_mirror else "muted"
            )
            out_path, did_mirror = mirror_document(json_p, src, tgt, DIR_MIRROR)
            sz = os.path.getsize(out_path)
            self._index.setdefault(name, {})
            self._index[name]["mirrored"]       = _now()
            self._index[name]["mirror_applied"]  = did_mirror
            _save_index(self._index)
            self._log(f"  OK  mirrored_json/{os.path.basename(out_path)}  ({sz//1024} KB)", "ok")
            self._status(f"Mirror done: {name}  ({'flipped' if did_mirror else 'passthrough'})", MIRROR_COL)
        except Exception as e:
            self._log(f"  FAIL {e}", "err")
            import traceback; self._log(traceback.format_exc(), "err")
            self._status(f"Mirror failed: {e}", DANGER)
        finally:
            self.after(0, self._refresh_list); self.after(0, self._update_detail)

    def _job_compact(self, name):
        mirror_p = _mirror_path(name)
        json_p   = _json_path(name)
        if os.path.exists(mirror_p):
            src_json = mirror_p; label = "mirrored_json"
        elif os.path.exists(json_p):
            src_json = json_p;   label = "to_json"
        else:
            self._log(f"  FAIL No JSON for {name} - run Extract first", "warn"); return
        self._log(f"-- Step 3 Compact: {name}  (source: {label})", "compact")
        self._status(f"Compacting {name}...", COMPACT_COL)
        try:
            sys.path.insert(0, BASE_DIR)
            from compact_json import extract_compact
            out_path, run_count = extract_compact(src_json, DIR_COMPACT)
            sz = os.path.getsize(out_path)
            compact_chars = _count_compact_chars(out_path)
            self._index.setdefault(name, {})
            self._index[name]["compacted"]     = _now()
            self._index[name]["compact_count"] = run_count
            self._index[name]["compact_size"]  = sz
            self._index[name].pop("translated", None)
            if compact_chars: self._index[name]["compact_chars"] = compact_chars
            _save_index(self._index)
            orig_chars = self._index.get(name, {}).get("orig_chars")
            note = f"  (~{(1-compact_chars/orig_chars)*100:.0f}% smaller)" if compact_chars and orig_chars else ""
            self._log(f"  OK  {run_count} run texts  {_fmt_chars(compact_chars)}{note}", "compact")
            self._log(f"  OK  compact/{os.path.basename(out_path)}  ({sz//1024} KB)", "compact")
            self._status(f"Compacted {name}", COMPACT_COL)
        except Exception as e:
            self._log(f"  FAIL {e}", "err")
            import traceback; self._log(traceback.format_exc(), "err")
            self._status(f"Compact failed: {e}", DANGER)
        finally:
            self.after(0, self._refresh_list); self.after(0, self._update_detail)

    def _job_translate(self, name):
        compact_p = _compact_path(name)
        if not os.path.exists(compact_p):
            self._log(f"  FAIL No compact file for {name} - run Compact first", "warn"); return

        src      = self.var_src.get()
        tgt      = self.var_tgt.get()
        model    = self.var_model.get()
        demo     = self.var_demo.get()
        doc_type = self.var_doc_type.get().strip() or None
        try:
            batch = int(self.var_batch.get())
        except ValueError:
            batch = 40

        demo_label = "  [DEMO: first 2 pages only]" if demo else ""
        self._log(f"-- Step 4 Translate: {name}  [{src} -> {tgt}]  model={model}  batch={batch}{demo_label}", "xlate")
        if demo:
            self._log(f"  ⚠ Demo mode ON — only translating first ~2 pages of content", "demo")
        self._status(f"Translating {name}{'  (demo)' if demo else ''}...", XLATE_COL)
        try:
            sys.path.insert(0, BASE_DIR)
            from translate_compact import translate_compact

            translate_compact(
                compact_path = compact_p,
                src_lang     = src,
                tgt_lang     = tgt,
                batch_size   = batch,
                model        = model,
                demo_pages   = 2 if demo else None,
                doc_type     = doc_type,
            )

            compact_chars = _count_compact_chars(compact_p)
            self._index.setdefault(name, {})
            self._index[name]["translated"] = _now() + (" [demo]" if demo else "")
            if compact_chars: self._index[name]["compact_chars"] = compact_chars
            _save_index(self._index)

            self._log(f"  OK  compact/{os.path.basename(compact_p)} translated{demo_label}", "xlate")
            self._status(f"Translated {name}{demo_label}", XLATE_COL)
        except Exception as e:
            self._log(f"  FAIL {e}", "err")
            import traceback; self._log(traceback.format_exc(), "err")
            self._status(f"Translate failed: {e}", DANGER)
        finally:
            self.after(0, self._refresh_list); self.after(0, self._update_detail)

    def _job_merge(self, name):
        compact_p = _compact_path(name)
        mirror_p  = _mirror_path(name)
        json_p    = _json_path(name)
        if os.path.exists(mirror_p):
            base_json = mirror_p; base_label = "mirrored_json"
        elif os.path.exists(json_p):
            base_json = json_p;   base_label = "to_json"
        else:
            self._log(f"  FAIL No JSON for {name} - run Extract first", "warn"); return
        if not os.path.exists(compact_p):
            self._log(f"  FAIL No compact for {name} - run Compact first", "warn"); return
        self._log(f"-- Step 5 Merge: {name}  (base: {base_label})", "merge")
        self._status(f"Merging {name}...", MERGE_COL)
        try:
            sys.path.insert(0, BASE_DIR)
            from merge_compact import merge_and_diff
            out_path, total_runs, changed = merge_and_diff(compact_p, base_json, DIR_MERGED)
            sz = os.path.getsize(out_path)
            self._index.setdefault(name, {})
            self._index[name]["merged"]        = _now()
            self._index[name]["merge_changed"] = changed
            # Reset normalize flag when re-merged
            self._index[name].pop("normalized", None)
            _save_index(self._index)
            if changed == 0:
                self._log(f"  OK  No differences - texts identical to original", "muted")
            else:
                self._log(f"  OK  {changed} change{'s' if changed!=1 else ''} / {total_runs} runs  (see terminal for diff)", "merge")
            self._log(f"  OK  merged_json/{os.path.basename(out_path)}  ({sz//1024} KB)", "merge")
            self._status(f"Merged {name}  ({changed} change{'s' if changed!=1 else ''})", MERGE_COL)
        except Exception as e:
            self._log(f"  FAIL {e}", "err")
            import traceback; self._log(traceback.format_exc(), "err")
            self._status(f"Merge failed: {e}", DANGER)
        finally:
            self.after(0, self._refresh_list); self.after(0, self._update_detail)

    def _job_normalize(self, name):
        merged_p = _merged_path(name)
        if not os.path.exists(merged_p):
            self._log(f"  FAIL No merged JSON for {name} - run Merge first", "warn"); return
        out_path = _norm_path(name)
        self._log(f"-- Step 5b Normalize: {name}", "norm")
        self._status(f"Normalizing {name}...", NORM_COL)
        try:
            import shutil
            shutil.copy2(merged_p, out_path)          # start from merged copy
            _data = json.load(open(out_path, encoding="utf-8"))
            n_norm = normalize_document(_data)
            with open(out_path, "w", encoding="utf-8") as fh:
                json.dump(_data, fh, ensure_ascii=False, indent=2)
            sz = os.path.getsize(out_path)
            self._index.setdefault(name, {})
            self._index[name]["normalized"] = _now()
            _save_index(self._index)
            self._log(f"  OK  {n_norm} field paragraph{'s' if n_norm!=1 else ''} normalized → uniform colon alignment", "norm")
            self._log(f"  OK  normalized_json/{os.path.basename(out_path)}  ({sz//1024} KB)", "norm")
            self._status(f"Normalized {name}  ({n_norm} paragraphs)", NORM_COL)
        except Exception as e:
            self._log(f"  FAIL {e}", "err")
            import traceback; self._log(traceback.format_exc(), "err")
            self._status(f"Normalize failed: {e}", DANGER)
        finally:
            self.after(0, self._refresh_list); self.after(0, self._update_detail)

    def _job_rebuild(self, name):
        # Priority: normalized > merged > mirrored > raw json
        norm_p   = _norm_path(name)
        merged_p = _merged_path(name)
        mirror_p = _mirror_path(name)
        json_p   = _json_path(name)
        if os.path.exists(norm_p):
            src_json = norm_p;   label = "normalized_json"
        elif os.path.exists(merged_p):
            src_json = merged_p; label = "merged_json"
        elif os.path.exists(mirror_p):
            src_json = mirror_p; label = "mirrored_json"
        elif os.path.exists(json_p):
            src_json = json_p;   label = "to_json"
        else:
            self._log(f"  FAIL No JSON for {name} - run Extract first", "warn"); return
        self._log(f"-- Step 6 Rebuild: {name}  (source: {label})", "info")
        self._status(f"Rebuilding {name}...", ACCENT)
        try:
            orig_chars = self._index.get(name, {}).get("orig_chars")
            if orig_chars: self._log(f"  Original:  {_fmt_chars(orig_chars)}", "muted")
            sys.path.insert(0, BASE_DIR)
            from json_to_word import reconstruct_to_file
            out = reconstruct_to_file(src_json, DIR_WORD)
            sz  = os.path.getsize(out)
            out_chars = _count_docx_chars(out)
            self._index.setdefault(name, {})
            self._index[name]["rebuilt"]   = _now()
            self._index[name]["docx_size"] = sz
            if out_chars: self._index[name]["out_chars"] = out_chars
            _save_index(self._index)
            if out_chars and orig_chars:
                d = out_chars - orig_chars; sg = "+" if d >= 0 else ""
                ok = abs(d) <= orig_chars * 0.02
                self._log(f"  Output: {_fmt_chars(out_chars)}  ({sg}{d:,}  {sg}{d/orig_chars*100:.1f}%)  {'OK' if ok else 'Mismatch'}",
                          "ok" if ok else "warn")
            self._log(f"  OK  to_word/{os.path.basename(out)}  ({sz//1024} KB)", "ok")
            self._status(f"Rebuilt {name}", SUCCESS)
        except Exception as e:
            self._log(f"  FAIL {e}", "err")
            import traceback; self._log(traceback.format_exc(), "err")
            self._status(f"Rebuild failed: {e}", DANGER)
        finally:
            self.after(0, self._refresh_list); self.after(0, self._update_detail)

    def _job_rebuild_raw(self, name):
        """
        Rebuild directly from merged_json (skipping normalize) → to_word_raw/.
        Use this to compare normalized vs un-normalized output side-by-side.
        """
        merged_p = _merged_path(name)
        mirror_p = _mirror_path(name)
        json_p   = _json_path(name)
        if os.path.exists(merged_p):
            src_json = merged_p; label = "merged_json"
        elif os.path.exists(mirror_p):
            src_json = mirror_p; label = "mirrored_json"
        elif os.path.exists(json_p):
            src_json = json_p;   label = "to_json"
        else:
            self._log(f"  FAIL No JSON for {name} - run Extract first", "warn"); return
        self._log(f"-- Step 6r Raw Rebuild: {name}  (source: {label}, NO normalize)", "info")
        self._status(f"Raw rebuilding {name}...", "#6366f1")
        try:
            sys.path.insert(0, BASE_DIR)
            from json_to_word import reconstruct_to_file
            out = reconstruct_to_file(src_json, DIR_RAW)
            sz  = os.path.getsize(out)
            self._log(f"  OK  to_word_raw/{os.path.basename(out)}  ({sz//1024} KB)", "ok")
            self._log(f"  Compare with to_word/ to see normalize effect", "muted")
            self._status(f"Raw rebuilt {name}", SUCCESS)
            _open_folder(DIR_RAW)
        except Exception as e:
            self._log(f"  FAIL {e}", "err")
            import traceback; self._log(traceback.format_exc(), "err")
            self._status(f"Raw rebuild failed: {e}", DANGER)
        finally:
            self.after(0, self._refresh_list); self.after(0, self._update_detail)

    def _job_full_pipeline(self, name):
        """
        Full pipeline:
          1 Extract → 2 Mirror → 3 Compact → 4 Translate → 5 Merge → 5b Normalize → 6 Rebuild
        """
        demo = self.var_demo.get()
        src  = self.var_src.get()
        tgt  = self.var_tgt.get()
        demo_tag = "  [DEMO]" if demo else ""

        self._log(f"== Full pipeline: {name}  [{src}→{tgt}]{demo_tag} ==", "info")

        self._job_extract(name);    time.sleep(0.05)
        self._job_mirror(name);     time.sleep(0.05)
        self._job_compact(name);    time.sleep(0.05)
        self._job_translate(name);  time.sleep(0.05)
        self._job_merge(name);      time.sleep(0.05)
        self._job_normalize(name);  time.sleep(0.05)
        self._job_rebuild(name)

        self._log(f"== Done: {name}{demo_tag} ==", "ok")

    def _job_run_all(self, files):
        self._log(f"== Full pipeline: {len(files)} files ==", "info")
        for name in files:
            self._log("", "muted")
            self._job_full_pipeline(name)
        self._log("== All done ==", "ok")
        self._status(f"All {len(files)} files processed", SUCCESS)

    # -----------------------------------------------  Restore backup
    def _restore_xlate_backup(self):
        if not self._need_selection(): return
        name      = self._selected
        compact_p = _compact_path(name)
        backup_p  = compact_p.replace(".compact.json", ".compact.orig.json")
        if not os.path.exists(backup_p):
            self._log(f"  No backup found for {name}", "warn"); return
        import shutil
        shutil.copy2(backup_p, compact_p)
        self._index.get(name, {}).pop("translated", None)
        _save_index(self._index)
        self._log(f"  Restored backup -> {os.path.basename(compact_p)}", "xlate")
        self._status(f"Restored backup for {name}", XLATE_COL)
        self.after(0, self._refresh_list); self.after(0, self._update_detail)

    # -----------------------------------------------  Folder / file openers
    def _open_originals(self):
        _open_folder(DIR_ORIG); self._log(f"Opened: {DIR_ORIG}", "muted")

    def _open_output(self):
        _open_folder(DIR_WORD); self._log(f"Opened: {DIR_WORD}", "muted")

    def _open_result(self):
        if not self._selected: return
        path = _out_path(self._selected)
        if os.path.exists(path):
            _open_file(path); self._log(f"Opened result: {os.path.basename(path)}", "ok")
        else:
            self._log(f"No output yet for {self._selected}", "warn")

    def _open_original(self):
        if not self._selected: return
        path = _orig_path(self._selected)
        if os.path.exists(path):
            _open_file(path); self._log(f"Opened original: {os.path.basename(path)}", "ok")
        else:
            self._log(f"Original file not found: {path}", "warn")

    # -----------------------------------------------  Debug XML
    def _run_debug(self):
        if not self._need_selection() or not self._guard(): return
        self._start_job(self._job_debug, [self._selected])

    def _job_debug(self, name):
        src = os.path.join(DIR_ORIG, name)
        if not os.path.exists(src):
            self._log(f"  FAIL File not found: {src}", "err"); return
        self._log(f"-- Debug XML: {name} --", "info")
        try:
            import zipfile as zf_mod
            from lxml import etree
            NS_W   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            NS_WP  = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            NS_WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            NS_V   = "urn:schemas-microsoft-com:vml"
            with zf_mod.ZipFile(src, "r") as zf:
                doc_xml = zf.read("word/document.xml")
            tree = etree.fromstring(doc_xml)
            body = tree.find(f"{{{NS_W}}}body")

            def flat_text(el):
                parts = []
                for child in el:
                    try:    local = etree.QName(child.tag).localname
                    except: continue
                    if local in ("txbxContent","drawing","pict"): continue
                    if local == "r":
                        for t in child.findall(f"{{{NS_W}}}t"):  parts.append(t.text or "")
                        for _ in child.findall(f"{{{NS_W}}}tab"): parts.append("->")
                    else: parts.extend(flat_text(child))
                return parts

            pc = 0
            for idx, child in enumerate(body):
                try:    local = etree.QName(child.tag).localname
                except: local = "?"
                if local == "tbl":
                    self._log(f"  [{idx:03d}] <tbl>", "muted")
                    for ri, tr in enumerate(child.findall(f"{{{NS_W}}}tr")):
                        for ci, tc in enumerate(tr.findall(f"{{{NS_W}}}tc")):
                            tcPr = tc.find(f"{{{NS_W}}}tcPr"); cs = 1; vm = None
                            if tcPr is not None:
                                gs = tcPr.find(f"{{{NS_W}}}gridSpan")
                                if gs is not None: cs = int(gs.get(f"{{{NS_W}}}val",1))
                                v = tcPr.find(f"{{{NS_W}}}vMerge")
                                if v is not None:
                                    vm = "restart" if v.get(f"{{{NS_W}}}val")=="restart" else "continue"
                            ct = "".join(t.text or "" for t in tc.findall(f".//{{{NS_W}}}t"))[:60]
                            ex = []
                            if cs>1: ex.append(f"colspan={cs}")
                            if vm:   ex.append(f"vmerge={vm}")
                            estr = f"  [{', '.join(ex)}]" if ex else ""
                            self._log(f"      row{ri} col{ci}{estr}  text={repr(ct)}",
                                      "warn" if vm=="continue" or cs>1 else "muted")
                    continue
                if local != "p":
                    self._log(f"  [{idx:03d}] <{local}>", "muted"); continue
                pc += 1
                bt = "".join(flat_text(child))[:80]; fi = []
                for dr in child.findall(f"{{{NS_W}}}drawing")+child.findall(f".//{{{NS_W}}}drawing"):
                    if dr.find(f".//{{{NS_WPS}}}txbx") is not None:
                        fi.append(f"[wps-textbox: {repr(''.join(t.text or '' for t in dr.findall(f'.//{{{NS_W}}}t'))[:40])}]")
                    elif dr.find(f"{{{NS_WP}}}anchor") is not None:
                        fi.append("[floating-image]")
                for ps in child.findall(f".//{{{NS_V}}}shape"):
                    tb = [t.text or "" for t in ps.findall(f".//{{{NS_W}}}t")]
                    if tb: fi.append(f"[vml-textbox: {repr(''.join(tb)[:40])}]")
                col = "warn" if fi else ("ok" if bt.strip() else "muted")
                self._log(f"  [{idx:03d}] <p> text={repr(bt)}{'  '+'  '.join(fi) if fi else ''}", col)

            self._log(f"-- Done: {pc} paragraphs --", "ok")
            self._status(f"Debug complete", SUCCESS)
        except Exception as e:
            import traceback
            self._log(f"  FAIL {e}\n{traceback.format_exc()}", "err")

    # -----------------------------------------------  Clean helpers
    def _clean_json(self):
        if not self._need_selection(): return
        name = self._selected; p = _json_path(name)
        if os.path.exists(p): os.remove(p); self._log(f"  Removed: {os.path.basename(p)}", "warn")
        for k in ("extracted","json_size","json_chars"): self._index.get(name,{}).pop(k,None)
        _save_index(self._index); self._status(f"Cleaned JSON", WARN)
        self.after(0,self._refresh_list); self.after(0,self._update_detail)

    def _clean_mirror(self):
        if not self._need_selection(): return
        name = self._selected; p = _mirror_path(name)
        if os.path.exists(p): os.remove(p); self._log(f"  Removed: {os.path.basename(p)}", "warn")
        for k in ("mirrored","mirror_applied"): self._index.get(name,{}).pop(k,None)
        _save_index(self._index); self._status(f"Cleaned mirrored JSON", WARN)
        self.after(0,self._refresh_list); self.after(0,self._update_detail)

    def _clean_compact(self):
        if not self._need_selection(): return
        name = self._selected; p = _compact_path(name)
        bak = p.replace(".compact.json",".compact.orig.json")
        for fp in [p, bak]:
            if os.path.exists(fp): os.remove(fp); self._log(f"  Removed: {os.path.basename(fp)}", "warn")
        for k in ("compacted","compact_count","compact_size","compact_chars","translated"): self._index.get(name,{}).pop(k,None)
        _save_index(self._index); self._status(f"Cleaned compact", WARN)
        self.after(0,self._refresh_list); self.after(0,self._update_detail)

    def _clean_merged(self):
        if not self._need_selection(): return
        name = self._selected; p = _merged_path(name)
        if os.path.exists(p): os.remove(p); self._log(f"  Removed: {os.path.basename(p)}", "warn")
        for k in ("merged","merge_changed"): self._index.get(name,{}).pop(k,None)
        _save_index(self._index); self._status(f"Cleaned merged", WARN)
        self.after(0,self._refresh_list); self.after(0,self._update_detail)

    def _clean_normalize(self):
        if not self._need_selection(): return
        name = self._selected; p = _norm_path(name)
        if os.path.exists(p): os.remove(p); self._log(f"  Removed: {os.path.basename(p)}", "warn")
        for k in ("normalized",): self._index.get(name,{}).pop(k,None)
        _save_index(self._index); self._status(f"Cleaned normalized JSON", WARN)
        self.after(0,self._refresh_list); self.after(0,self._update_detail)

    def _clean_docx(self):
        if not self._need_selection(): return
        name = self._selected; p = _out_path(name)
        if os.path.exists(p): os.remove(p); self._log(f"  Removed: {os.path.basename(p)}", "warn")
        for k in ("rebuilt","docx_size","out_chars"): self._index.get(name,{}).pop(k,None)
        _save_index(self._index); self._status(f"Cleaned DOCX", WARN)
        self.after(0,self._refresh_list); self.after(0,self._update_detail)

    def _clean_all(self):
        if not self._need_selection(): return
        self._log(f"-- Cleaning all intermediates: {self._selected} --", "warn")
        self._clean_json(); self._clean_mirror(); self._clean_compact()
        self._clean_merged(); self._clean_normalize(); self._clean_docx()

    def _clean_all_compact(self):
        if not messagebox.askyesno(
            "Delete ALL compact files",
            "This will delete every .compact.json and .compact.orig.json in the compact/ folder "
            "for ALL documents, and reset translation status.\n\nContinue?",
        ):
            return
        removed = []
        try:
            for fname in os.listdir(DIR_COMPACT):
                if fname.endswith(".compact.json") or fname.endswith(".compact.orig.json"):
                    fp = os.path.join(DIR_COMPACT, fname)
                    os.remove(fp)
                    removed.append(fname)
        except Exception as e:
            self._log(f"  FAIL {e}", "err")
        for name in list(self._index.keys()):
            for k in ("compacted","compact_count","compact_size","compact_chars","translated"):
                self._index[name].pop(k, None)
        _save_index(self._index)
        if removed:
            self._log(f"-- Del ALL compact: removed {len(removed)} file(s) --", "warn")
            for f in removed: self._log(f"  Removed: {f}", "warn")
        else:
            self._log("  compact/ was already empty", "muted")
        self._status(f"Cleared compact/ ({len(removed)} files removed)", WARN)
        self.after(0, self._refresh_list)
        self.after(0, self._update_detail)


# ---------- OS helpers --------------------------------------------------------
def _open_folder(path):
    if sys.platform   == "win32":  os.startfile(path)
    elif sys.platform == "darwin": subprocess.run(["open",     path])
    else:                          subprocess.run(["xdg-open", path])

def _open_file(path):
    if sys.platform   == "win32":  os.startfile(path)
    elif sys.platform == "darwin": subprocess.run(["open",     path])
    else:                          subprocess.run(["xdg-open", path])


if __name__ == "__main__":
    app = PipelineApp()
    app.mainloop()