"""
word_to_html.py
───────────────
Extracts a Word (.docx) document into a rich JSON+HTML intermediate file
that captures layout, fonts, positioning, text boxes, tables, etc.
so it can be faithfully recreated by html_to_word.py

Output goes to:  to_html/<basename>.json
"""

import json
import os
import re
import zipfile
from copy import deepcopy

from lxml import etree

# ── XML namespaces ─────────────────────────────────────────────────────────────
NS = {
    "w":   "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp":  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a":   "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "mc":  "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "v":   "urn:schemas-microsoft-com:vml",
    "o":   "urn:schemas-microsoft-com:office:office",
    "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "ct":  "http://schemas.openxmlformats.org/package/2006/content-types",
}

def _q(ns, tag):
    return f"{{{NS[ns]}}}{tag}"

# ── Twips/EMU helpers ─────────────────────────────────────────────────────────
def _twips(val, default=0):
    """Parse a twips (DXA) integer safely."""
    try: return int(val)
    except: return default

def _emu_to_twips(emu):
    """EMU → twips (1 inch = 914400 EMU = 1440 twips)."""
    try: return int(int(emu) * 1440 / 914400)
    except: return 0

def _pt_half_to_pt(half):
    """Half-points → points."""
    try: return int(half) / 2
    except: return 0

# ── Color helpers ─────────────────────────────────────────────────────────────
def _norm_color(val):
    if not val or val.upper() in ("AUTO", "NONE", ""):
        return None
    val = val.strip().lstrip("#")
    if len(val) == 6:
        return "#" + val.upper()
    return None

# ── Low-level XML getters ─────────────────────────────────────────────────────
def _w(tag, el):
    return el.find(_q("w", tag))

def _wval(tag, el):
    ch = el.find(_q("w", tag))
    if ch is None: return None
    return ch.get(_q("w", "val"))

def _bool_prop(tag, rpr):
    """Return True/False/None for a w:tag that can be present (=True), absent (=None), or val=0 (=False)."""
    el = rpr.find(_q("w", tag)) if rpr is not None else None
    if el is None: return None
    val = el.get(_q("w","val"))
    if val in ("0","false"): return False
    return True

# ── Run formatting ─────────────────────────────────────────────────────────────
def _extract_rpr(rpr):
    """Extract run properties from a <w:rPr> element."""
    if rpr is None:
        return {}
    out = {}

    # Font — store Latin and Complex Script fonts separately.
    # In Arabic/mixed documents these are often different fonts.
    # Losing the CS font causes wrong glyph metrics → text clipping.
    rFonts = rpr.find(_q("w","rFonts"))
    if rFonts is not None:
        latin = rFonts.get(_q("w","ascii")) or rFonts.get(_q("w","hAnsi"))
        cs    = rFonts.get(_q("w","cs"))
        eastAsia = rFonts.get(_q("w","eastAsia"))
        if latin:    out["font"]         = latin
        if cs:       out["font_cs"]      = cs
        if eastAsia: out["font_eastAsia"]= eastAsia
        # Fallback: if only CS font is present (pure Arabic run), use it as font too
        if not latin and cs:
            out["font"] = cs

    # Size — store Latin size and Complex Script size separately.
    # A mismatch between sz and szCs causes text clipping in bidi paragraphs.
    sz = rpr.find(_q("w","sz"))
    if sz is not None:
        out["size"] = _pt_half_to_pt(sz.get(_q("w","val")))

    szCs = rpr.find(_q("w","szCs"))
    if szCs is not None:
        val_cs = _pt_half_to_pt(szCs.get(_q("w","val")))
        if "size" not in out:
            out["size"] = val_cs          # use as primary if no sz
        if val_cs != out.get("size"):
            out["size_cs"] = val_cs       # store separately only when different

    # Bold / Italic / Underline / Strike
    b  = _bool_prop("b",  rpr)
    i  = _bool_prop("i",  rpr)
    u  = rpr.find(_q("w","u"))
    s  = _bool_prop("strike", rpr)
    cs = _bool_prop("cs", rpr)   # complex script (Arabic etc)
    bCs = _bool_prop("bCs", rpr)
    iCs = _bool_prop("iCs", rpr)

    if b  is not None: out["bold"]      = b
    if bCs is True:    out["bold"]      = True
    if i  is not None: out["italic"]    = i
    if iCs is True:    out["italic"]    = True
    if s  is not None: out["strike"]    = s

    if u is not None:
        uval = u.get(_q("w","val"))
        out["underline"] = (uval not in ("none","",None))

    # Color
    color = rpr.find(_q("w","color"))
    if color is not None:
        out["color"] = _norm_color(color.get(_q("w","val")))

    # Highlight
    hl = rpr.find(_q("w","highlight"))
    if hl is not None:
        out["highlight"] = hl.get(_q("w","val"))

    # Character spacing (w:spacing) — tight/loose tracking.
    # Missing this causes glyphs to overflow their metric boxes → clipping.
    spacing_el = rpr.find(_q("w","spacing"))
    if spacing_el is not None:
        val = spacing_el.get(_q("w","val"))
        if val is not None:
            out["char_spacing"] = int(val)   # in twentieth-points (twips/20)

    # RTL
    rtl = _bool_prop("rtl", rpr)
    if rtl: out["rtl"] = True

    return out

# ── Paragraph formatting ───────────────────────────────────────────────────────
def _extract_ppr(ppr):
    """Extract paragraph properties from a <w:pPr> element."""
    if ppr is None:
        return {}
    out = {}

    # Alignment
    jc = ppr.find(_q("w","jc"))
    if jc is not None:
        out["align"] = jc.get(_q("w","val"), "left")

    # Indentation
    ind = ppr.find(_q("w","ind"))
    if ind is not None:
        lv = ind.get(_q("w","left")); rv = ind.get(_q("w","right"))
        hv = ind.get(_q("w","hanging")); fv = ind.get(_q("w","firstLine"))
        if lv: out["indent_left"]    = _twips(lv)
        if rv: out["indent_right"]   = _twips(rv)
        if hv: out["indent_hanging"] = _twips(hv)
        if fv: out["indent_first"]   = _twips(fv)

    # Spacing
    sp = ppr.find(_q("w","spacing"))
    if sp is not None:
        bv = sp.get(_q("w","before")); av = sp.get(_q("w","after"))
        lnv = sp.get(_q("w","line")); lnrule = sp.get(_q("w","lineRule"))
        if bv: out["space_before"] = _twips(bv)
        if av: out["space_after"]  = _twips(av)
        if lnv:
            out["line_spacing"]      = _twips(lnv)
            out["line_spacing_rule"] = lnrule or "auto"

    # Tab stops
    tabs = ppr.find(_q("w","tabs"))
    if tabs is not None:
        tab_list = []
        for tab in tabs.findall(_q("w","tab")):
            pos = tab.get(_q("w","pos"))
            leader = tab.get(_q("w","leader"))
            kind   = tab.get(_q("w","val"), "left")
            if pos:
                tab_list.append({"pos": _twips(pos), "type": kind,
                                  "leader": leader or "none"})
        if tab_list: out["tab_stops"] = tab_list

    # RTL bidi
    bidi = ppr.find(_q("w","bidi"))
    if bidi is not None:
        val = bidi.get(_q("w","val"))
        out["bidi"] = (val not in ("0","false")) if val else True

    # Keep together / keep with next
    if ppr.find(_q("w","keepLines"))  is not None: out["keep_lines"] = True
    if ppr.find(_q("w","keepNext"))   is not None: out["keep_next"]  = True
    if ppr.find(_q("w","pageBreakBefore")) is not None: out["page_break_before"] = True

    # Contextual spacing (suppresses space between same-style paragraphs)
    if ppr.find(_q("w","contextualSpacing")) is not None:
        cs_el = ppr.find(_q("w","contextualSpacing"))
        val = cs_el.get(_q("w","val"))
        out["contextual_spacing"] = (val not in ("0","false")) if val else True

    # Auto-spacing flags
    sp = ppr.find(_q("w","spacing"))
    if sp is not None:
        if sp.get(_q("w","beforeAutospacing")) in ("1","true","on"):
            out["before_autospacing"] = True
        if sp.get(_q("w","afterAutospacing")) in ("1","true","on"):
            out["after_autospacing"] = True

    # Style name
    pStyle = ppr.find(_q("w","pStyle"))
    if pStyle is not None:
        out["style"] = pStyle.get(_q("w","val"))

    # numPr (list)
    numPr = ppr.find(_q("w","numPr"))
    if numPr is not None:
        ilvl = numPr.find(_q("w","ilvl"))
        numId = numPr.find(_q("w","numId"))
        if ilvl is not None and numId is not None:
            out["list_level"] = int(ilvl.get(_q("w","val"),0))
            out["list_id"]    = int(numId.get(_q("w","val"),0))

    # rPr inside pPr (paragraph mark formatting)
    rpr = ppr.find(_q("w","rPr"))
    if rpr is not None:
        out["pmark_fmt"] = _extract_rpr(rpr)

    return out

# ── Extract runs from a paragraph element ────────────────────────────────────
_RUN_WRAPPER_TAGS = {
    "sdt", "sdtContent", "ins", "del", "hyperlink",
    "smartTag", "customXml", "fldSimple", "bdo", "dir",
}

def _iter_runs_recursive(el):
    """Yield every w:r in el. Hard-stops at txbxContent/drawing/pict."""
    for child in el:
        try:
            local = etree.QName(child.tag).localname
        except Exception:
            continue
        if local == "r":
            yield child
        elif local in ("txbxContent", "drawing", "pict"):
            pass  # hard stop — never cross into embedded content
        elif local in _RUN_WRAPPER_TAGS:
            yield from _iter_runs_recursive(child)

def _extract_runs(p_el):
    runs = []
    for r in _iter_runs_recursive(p_el):
        rpr = r.find(_q("w","rPr"))
        fmt = _extract_rpr(rpr)
        # Iterate children in document order to preserve t/tab/br sequence
        text_parts = []
        for child in r:
            try:
                local = etree.QName(child.tag).localname
            except Exception:
                continue
            if local == "t":
                text_parts.append(child.text or "")
            elif local == "tab":
                text_parts.append("\t")
            elif local == "br":
                text_parts.append("\n")
        text = "".join(text_parts)
        if not text:
            continue
        run_data = {"text": text}
        run_data.update(fmt)
        runs.append(run_data)
    return runs

# ── Paragraph text (flat) ─────────────────────────────────────────────────────
def _para_text(p_el):
    parts = []
    for r in _iter_runs_recursive(p_el):
        for t in r.findall(_q("w","t")):
            parts.append(t.text or "")
        for _ in r.findall(_q("w","tab")):
            parts.append("\t")
    return "".join(parts)

# ── Border extraction ─────────────────────────────────────────────────────────
def _extract_border(b_el):
    if b_el is None: return None
    return {
        "style": b_el.get(_q("w","val"), "single"),
        "size":  _twips(b_el.get(_q("w","sz"), 4)),
        "color": _norm_color(b_el.get(_q("w","color"), "000000")),
        "space": _twips(b_el.get(_q("w","space"), 0)),
    }

def _extract_borders_set(borders_el):
    if borders_el is None: return {}
    out = {}
    for side in ("top","bottom","left","right","insideH","insideV"):
        b = borders_el.find(_q("w",side))
        if b is not None:
            bdata = _extract_border(b)
            if bdata: out[side] = bdata
    return out

# ── Table extraction ──────────────────────────────────────────────────────────
def _extract_table(tbl_el):
    """Extract a <w:tbl> into a dict."""
    result = {"type": "table", "rows": []}

    # Table properties
    tblPr = tbl_el.find(_q("w","tblPr"))
    if tblPr is not None:
        # Width
        tblW = tblPr.find(_q("w","tblW"))
        if tblW is not None:
            result["width"] = _twips(tblW.get(_q("w","w"),0))
            result["width_type"] = tblW.get(_q("w","type"),"dxa")

        # Alignment
        jc = tblPr.find(_q("w","jc"))
        if jc is not None:
            result["align"] = jc.get(_q("w","val"),"left")

        # bidiVisual — RTL table (columns flow right-to-left, table anchors right).
        # This is the correct RTL mechanism; w:jc="right" is NOT used for RTL tables.
        bv = tblPr.find(_q("w","bidiVisual"))
        if bv is not None:
            val = bv.get(_q("w","val"))
            result["bidi_visual"] = (val not in ("0","false")) if val else True

        # Borders
        tblBorders = tblPr.find(_q("w","tblBorders"))
        if tblBorders is not None:
            result["borders"] = _extract_borders_set(tblBorders)

        # Spacing
        tblCellMar = tblPr.find(_q("w","tblCellMar"))
        if tblCellMar is not None:
            margins = {}
            for side in ("top","bottom","left","right"):
                m = tblCellMar.find(_q("w",side))
                if m is not None:
                    margins[side] = _twips(m.get(_q("w","w"),0))
            if margins: result["cell_margins"] = margins

    # Column widths from tblGrid
    tblGrid = tbl_el.find(_q("w","tblGrid"))
    if tblGrid is not None:
        col_widths = []
        for gc in tblGrid.findall(_q("w","gridCol")):
            col_widths.append(_twips(gc.get(_q("w","w"),0)))
        result["col_widths"] = col_widths

    # Rows
    for tr_el in tbl_el.findall(_q("w","tr")):
        row = {"cells": []}

        trPr = tr_el.find(_q("w","trPr"))
        if trPr is not None:
            trH = trPr.find(_q("w","trHeight"))
            if trH is not None:
                row["height"] = _twips(trH.get(_q("w","val"),0))
                row["height_rule"] = trH.get(_q("w","hRule"),"auto")
            if trPr.find(_q("w","tblHeader")) is not None:
                row["header"] = True

        for tc_el in tr_el.findall(_q("w","tc")):
            cell = {"paragraphs": []}

            tcPr = tc_el.find(_q("w","tcPr"))
            if tcPr is not None:
                # Width
                tcW = tcPr.find(_q("w","tcW"))
                if tcW is not None:
                    cell["width"] = _twips(tcW.get(_q("w","w"),0))

                # Borders
                tcBorders = tcPr.find(_q("w","tcBorders"))
                if tcBorders is not None:
                    cell["borders"] = _extract_borders_set(tcBorders)

                # Shading / background
                shd = tcPr.find(_q("w","shd"))
                if shd is not None:
                    fill = shd.get(_q("w","fill"))
                    cell["bg"] = _norm_color(fill)

                # Vertical alignment
                vAlign = tcPr.find(_q("w","vAlign"))
                if vAlign is not None:
                    cell["valign"] = vAlign.get(_q("w","val"),"top")

                # Span
                gridSpan = tcPr.find(_q("w","gridSpan"))
                if gridSpan is not None:
                    gs = int(gridSpan.get(_q("w","val"),1))
                    if gs > 1: cell["colspan"] = gs

                # Vertical merge
                vMerge = tcPr.find(_q("w","vMerge"))
                if vMerge is not None:
                    vmval = vMerge.get(_q("w","val"))
                    cell["vmerge"] = "restart" if vmval == "restart" else "continue"

                # Cell margins override
                tcMar = tcPr.find(_q("w","tcMar"))
                if tcMar is not None:
                    cmarg = {}
                    for side in ("top","bottom","left","right"):
                        m = tcMar.find(_q("w",side))
                        if m is not None:
                            cmarg[side] = _twips(m.get(_q("w","w"),0))
                    if cmarg: cell["margins"] = cmarg

            # Paragraphs inside cell
            for p_el in tc_el.findall(_q("w","p")):
                pdata = _extract_paragraph(p_el)
                cell["paragraphs"].append(pdata)

            row["cells"].append(cell)

        result["rows"].append(row)

    return result

# ── Paragraph extraction ──────────────────────────────────────────────────────
def _extract_paragraph(p_el):
    ppr  = p_el.find(_q("w","pPr"))
    out  = _extract_ppr(ppr)
    out["type"] = "paragraph"
    out["runs"] = _extract_runs(p_el)
    out["text"] = _para_text(p_el)
    return out

# ── Text-box extraction (modern wps:wsp) ─────────────────────────────────────
def _extract_textbox_wps(anchor_el):
    """Extract an anchor-based text box (Word 2010+)."""
    result = {"type": "textbox"}

    # Position
    pos_h = anchor_el.find(_q("wp","positionH"))
    pos_v = anchor_el.find(_q("wp","positionV"))
    if pos_h is not None:
        posOffset = pos_h.find(_q("wp","posOffset"))
        if posOffset is not None and posOffset.text:
            result["pos_x"] = _emu_to_twips(posOffset.text)
        result["pos_h_anchor"] = pos_h.get("relativeFrom","page")
    if pos_v is not None:
        posOffset = pos_v.find(_q("wp","posOffset"))
        if posOffset is not None and posOffset.text:
            result["pos_y"] = _emu_to_twips(posOffset.text)
        result["pos_v_anchor"] = pos_v.get("relativeFrom","page")

    # Size
    extent = anchor_el.find(_q("wp","extent"))
    if extent is not None:
        result["width"]  = _emu_to_twips(extent.get("cx","0"))
        result["height"] = _emu_to_twips(extent.get("cy","0"))

    # Find wps:txbx
    txbx = anchor_el.find(".//" + _q("wps","txbx"))
    if txbx is None: return None

    # Shape properties (border)
    spPr = anchor_el.find(".//" + _q("wps","spPr"))
    if spPr is not None:
        ln = spPr.find(".//" + _q("a","ln"))
        result["has_border"] = ln is not None

    # Extract paragraphs from text box
    result["paragraphs"] = []
    for p_el in txbx.findall(".//" + _q("w","p")):
        result["paragraphs"].append(_extract_paragraph(p_el))

    return result

# ── Text-box extraction (legacy VML v:shape) ─────────────────────────────────
def _extract_textbox_vml(pict_el):
    """Extract a VML-style text box."""
    result = {"type": "textbox", "paragraphs": []}

    # Find v:shape style for positioning
    shape = pict_el.find(".//{urn:schemas-microsoft-com:vml}shape")
    if shape is not None:
        style_str = shape.get("style","")
        def _css_val(prop):
            m = re.search(rf"{prop}:\s*([^;]+)", style_str)
            return m.group(1).strip() if m else None

        def _parse_dim(val):
            """Parse CSS dimension like '1.5in', '120pt', '3cm' → twips."""
            if not val: return 0
            val = val.strip()
            m = re.match(r"([\d.]+)\s*(in|pt|cm|mm|px|twips?)?", val)
            if not m: return 0
            n = float(m.group(1)); unit = (m.group(2) or "").lower()
            if unit in ("in",""):   return int(n * 1440)
            if unit == "pt":        return int(n * 20)
            if unit == "cm":        return int(n * 1440 / 2.54)
            if unit == "mm":        return int(n * 1440 / 25.4)
            if unit == "px":        return int(n * 15)
            return int(n)

        result["width"]  = _parse_dim(_css_val("width"))
        result["height"] = _parse_dim(_css_val("height"))

        ml  = _css_val("margin-left")
        mt  = _css_val("margin-top")
        msl = _css_val("mso-left")
        mst = _css_val("mso-top")
        result["pos_x"] = _parse_dim(ml or msl or "0")
        result["pos_y"] = _parse_dim(mt or mst or "0")

        # Border
        result["has_border"] = bool(shape.get("stroked","f") not in ("f","false",""))

    # Find w:txbxContent paragraphs
    ns_w = NS["w"]
    for p_el in pict_el.findall(f".//{{{ns_w}}}p"):
        result["paragraphs"].append(_extract_paragraph(p_el))

    return result if result["paragraphs"] else None

# ── Drawing / image ───────────────────────────────────────────────────────────
def _extract_drawing(drawing_el, rels):
    """Extract an inline drawing or anchor."""
    # Check anchor (text box or floating image)
    for anchor in drawing_el.findall(_q("wp","anchor")):
        # Check if it has wps:txbx (text box)
        if anchor.find(".//" + _q("wps","txbx")) is not None:
            return _extract_textbox_wps(anchor)
        # Otherwise it's a floating image
        return _extract_image_anchor(anchor, rels, floating=True)

    # Inline image
    for inline in drawing_el.findall(_q("wp","inline")):
        return _extract_image_inline(inline, rels)

    return None

def _extract_image_anchor(anchor_el, rels, floating=False):
    blip = anchor_el.find(".//" + _q("a","blip"))
    result = {"type": "image", "floating": floating}
    if blip is not None:
        embed = blip.get(_q("r","embed"))
        if embed and rels:
            result["rId"] = embed
            result["target"] = rels.get(embed,{}).get("target","")
    extent = anchor_el.find(_q("wp","extent"))
    if extent is not None:
        result["width"]  = _emu_to_twips(extent.get("cx",0))
        result["height"] = _emu_to_twips(extent.get("cy",0))
    return result

def _extract_image_inline(inline_el, rels):
    blip = inline_el.find(".//" + _q("a","blip"))
    result = {"type": "image", "floating": False}
    if blip is not None:
        embed = blip.get(_q("r","embed"))
        if embed and rels:
            result["rId"] = embed
            result["target"] = rels.get(embed,{}).get("target","")
    extent = inline_el.find(_q("wp","extent"))
    if extent is not None:
        result["width"]  = _emu_to_twips(extent.get("cx",0))
        result["height"] = _emu_to_twips(extent.get("cy",0))
    return result

# ── Main paragraph body element dispatcher ────────────────────────────────────
def _collect_body_elements(body_el, rels):
    """Walk the document body and collect all elements in order."""
    elements = []

    for child in body_el:
        tag = etree.QName(child.tag).localname if child.tag != etree.Comment else ""

        if tag == "tbl":
            elements.append(_extract_table(child))

        elif tag == "p":
            floats = []
            found_wps_textbox = False

            for drawing in child.findall(".//" + _q("w","drawing")):
                drawn = _extract_drawing(drawing, rels)
                if drawn:
                    floats.append(drawn)
                    if drawn.get("type") == "textbox":
                        found_wps_textbox = True

            if not found_wps_textbox:
                pict = child.find(_q("w","pict"))
                if pict is not None:
                    tb = _extract_textbox_vml(pict)
                    if tb:
                        floats.append(tb)

            pdata = _extract_paragraph(child)
            has_text = bool(pdata["text"].strip() or
                            "\t" in pdata.get("text", "") or
                            pdata.get("runs"))

            if floats and not has_text:
                elements.extend(floats)
            elif floats and has_text:
                elements.append(pdata)
                elements.extend(floats)
            elif has_text:
                elements.append(pdata)
            else:
                elements.append({"type": "empty_paragraph",
                                  **{k: v for k, v in pdata.items()
                                     if k not in ("type", "runs", "text")}})

        elif tag == "sectPr":
            pass

    return elements

# ── Page / section properties ─────────────────────────────────────────────────
def _extract_section(sect_el):
    if sect_el is None:
        return {"width": 12240, "height": 15840,
                "margin_top": 1440, "margin_bottom": 1440,
                "margin_left": 1800, "margin_right": 1800}
    out = {}
    pgSz = sect_el.find(_q("w","pgSz"))
    if pgSz is not None:
        out["width"]  = _twips(pgSz.get(_q("w","w"), 12240))
        out["height"] = _twips(pgSz.get(_q("w","h"), 15840))
        orient = pgSz.get(_q("w","orient"))
        if orient: out["orientation"] = orient
    pgMar = sect_el.find(_q("w","pgMar"))
    if pgMar is not None:
        out["margin_top"]    = _twips(pgMar.get(_q("w","top"),    1440))
        out["margin_bottom"] = _twips(pgMar.get(_q("w","bottom"), 1440))
        out["margin_left"]   = _twips(pgMar.get(_q("w","left"),   1800))
        out["margin_right"]  = _twips(pgMar.get(_q("w","right"),  1800))
        out["margin_header"] = _twips(pgMar.get(_q("w","header"),  720))
        out["margin_footer"] = _twips(pgMar.get(_q("w","footer"),  720))
    cols = sect_el.find(_q("w","cols"))
    if cols is not None:
        out["columns"] = int(cols.get(_q("w","num"), 1))
    return out

# ── Default styles ─────────────────────────────────────────────────────────────
def _extract_default_styles(styles_xml):
    if styles_xml is None:
        return {"font": "Calibri", "size": 11}
    tree = etree.fromstring(styles_xml)
    out  = {}
    docDefaults = tree.find(_q("w","docDefaults"))
    if docDefaults is not None:
        # Run defaults (font, size)
        rPrDf = docDefaults.find(".//" + _q("w","rPrDefault"))
        if rPrDf is not None:
            rpr = rPrDf.find(_q("w","rPr"))
            if rpr is not None:
                rFonts = rpr.find(_q("w","rFonts"))
                if rFonts is not None:
                    out["font"] = (rFonts.get(_q("w","ascii")) or
                                   rFonts.get(_q("w","hAnsi")) or "Calibri")
                sz = rpr.find(_q("w","sz"))
                if sz is not None:
                    out["size"] = _pt_half_to_pt(sz.get(_q("w","val"),22))

        # Paragraph defaults (spacing, line spacing)
        pPrDf = docDefaults.find(".//" + _q("w","pPrDefault"))
        if pPrDf is not None:
            ppr = pPrDf.find(_q("w","pPr"))
            if ppr is not None:
                sp = ppr.find(_q("w","spacing"))
                if sp is not None:
                    bv = sp.get(_q("w","before")); av = sp.get(_q("w","after"))
                    lnv = sp.get(_q("w","line")); lnrule = sp.get(_q("w","lineRule"))
                    if bv is not None: out["default_space_before"] = _twips(bv)
                    if av is not None: out["default_space_after"]  = _twips(av)
                    if lnv is not None:
                        out["default_line_spacing"]      = _twips(lnv)
                        out["default_line_spacing_rule"] = lnrule or "auto"

    out.setdefault("font","Calibri")
    out.setdefault("size",11)

    # Extract named style spacing (especially "Normal")
    style_spacing = {}
    for style_el in tree.findall(_q("w","style")):
        style_id = style_el.get(_q("w","styleId"))
        if not style_id:
            continue
        pPr = style_el.find(_q("w","pPr"))
        if pPr is not None:
            sp = pPr.find(_q("w","spacing"))
            if sp is not None:
                sdata = {}
                bv = sp.get(_q("w","before")); av = sp.get(_q("w","after"))
                lnv = sp.get(_q("w","line")); lnrule = sp.get(_q("w","lineRule"))
                if bv is not None: sdata["space_before"] = _twips(bv)
                if av is not None: sdata["space_after"]  = _twips(av)
                if lnv is not None:
                    sdata["line_spacing"]      = _twips(lnv)
                    sdata["line_spacing_rule"] = lnrule or "auto"
                if sdata:
                    style_spacing[style_id] = sdata

    if style_spacing:
        out["style_spacing"] = style_spacing

    return out

# ── Relationships ─────────────────────────────────────────────────────────────
def _load_rels(zf, rel_path):
    if rel_path not in zf.namelist():
        return {}
    tree = etree.fromstring(zf.read(rel_path))
    out = {}
    for rel in tree:
        rid    = rel.get("Id")
        target = rel.get("Target","")
        rtype  = rel.get("Type","")
        if rid:
            out[rid] = {"target": target, "type": rtype}
    return out

# ── Main extraction function ───────────────────────────────────────────────────
def extract_docx(docx_path):
    """
    Extract a .docx file into a document dict.
    Returns a dict ready for json.dumps().
    """
    doc_data = {
        "source": os.path.basename(docx_path),
        "meta":   {},
        "elements": []
    }

    with zipfile.ZipFile(docx_path, "r") as zf:
        names = zf.namelist()

        # Parse document XML
        doc_xml    = zf.read("word/document.xml")
        styles_xml = zf.read("word/styles.xml") if "word/styles.xml" in names else None
        rels       = _load_rels(zf, "word/_rels/document.xml.rels")

        tree = etree.fromstring(doc_xml)
        body = tree.find(".//" + _q("w","body"))
        if body is None:
            raise ValueError("No <w:body> found in document.xml")

        # Section / page properties
        sectPr = body.find(_q("w","sectPr"))
        doc_data["meta"]["page"] = _extract_section(sectPr)

        # Default styles
        doc_data["meta"]["defaults"] = _extract_default_styles(styles_xml)

        # Body elements
        doc_data["elements"] = _collect_body_elements(body, rels)

    return doc_data

# ── Convenience entry point ───────────────────────────────────────────────────
def extract_to_file(docx_path, out_dir="to_html"):
    """Extract docx_path → out_dir/<basename>.json"""
    os.makedirs(out_dir, exist_ok=True)
    basename = os.path.splitext(os.path.basename(docx_path))[0]
    out_path = os.path.join(out_dir, basename + ".json")
    data = extract_docx(docx_path)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return out_path

if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage: python word_to_html.py <file.docx> [out_dir]")
        sys.exit(1)
    out_dir = sys.argv[2] if len(sys.argv) > 2 else "to_html"
    result  = extract_to_file(sys.argv[1], out_dir)
    print(f"Extracted → {result}")