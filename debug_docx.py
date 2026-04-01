"""
debug_docx.py
─────────────
Run this on your ORIGINAL .docx to dump the raw XML of every body element.
This lets us see exactly why "Resident Identity No." is not being extracted.

Usage:
    python debug_docx.py "path/to/your.docx"

Output: debug_output.txt  (in the same folder as this script)
"""

import sys
import zipfile
import os
from lxml import etree

def main():
    if len(sys.argv) < 2:
        print("Usage: python debug_docx.py <your_file.docx>")
        sys.exit(1)

    docx_path = sys.argv[1]
    out_path  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "debug_output.txt")

    NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    with zipfile.ZipFile(docx_path, "r") as zf:
        doc_xml = zf.read("word/document.xml")

    tree = etree.fromstring(doc_xml)
    body = tree.find(f"{{{NS_W}}}body")

    lines = []
    lines.append(f"=== debug_docx.py : {os.path.basename(docx_path)} ===\n")

    def flat_text(el):
        """Extract all text inside element, including inside w:sdt wrappers."""
        parts = []
        for t in el.iter(f"{{{NS_W}}}t"):
            parts.append(t.text or "")
        for _ in el.iter(f"{{{NS_W}}}tab"):
            parts.append("\t")
        return "".join(parts)

    for idx, child in enumerate(body):
        try:
            local = etree.QName(child.tag).localname
        except Exception:
            local = "?"

        txt = flat_text(child)[:120].replace("\n", "↵").replace("\t", "→")

        lines.append(f"[{idx:03d}] <{local}>  text={repr(txt)}")

        # For paragraphs: show child tag names so we can see w:sdt etc.
        if local == "p":
            child_tags = []
            for c in child:
                try:
                    child_tags.append(etree.QName(c.tag).localname)
                except Exception:
                    child_tags.append("?")
            lines.append(f"       children: {child_tags}")

            # If it contains sdt, show the sdt children too
            for sdt in child.findall(f".//{{{NS_W}}}sdt"):
                sdt_text = flat_text(sdt)[:80].replace("\t", "→")
                sdt_children = [etree.QName(c.tag).localname for c in sdt]
                lines.append(f"       w:sdt text={repr(sdt_text)}  sdt-children={sdt_children}")

            # Show raw XML of the paragraph (truncated)
            raw = etree.tostring(child, encoding="unicode")
            lines.append(f"       RAW XML: {raw[:600]}")

        lines.append("")

    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    print(f"Done! Output written to: {out_path}")
    print(f"Total body children: {idx+1}")
    print("\nLooking for 'Resident':")
    for line in lines:
        if "Resident" in line or "resident" in line:
            print(" ", line)

if __name__ == "__main__":
    main()