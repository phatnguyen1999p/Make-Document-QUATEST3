"""
DOCX Inspector & Dumper for QUATEST 3 (Room N2).
Replaces dump-all.ps1 and set-cells.ps1 -Dump, plus full header/footer/media/crop/docPr inspection.
"""
import os
import sys
import zipfile
import hashlib
import argparse
from lxml import etree
import docx

sys.stdout.reconfigure(encoding="utf-8")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"


def w_tag(name: str) -> str:
    return f"{{{W_NS}}}{name}"


def dump_xml_runs(zf: zipfile.ZipFile, entry_name: str, out_lines: list[str]):
    """Dump all <w:t> runs with 0-based index, highlight, bold, vertAlign, plus <w:instrText>."""
    raw = zf.read(entry_name)
    root = etree.fromstring(raw)
    t_elems = root.findall(f".//{{{W_NS}}}t")
    instr_elems = root.findall(f".//{{{W_NS}}}instrText")

    out_lines.append(f"\n--- ENTRY: {entry_name} (w:t count={len(t_elems)}, w:instrText count={len(instr_elems)}) ---")
    if instr_elems:
        instr_txts = [el.text or "" for el in instr_elems]
        out_lines.append(f"  [FIELDS instrText]: {instr_txts}")

    for idx, t_el in enumerate(t_elems):
        txt = t_el.text or ""
        r_el = t_el.getparent()
        flags = []
        if r_el is not None and r_el.tag == w_tag("r"):
            rPr = r_el.find(w_tag("rPr"))
            if rPr is not None:
                if rPr.find(w_tag("b")) is not None:
                    b_val = rPr.find(w_tag("b")).get(w_tag("val"))
                    if b_val not in ("0", "false"):
                        flags.append("B")
                hl = rPr.find(w_tag("highlight"))
                if hl is not None:
                    flags.append(f"HL:{hl.get(w_tag('val'), 'yellow')}")
                va = rPr.find(w_tag("vertAlign"))
                if va is not None:
                    flags.append(f"VA:{va.get(w_tag('val'))}")
        # Also check paragraph-level highlight
        flag_str = f" [{','.join(flags)}]" if flags else ""
        disp = txt.replace("\n", "\\n").replace("\t", "\\t")
        out_lines.append(f"  [{idx:04d}]{flag_str}: '{disp}'")


def dump_tables(docx_path: str, out_lines: list[str]):
    """Dump all tables with (table, row, col) 0-based coordinates, vMerge/hMerge, bold, highlight, and drawings."""
    doc = docx.Document(docx_path)
    out_lines.append(f"\n--- TABLES MAP (Total tables: {len(doc.tables)}) ---")
    for t_idx, tbl in enumerate(doc.tables):
        rows = len(tbl.rows)
        cols = len(tbl.columns)
        out_lines.append(f"\n[TABLE #{t_idx} ({rows}x{cols})]")
        for r_idx, row in enumerate(tbl.rows):
            cells_desc = []
            prev_tc = None
            for c_idx, cell in enumerate(row.cells):
                tc = cell._tc
                if tc is prev_tc:
                    cells_desc.append(f"C{c_idx}:<H-MERGE>")
                    continue
                prev_tc = tc
                vMerge = tc.findall(f".//{{{W_NS}}}vMerge")
                vm_tag = ""
                if vMerge:
                    val = vMerge[0].get(w_tag("val"))
                    vm_tag = f"<VM:{val or 'cont'}>"
                drawings = tc.findall(f".//{{{W_NS}}}drawing") + tc.findall(f".//{{{W_NS}}}pict")
                drw_tag = f"<IMG:{len(drawings)}>" if drawings else ""
                hl_elems = tc.findall(f".//{{{W_NS}}}highlight")
                hl_tag = "<HL>" if hl_elems else ""
                b_elems = [
                    b for b in tc.findall(f".//{{{W_NS}}}b")
                    if b.get(w_tag("val")) not in ("0", "false")
                ]
                b_tag = "<B>" if b_elems else ""
                txt = cell.text.strip().replace("\n", " \\n ")
                if len(txt) > 65:
                    txt = txt[:62] + "..."
                cells_desc.append(f"C{c_idx}:{vm_tag}{drw_tag}{hl_tag}{b_tag}'{txt}'")
            out_lines.append(f"  R{r_idx:02d}: " + " | ".join(cells_desc))


def dump_images_and_rels(docx_path: str, out_lines: list[str]):
    """Inspect all drawings, rId -> media target, dimensions, srcRect crop, docPr IDs, and MD5."""
    out_lines.append("\n--- MEDIA & DRAWINGS MAP ---")
    with zipfile.ZipFile(docx_path, "r") as zf:
        rels_map = {}
        if "word/_rels/document.xml.rels" in zf.namelist():
            rels_root = etree.fromstring(zf.read("word/_rels/document.xml.rels"))
            for rel in rels_root:
                r_id = rel.get("Id")
                target = rel.get("Target")
                rels_map[r_id] = target

        doc_root = etree.fromstring(zf.read("word/document.xml"))
        drawings = doc_root.findall(f".//{{{W_NS}}}drawing")
        docPr_ids = []
        for d_idx, drw in enumerate(drawings):
            docPr = drw.find(f".//{{{WP_NS}}}docPr")
            dp_id = docPr.get("id") if docPr is not None else None
            dp_name = docPr.get("name") if docPr is not None else ""
            if dp_id is not None:
                docPr_ids.append(dp_id)
            ext_el = drw.find(f".//{{{WP_NS}}}extent")
            cx_cm = int(ext_el.get("cx", "0")) / 360000.0 if ext_el is not None else 0
            cy_cm = int(ext_el.get("cy", "0")) / 360000.0 if ext_el is not None else 0
            blip = drw.find(f".//{{{A_NS}}}blip")
            r_embed = blip.get(f"{{{R_NS}}}embed") if blip is not None else None
            srcRect = drw.find(f".//{{{A_NS}}}srcRect")
            crop_info = dict(srcRect.attrib) if srcRect is not None and srcRect.attrib else None
            target = rels_map.get(r_embed, "")
            media_path = f"word/{target}" if target and not target.startswith("word/") else target
            md5_str = ""
            size_b = 0
            if media_path and media_path in zf.namelist():
                data = zf.read(media_path)
                size_b = len(data)
                md5_str = hashlib.md5(data).hexdigest()[:10]
            out_lines.append(
                f"  Drawing #{d_idx:02d}: docPr_id={dp_id} ('{dp_name}') | "
                f"rId={r_embed} -> {media_path} ({size_b:,}B, md5={md5_str}) | "
                f"size={cx_cm:.2f}x{cy_cm:.2f}cm | crop={crop_info}"
            )
        # Check duplicate docPr IDs (Bẫy 34)
        dup_ids = set(x for x in docPr_ids if docPr_ids.count(x) > 1)
        if dup_ids:
            out_lines.append(f"  [WARNING - Bẫy 34]: Duplicate wp:docPr IDs detected: {sorted(dup_ids)}")


def inspect_docx(docx_path: str, dump_runs: bool = True, dump_tbls: bool = True, dump_imgs: bool = True) -> str:
    lines = [
        "=" * 95,
        f"FILE: {docx_path}",
        "=" * 95,
    ]
    with zipfile.ZipFile(docx_path, "r") as zf:
        names = zf.namelist()
        if dump_runs:
            # Order: document.xml, header1..9.xml, footer1..9.xml, comments.xml
            entries = ["word/document.xml"]
            entries += sorted(n for n in names if n.startswith("word/header") and n.endswith(".xml"))
            entries += sorted(n for n in names if n.startswith("word/footer") and n.endswith(".xml"))
            entries += sorted(n for n in names if n.startswith("word/comments") and n.endswith(".xml"))
            for ent in entries:
                if ent in names:
                    dump_xml_runs(zf, ent, lines)

    if dump_tbls:
        dump_tables(docx_path, lines)
    if dump_imgs:
        dump_images_and_rels(docx_path, lines)

    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(description="DOCX Inspector (Runs, Headers/Footers, Tables, Media)")
    parser.add_argument("target", help="Path to a .docx file or directory containing .docx files")
    parser.add_argument("-o", "--out", help="Output text file path")
    parser.add_argument("--no-runs", action="store_true", help="Skip dumping <w:t> runs")
    parser.add_argument("--no-tables", action="store_true", help="Skip dumping table map")
    parser.add_argument("--no-images", action="store_true", help="Skip dumping media/drawing map")
    args = parser.parse_args()

    targets = []
    if os.path.isdir(args.target):
        for root, _, files in os.walk(args.target):
            for f in sorted(files):
                if f.lower().endswith(".docx") and not f.startswith("~$"):
                    targets.append(os.path.join(root, f))
    else:
        targets.append(args.target)

    all_out = []
    for t in targets:
        all_out.append(
            inspect_docx(
                t,
                dump_runs=not args.no_runs,
                dump_tbls=not args.no_tables,
                dump_imgs=not args.no_images,
            )
        )
    result = "\n\n".join(all_out)
    if args.out:
        os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(result)
        print(f"Saved dump of {len(targets)} file(s) -> {args.out}")
    else:
        print(result)


if __name__ == "__main__":
    main()
