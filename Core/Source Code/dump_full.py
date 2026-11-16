import os
import docx
from docx.document import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph

def iter_block_items(parent):
    if isinstance(parent, Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("something's not right")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

base_dir = r"c:\Users\phat.nt\Desktop\Source\Làm KQ"
out_path = r"C:\Users\phat.nt\.gemini\antigravity\brain\534de527-a676-4430-9dfa-9b3970a21af9\scratch\full_dump.txt"

with open(out_path, "w", encoding="utf-8") as f:
    for folder in ["New From 2026", "Old Form 2025"]:
        folder_path = os.path.join(base_dir, folder)
        f.write("=" * 100 + "\n")
        f.write(f"FOLDER: {folder}\n")
        f.write("=" * 100 + "\n")
        for fname in sorted(os.listdir(folder_path)):
            if not fname.endswith(".docx") or fname.startswith("~$"):
                continue
            fpath = os.path.join(folder_path, fname)
            f.write(f"\n### FILE: {fname} ###\n")
            doc = docx.Document(fpath)
            
            # Check headers/footers
            for s_i, sec in enumerate(doc.sections):
                hdr_txt = " | ".join([p.text.strip() for p in sec.header.paragraphs if p.text.strip()])
                ftr_txt = " | ".join([p.text.strip() for p in sec.footer.paragraphs if p.text.strip()])
                f.write(f"  Section {s_i}: orient={sec.orientation}, page_w={sec.page_width.mm:.1f}mm, page_h={sec.page_height.mm:.1f}mm, hdr='{hdr_txt}', ftr='{ftr_txt}'\n")

            t_idx = 0
            for item in iter_block_items(doc):
                if isinstance(item, Paragraph):
                    txt = item.text.strip()
                    # Check numbering
                    numPr = item._p.xpath('.//w:numPr')
                    num_info = " [NUM]" if numPr else ""
                    if txt:
                        f.write(f"[P ({item.style.name}){num_info}]: {txt}\n")
                elif isinstance(item, Table):
                    rows = len(item.rows)
                    cols = len(item.columns)
                    f.write(f"\n[TABLE #{t_idx} ({rows}x{cols}) style='{item.style.name if item.style else None}']:\n")
                    for r_i, row in enumerate(item.rows):
                        row_cells = []
                        prev_tc = None
                        for c_i, cell in enumerate(row.cells):
                            if cell._tc is prev_tc:
                                row_cells.append("<H-MERGE>")
                            else:
                                prev_tc = cell._tc
                                vMerge = cell._tc.xpath('.//w:vMerge')
                                vm_str = ""
                                if vMerge:
                                    val = vMerge[0].get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                                    vm_str = f"<V-MERGE:{val or 'cont'}>"
                                c_txt = cell.text.strip().replace('\n', ' \\n ')
                                row_cells.append(f"{vm_str}{c_txt}")
                        f.write(f"  R{r_i:02d}: " + " || ".join(row_cells) + "\n")
                    t_idx += 1
