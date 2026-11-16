import os
import sys
import docx
from docx.document import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph

sys.stdout.reconfigure(encoding='utf-8')

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

for folder in ["New From 2026", "Old Form 2025"]:
    folder_path = os.path.join(base_dir, folder)
    print("=" * 80)
    print(f"FOLDER: {folder}")
    print("=" * 80)
    for fname in sorted(os.listdir(folder_path)):
        if not fname.endswith(".docx") or fname.startswith("~$"):
            continue
        fpath = os.path.join(folder_path, fname)
        print(f"\n--- FILE: {fname} ---")
        doc = docx.Document(fpath)
        print(f"Total paragraphs: {len(doc.paragraphs)}, Total tables: {len(doc.tables)}, Sections: {len(doc.sections)}")
        
        t_idx = 0
        for item in iter_block_items(doc):
            if isinstance(item, Paragraph):
                txt = item.text.strip()
                if txt:
                    style_name = item.style.name if item.style else "None"
                    print(f"  [P ({style_name})]: {txt[:120]}")
            elif isinstance(item, Table):
                rows = len(item.rows)
                cols = len(item.columns)
                # get first 2 rows text summary
                preview_rows = []
                for r in range(min(3, rows)):
                    row_cells = []
                    for c in range(min(10, len(item.rows[r].cells))):
                        cell_txt = item.rows[r].cells[c].text.strip().replace('\n', ' ')
                        if len(cell_txt) > 35:
                            cell_txt = cell_txt[:35] + "..."
                        row_cells.append(cell_txt)
                    # deduplicate adjacent identical merged cells for readability
                    dedup = []
                    for rc in row_cells:
                        if not dedup or dedup[-1] != rc:
                            dedup.append(rc)
                    preview_rows.append(" | ".join(dedup))
                print(f"  [TABLE #{t_idx} ({rows}x{cols})]:")
                for pr_i, pr in enumerate(preview_rows):
                    print(f"    R{pr_i}: {pr}")
                t_idx += 1
