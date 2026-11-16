import os
import docx
from docx.oxml.ns import qn

base_dir = r"c:\Users\phat.nt\Desktop\Source\Làm KQ"

for folder in ["New From 2026", "Old Form 2025"]:
    folder_path = os.path.join(base_dir, folder)
    print(f"=== {folder} ===")
    for fname in sorted(os.listdir(folder_path)):
        if not fname.endswith(".docx") or fname.startswith("~$"):
            continue
        fpath = os.path.join(folder_path, fname)
        doc = docx.Document(fpath)
        print(f"\n--- {fname} ---")
        sec_idx = 0
        t_idx = 0
        for child in doc.element.body.iterchildren():
            if child.tag == qn('w:p'):
                txt = "".join(child.itertext()).strip()
                sectPr = child.find(f".//{qn('w:sectPr')}")
                drawings = child.findall(f".//{qn('w:drawing')}") + child.findall(f".//{qn('w:pict')}")
                img_str = f" [IMAGES: {len(drawings)}]" if drawings else ""
                if txt or img_str or sectPr is not None:
                    print(f"  [Sec {sec_idx}] P: {txt[:80]}{img_str}")
                if sectPr is not None:
                    orient = sectPr.find(f".//{qn('w:pgSz')}")
                    o_val = orient.get(qn('w:orient')) if orient is not None else 'portrait'
                    print(f"  >>> SECTION BREAK (end of Sec {sec_idx}, orient={o_val}) <<<")
                    sec_idx += 1
            elif child.tag == qn('w:tbl'):
                drawings = child.findall(f".//{qn('w:drawing')}") + child.findall(f".//{qn('w:pict')}")
                img_str = f" [IMAGES: {len(drawings)}]" if drawings else ""
                first_cell = "".join(child.findall(f".//{qn('w:tc')}")[0].itertext()).strip().replace('\n', ' ')[:50]
                print(f"  [Sec {sec_idx}] TABLE #{t_idx} ({first_cell}){img_str}")
                t_idx += 1
