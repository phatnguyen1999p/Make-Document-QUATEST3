import os
import docx
from docx.oxml.ns import qn

base_dir = r"c:\Users\phat.nt\Desktop\Source\Làm KQ"
out_path = r"C:\Users\phat.nt\.gemini\antigravity\brain\534de527-a676-4430-9dfa-9b3970a21af9\scratch\body_indices.txt"

with open(out_path, "w", encoding="utf-8") as f:
    for folder in ["New From 2026", "Old Form 2025"]:
        folder_path = os.path.join(base_dir, folder)
        f.write(f"=== {folder} ===\n")
        for fname in sorted(os.listdir(folder_path)):
            if not fname.endswith(".docx") or fname.startswith("~$"):
                continue
            doc = docx.Document(os.path.join(folder_path, fname))
            f.write(f"\n--- {fname} ---\n")
            for idx, child in enumerate(doc.element.body):
                tag = child.tag.split('}')[-1]
                if tag == 'p':
                    txt = "".join(child.itertext()).strip().replace('\n', ' ')
                    sectPr = child.find(f".//{qn('w:sectPr')}")
                    numPr = child.find(f".//{qn('w:numPr')}")
                    drawings = child.findall(f".//{qn('w:drawing')}") + child.findall(f".//{qn('w:pict')}")
                    extra = []
                    if sectPr is not None: extra.append("SECT_BRK")
                    if numPr is not None:
                        nid = numPr.find(qn('w:numId')).get(qn('w:val'))
                        ilvl_el = numPr.find(qn('w:ilvl'))
                        lvl = ilvl_el.get(qn('w:val')) if ilvl_el is not None else '0'
                        extra.append(f"NUM({nid}:{lvl})")
                    if drawings: extra.append(f"DRAW({len(drawings)})")
                    ex_str = f" [{', '.join(extra)}]" if extra else ""
                    f.write(f"  [{idx:02d}] P{ex_str}: '{txt[:85]}'\n")
                elif tag == 'tbl':
                    rows = child.findall(qn('w:tr'))
                    c0_txt = "".join(rows[0].itertext()).strip().replace('\n', ' ')[:70] if rows else ""
                    r1_txt = "".join(rows[1].itertext()).strip().replace('\n', ' ')[:70] if len(rows) > 1 else ""
                    f.write(f"  [{idx:02d}] TBL ({len(rows)} rows): R0='{c0_txt}' | R1='{r1_txt}'\n")
                elif tag == 'sectPr':
                    f.write(f"  [{idx:02d}] FINAL_SECTPR\n")
