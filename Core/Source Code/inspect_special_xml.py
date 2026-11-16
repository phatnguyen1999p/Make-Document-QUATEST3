import os
import docx
from docx.oxml.ns import qn
from lxml import etree

base_dir = r"c:\Users\phat.nt\Desktop\Source\Làm KQ"

for folder in ["New From 2026", "Old Form 2025"]:
    folder_path = os.path.join(base_dir, folder)
    for fname in sorted(os.listdir(folder_path)):
        if not fname.endswith(".docx") or fname.startswith("~$"):
            continue
        fpath = os.path.join(folder_path, fname)
        doc = docx.Document(fpath)
        syms = doc.element.body.findall(f".//{qn('w:sym')}")
        picts = doc.element.body.findall(f".//{qn('w:pict')}")
        drawings = doc.element.body.findall(f".//{qn('w:drawing')}")
        print(f"{folder}/{fname}: syms={len(syms)}, picts={len(picts)}, drawings={len(drawings)}")
        for s in syms[:5]:
            print("  sym:", s.attrib)
        # Check Table #9 col 1 in dot 3
        if "đợt 3" in fname:
            t9 = doc.tables[9]
            for r_i, row in enumerate(t9.rows):
                c1_xml = etree.tostring(row.cells[1]._tc, encoding='unicode')
                has_pict = 'pict' in c1_xml or 'drawing' in c1_xml or 'sym' in c1_xml
                print(f"  Dot3 T9 R{r_i} C1 has_special={has_pict}, xml_len={len(c1_xml)}")
                if r_i == 1:
                    print("   ", c1_xml[:400])
