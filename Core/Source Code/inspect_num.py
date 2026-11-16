import os
import docx
from docx.oxml.ns import qn

base_dir = r"c:\Users\phat.nt\Desktop\Source\Làm KQ"
out_path = r"C:\Users\phat.nt\.gemini\antigravity\brain\534de527-a676-4430-9dfa-9b3970a21af9\scratch\num_out.txt"

with open(out_path, "w", encoding="utf-8") as f:
    for folder in ["New From 2026", "Old Form 2025"]:
        folder_path = os.path.join(base_dir, folder)
        f.write(f"=== {folder} ===\n")
        for fname in sorted(os.listdir(folder_path)):
            if not fname.endswith(".docx") or fname.startswith("~$"):
                continue
            fpath = os.path.join(folder_path, fname)
            doc = docx.Document(fpath)
            f.write(f"\n--- {fname} ---\n")
            num_part = doc.part.numbering_part
            num_map = {}
            abs_map = {}
            for num in num_part.element.findall(qn('w:num')):
                nid = num.get(qn('w:numId'))
                anid_el = num.find(qn('w:abstractNumId'))
                num_map[nid] = anid_el.get(qn('w:val')) if anid_el is not None else None
            for anum in num_part.element.findall(qn('w:abstractNum')):
                anid = anum.get(qn('w:abstractNumId'))
                lvls = {}
                for lvl in anum.findall(qn('w:lvl')):
                    ilvl = lvl.get(qn('w:ilvl'))
                    lt = lvl.find(qn('w:lvlText'))
                    st = lvl.find(qn('w:start'))
                    lvls[ilvl] = (lt.get(qn('w:val')) if lt is not None else '', st.get(qn('w:val')) if st is not None else '1')
                abs_map[anid] = lvls
            
            # Check Table 0 numbering too
            t0 = doc.tables[0]
            for r_i, row in enumerate(t0.rows):
                c0 = row.cells[0]
                for p in c0.paragraphs:
                    numPr = p._p.find(f".//{qn('w:numPr')}")
                    if numPr is not None:
                        nid = numPr.find(qn('w:numId')).get(qn('w:val'))
                        ilvl_el = numPr.find(qn('w:ilvl'))
                        lvl = ilvl_el.get(qn('w:val')) if ilvl_el is not None else '0'
                        anid = num_map.get(nid)
                        linfo = abs_map.get(anid, {}).get(lvl)
                        f.write(f"  T0.R{r_i:02d} numId={nid} ilvl={lvl} fmt={linfo}: {p.text.strip()}\n")

            for i, p in enumerate(doc.paragraphs):
                numPr = p._p.find(f".//{qn('w:numPr')}")
                if numPr is not None:
                    nid = numPr.find(qn('w:numId')).get(qn('w:val'))
                    ilvl_el = numPr.find(qn('w:ilvl'))
                    lvl = ilvl_el.get(qn('w:val')) if ilvl_el is not None else '0'
                    anid = num_map.get(nid)
                    linfo = abs_map.get(anid, {}).get(lvl)
                    f.write(f"  P{i:02d} numId={nid} ilvl={lvl} fmt={linfo}: {p.text.strip()}\n")
