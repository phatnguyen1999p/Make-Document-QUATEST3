import os
import docx
from docx.oxml.ns import qn

dir_core = r"c:\Users\phat.nt\Desktop\Source\Làm KQ\Core"
out_path = os.path.join(dir_core, "Source Code", "verify_templates.txt")

with open(out_path, "w", encoding="utf-8") as f:
    for fname in sorted(os.listdir(dir_core)):
        if not fname.startswith("TEMPLATE"):
            continue
        doc = docx.Document(os.path.join(dir_core, fname))
        f.write("=" * 90 + "\n")
        f.write(f"FILE: {fname} (Tables: {len(doc.tables)}, Paragraphs: {len(doc.paragraphs)})\n")
        f.write("=" * 90 + "\n")
        t_idx = 0
        for idx, child in enumerate(doc.element.body):
            tag = child.tag.split('}')[-1]
            if tag == 'p':
                txt = "".join(child.itertext()).strip().replace('\n', ' ')
                numPr = child.find(f".//{qn('w:numPr')}")
                num_str = ""
                if numPr is not None:
                    nid = numPr.find(qn('w:numId')).get(qn('w:val'))
                    ilvl = numPr.find(qn('w:ilvl'))
                    lvl = ilvl.get(qn('w:val')) if ilvl is not None else '0'
                    num_str = f" [NUM {nid}:{lvl}]"
                drawings = child.findall(f".//{qn('w:drawing')}") + child.findall(f".//{qn('w:pict')}")
                d_str = f" [DRAW:{len(drawings)}]" if drawings else ""
                kn_str = " [KN]" if child.find(f".//{qn('w:keepNext')}") is not None else ""
                kl_str = " [KL]" if child.find(f".//{qn('w:keepLines')}") is not None else ""
                if txt or num_str or d_str:
                    f.write(f"  P{idx:02d}{num_str}{d_str}{kn_str}{kl_str}: {txt[:90]}\n")
            elif tag == 'tbl':
                tbl = doc.tables[t_idx]
                rows = len(tbl.rows)
                cols = len(tbl.columns)
                tr_elms = child.findall(qn('w:tr'))
                kn_rows = [i for i, tr in enumerate(tr_elms) if tr.find(f".//{qn('w:keepNext')}") is not None]
                r0 = " | ".join([c.text.strip().replace('\n', ' ')[:25] for c in tbl.rows[0].cells[:6]])
                r_last = " | ".join([c.text.strip().replace('\n', ' ')[:25] for c in tbl.rows[-1].cells[:6]])
                f.write(f"  [TABLE #{t_idx:02d} ({rows}x{cols}) KN_rows={kn_rows}]:\n")
                f.write(f"    R0   : {r0}\n")
                if rows > 1:
                    f.write(f"    RLast: {r_last}\n")
                t_idx += 1
