import os
import docx

base_dir = r"c:\Users\phat.nt\Desktop\Source\Làm KQ"
out_path = r"C:\Users\phat.nt\.gemini\antigravity\brain\534de527-a676-4430-9dfa-9b3970a21af9\scratch\format_details.txt"

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
            doc = docx.Document(fpath)
            f.write(f"\n### FILE: {fname} ###\n")
            
            # Normal style font
            normal_style = doc.styles['Normal']
            f.write(f"Normal style font: name={normal_style.font.name}, size={normal_style.font.size.pt if normal_style.font.size else None}\n")
            
            # Inspect numbering on paragraphs
            for i, p in enumerate(doc.paragraphs):
                txt = p.text.strip()
                if not txt:
                    continue
                numPr = p._p.xpath('.//w:numPr')
                ilvl = p._p.xpath('.//w:numPr/w:ilvl/@w:val')
                numId = p._p.xpath('.//w:numPr/w:numId/@w:val')
                runs_info = []
                for r in p.runs:
                    if r.text.strip():
                        runs_info.append(f"'{r.text}'(b={r.bold},i={r.italic},u={r.underline},sz={r.font.size.pt if r.font.size else 'inh'},fn={r.font.name})")
                f.write(f"  P{i:02d} [style={p.style.name}, numId={numId}, ilvl={ilvl}]: {' + '.join(runs_info[:4])}\n")
            
            # Inspect tables borders and shading and column widths
            for t_i, tbl in enumerate(doc.tables):
                tblPr = tbl._tbl.tblPr
                borders = tblPr.xpath('.//w:tblBorders/*')
                b_str = ", ".join([f"{b.tag.split('}')[-1]}:{b.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')}" for b in borders])
                f.write(f"  T{t_i:02d} ({len(tbl.rows)}x{len(tbl.columns)}) borders=[{b_str}]\n")
                # Check first row and last row formatting
                for r_idx in [0, 1, len(tbl.rows)-1]:
                    if r_idx < len(tbl.rows):
                        row = tbl.rows[r_idx]
                        c0 = row.cells[0]
                        p0 = c0.paragraphs[0] if c0.paragraphs else None
                        r0 = p0.runs[0] if (p0 and p0.runs) else None
                        r_font = f"b={r0.bold},i={r0.italic},sz={r0.font.size.pt if r0 and r0.font.size else 'inh'}" if r0 else "no-run"
                        f.write(f"    Row {r_idx}: cell0='{c0.text.strip().replace(chr(10), ' ')[:40]}' ({r_font})\n")
