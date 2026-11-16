import os
import docx
from lxml import etree

base_dir = r"c:\Users\phat.nt\Desktop\Source\Làm KQ\New From 2026"
for fname in sorted(os.listdir(base_dir)):
    if not fname.endswith(".docx"):
        continue
    doc = docx.Document(os.path.join(base_dir, fname))
    print(f"=== {fname} ===")
    for i, sec in enumerate(doc.sections):
        print(f"  Section {i} different_first_page={sec.different_first_page_header_footer}")
        for p in sec.header.paragraphs:
            print(f"    Hdr P: {p.text}")
        for t in sec.header.tables:
            for r in t.rows:
                print(f"    Hdr Tbl Row: {[c.text.strip() for c in r.cells]}")
        for p in sec.footer.paragraphs:
            xml_str = etree.tostring(p._p, encoding='unicode')
            fields = [el.text for el in p._p.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}instrText')]
            print(f"    Ftr P: '{p.text}' fields={fields}")
        for t in sec.footer.tables:
            for r in t.rows:
                print(f"    Ftr Tbl Row: {[c.text.strip() for c in r.cells]}")
