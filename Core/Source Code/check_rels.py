import os
import docx
from lxml import etree

base_dir = r"c:\Users\phat.nt\Desktop\Source\Làm KQ"
for folder in ["New From 2026", "Old Form 2025"]:
    folder_path = os.path.join(base_dir, folder)
    print(f"=== {folder} ===")
    for fname in sorted(os.listdir(folder_path)):
        if not fname.endswith(".docx") or fname.startswith("~$"):
            continue
        doc = docx.Document(os.path.join(folder_path, fname))
        print(f"File: {fname}")
        for t_i, tbl in enumerate(doc.tables):
            xml_str = etree.tostring(tbl._tbl, encoding='unicode')
            has_rel = 'r:embed' in xml_str or 'r:id' in xml_str or 'r:link' in xml_str
            if has_rel:
                print(f"  Table #{t_i} has r:* relationship!")
