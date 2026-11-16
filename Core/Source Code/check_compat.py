import os
import docx
from docx.oxml.ns import qn

base_dir = r"c:\Users\phat.nt\Desktop\Source\Làm KQ"

def get_doc_defs(doc):
    styles = set(s.style_id for s in doc.styles)
    nums = set(n.get(qn('w:numId')) for n in doc.part.numbering_part.element.findall(qn('w:num')))
    return styles, nums

def get_used_defs(doc):
    p_styles = set(el.get(qn('w:val')) for el in doc.element.body.findall(f".//{qn('w:pStyle')}"))
    t_styles = set(el.get(qn('w:val')) for el in doc.element.body.findall(f".//{qn('w:tblStyle')}"))
    r_styles = set(el.get(qn('w:val')) for el in doc.element.body.findall(f".//{qn('w:rStyle')}"))
    nums = set(el.get(qn('w:val')) for el in doc.element.body.findall(f".//{qn('w:numId')}"))
    return p_styles | t_styles | r_styles, nums

for lang_26, lang_25 in [("viet - 1 - phong sach", "viet"), ("en - 1 - phong sach", "eng"), ("vn - 2 - khi nen", "viet - 4"), ("en - 2 - khi nen", "eng - 4")]:
    f26 = [f for f in os.listdir(os.path.join(base_dir, "New From 2026")) if lang_26 in f][0]
    doc26 = docx.Document(os.path.join(base_dir, "New From 2026", f26))
    s26, n26 = get_doc_defs(doc26)
    print(f"2026 Base: {f26}")
    for f25 in sorted(os.listdir(os.path.join(base_dir, "Old Form 2025"))):
        if lang_25 in f25 and f25.endswith(".docx"):
            doc25 = docx.Document(os.path.join(base_dir, "Old Form 2025", f25))
            u_s, u_n = get_used_defs(doc25)
            missing_s = u_s - s26
            missing_n = u_n - n26
            print(f"  vs 2025 {f25}: missing_styles={missing_s}, missing_nums={missing_n}")
