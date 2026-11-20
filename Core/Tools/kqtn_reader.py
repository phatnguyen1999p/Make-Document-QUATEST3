# -*- coding: utf-8 -*-
"""Đọc phiếu Kết quả thử nghiệm (BM105.02, PDF ký điện tử có lớp chữ) -> danh sách chỉ tiêu.

    from Core.Tools.kqtn_reader import read_kqtn
    kq = read_kqtn(path)   # dict: so_phieu, ngay, ten_mau, ngay_nhan_mau, rows=[KqRow…]
    kqs = read_kqtn_all(path)   # file gộp nhiều phiếu (.01, .02 …) -> list dict, mỗi phiếu một dict

Bảng (STT | Tên chỉ tiêu | Đơn vị | Phương pháp | Kết quả | Giới hạn phát hiện) lấy khung ô bằng
``page.find_tables()``, chữ từng ô bằng ``page.get_textbox`` (giữ đúng dấu phẩy, khoảng trắng); hàng có STT ``7.n`` /
``7.n.m`` (cấu tử của nhóm chất). Ký hiệu ``⁽¹⁾`` (VIMCERTS) bị bỏ khỏi tên, giữ ở ``marks``.
Xem nhanh: ``python Core/Tools/kqtn_reader.py <file.pdf>``.
"""
import re
import sys
import unicodedata
from dataclasses import dataclass, field

import pymupdf

HEAD_KEYS = ("STT", "Tên", "Đơn", "Phương", "Kết", "Giới")


@dataclass
class KqRow:
    stt: str
    name: str
    unit: str
    method: str
    result: str
    lod: str
    level: int = 0                       # 0 = chỉ tiêu, 1 = cấu tử (7.26.1), 2 = cấu tử con (7.27.3.1)
    marks: list = field(default_factory=list)

    @property
    def kph(self):
        return self.result.upper().startswith("KPH")

    def ket_qua(self):
        """Quy tắc từ 19/11/2026 (người dùng, hồ sơ 131873): KPH -> ghi GIỚI HẠN PHÁT HIỆN, KHÔNG dấu "<"
        ('KPH', LOD 0,5 -> '0,5'); còn lại như value() ('< 0,6' kèm loq=True khi LOQ). Trả (chuỗi, loq, kph).
        Đánh giá đạt / không đạt phải xét cờ kph (KPH luôn coi là dưới mức), không so số LOD với mức."""
        if self.kph and self.lod not in ("", "-"):
            return self.lod, False, True
        v, loq = self.value()
        return v, loq, self.kph

    def value(self):
        """Kết quả ghi lên chứng thư: KPH -> '< LOD'; '＜ 0,01 (LOQ)' -> ('< 0,01', loq=True)."""
        r = self.result.replace("＜", "<").strip()
        if self.kph:
            return (f"< {self.lod}", False) if self.lod not in ("", "-") else ("KPH", False)
        m = re.match(r"^<\s*([\d.,]+)\s*\(LOQ\)", r)
        if m:
            return f"< {m.group(1)}", True
        return r, False


def _norm(s):
    s = unicodedata.normalize("NFC", s or "").replace("Ð", "Đ").replace("ð", "đ")   # Ð (Eth) -> Đ (137037)
    return re.sub(r"\s+", " ", s).strip()


def read_kqtn_all(path):
    """Một file có thể gộp nhiều phiếu (134486: `KQ-…-MTTN26020478 (M4-M12).pdf` chứa .04 – .12) -> tách theo số phiếu
    đuôi `.nn` in trên từng trang, đọc từng phiếu; file một phiếu -> [read_kqtn(path)]."""
    doc = pymupdf.open(path)
    groups = {}
    for i, page in enumerate(doc):
        m = re.search(r"[A-Z]{2,5}TN\d+\.(\d+)", page.get_text())
        key = m.group(1) if m else (list(groups)[-1] if groups else "")
        groups.setdefault(key, []).append(i)
    if len(groups) <= 1:
        return [_parse(doc, path)]
    out = []
    for pages in groups.values():
        sub = pymupdf.open()
        for i in pages:
            sub.insert_pdf(doc, from_page=i, to_page=i)
        out.append(_parse(sub, path))
    return out


def read_kqtn(path):
    return _parse(pymupdf.open(path), path)


def _parse(doc, path):
    first = doc[0].get_text()
    info = {"path": path}
    m = re.search(r"\((MTTN\d+(?:\.\d+)?|HDTN\d+(?:\.\d+)?|[A-Z]{2,5}TN\d+(?:\.\d+)?)\)", first)   # phiếu 1 mẫu: không đuôi .01 (137037)
    info["so_phieu"] = m.group(1) if m else ""
    # mẫu BM105.02.VIL (phiếu vi sinh, 142705): ngày nằm sau số phiếu, nhãn "1." / "Tên mẫu" / ":" tách dòng
    m = re.search(r"T E S T R E P O R T[ \t]*\n[ \t]*(\d{2}/\d{2}/\d{4})", first) or \
        re.search(r"\([A-Z]{2,5}TN[\d.]+\)[ \t]*\n(?:[ \t]*\([^)\n]*\)[ \t]*\n)?[ \t]*(\d{2}/\d{2}/\d{4})", first)   # (MT.…/n) xen giữa
    info["ngay"] = m.group(1) if m else ""
    # số niêm phong (131873: "- Niêm phong: niêm nhựa số KT3 72818" -> "KT372818"); đối chiếu BBLM khi ghép mẫu
    m = re.search(r"Niêm phong:[^\n]*?số[ \t]*([A-Z]{1,3}[ \t]*\d[\d \t]*\d)", first)
    info["niem"] = re.sub(r"\s+", "", m.group(1)) if m else ""
    m = re.search(r"1\. Tên mẫu\s*\n\s*:\s*(.+)", first) or re.search(r"Tên mẫu\s*\n\s*:\s*\n\s*(.+)", first)
    info["ten_mau"] = _norm(m.group(1)) if m else ""
    m = re.search(r"Thời gian lấy mẫu:\s*(\d{2}/\d{2}/\d{4})", first)
    info["ngay_lay_mau"] = m.group(1) if m else ""
    m = re.search(r"4\. Ngày nhận mẫu\s*\n\s*:\s*(\d{2}/\d{2}/\d{4})", first) or \
        re.search(r"Ngày nhận mẫu\s*\n\s*:\s*\n\s*(\d{2}/\d{2}/\d{4})", first)
    info["ngay_nhan_mau"] = m.group(1) if m else ""
    m = re.search(r"5\. Thời gian thử nghiệm\s*\n\s*:\s*([\d/]+)\s*-\s*([\d/]+)", first) or \
        re.search(r"Thời gian thử nghiệm\s*\n\s*:\s*\n\s*([\d/]+)\s*-\s*([\d/]+)", first)
    info["thu_nghiem"] = (m.group(1), m.group(2)) if m else ("", "")
    rows = []
    for page in doc:
        for tab in page.find_tables().tables:
            # dòng tiêu đề STT thường là hàng đầu; phiếu nhiều trang có khi bị gộp trong khung cả trang (134486 N4:
            # bảng 9 cột, hàng đầu là tiêu đề trang) -> dò hàng / cột của từng tiêu đề
            idx, start = None, 0
            for h, tr in enumerate(tab.rows):
                texts = [page.get_textbox(c).strip() if c else "" for c in tr.cells]
                pos = [next((i for i, t in enumerate(texts) if t.startswith(k)), None) for k in HEAD_KEYS]
                if pos[0] is not None and all(x is not None for x in pos[:5]):   # VIL: không có cột giới hạn
                    idx, start = pos, h + 1
                    break
            if idx is None:
                continue
            for tr in tab.rows[start:]:
                cells = [_norm(page.get_textbox(c)) if c else "" for c in tr.cells]
                vals = [cells[i] if i is not None and i < len(cells) else "" for i in idx]
                if not re.match(r"^7\.", vals[0].replace(" ", "")):
                    continue
                rows.append({"stt": vals[0], "cols": vals[1:6]})
    out = []
    for r in rows:
        stt = re.sub(r"\s+", "", r["stt"]).rstrip(".")
        name, unit, meth, res, lod = (_norm(c) for c in r["cols"])
        marks = re.findall(r"⁽([\d¹²³⁴⁵⁶⁷⁸⁹])⁾", name)
        name = _norm(re.sub(r"⁽[\d¹²³⁴⁵⁶⁷⁸⁹]⁾", "", name))
        meth = re.sub(r"(\d)\s*-\s*(\d)", r"\1-\2", meth)          # "6001- 1:2021"
        meth = meth.replace("3500 Cr B", "3500-Cr B").replace("4500 S²⁻B", "4500-S²⁻ B")
        unit = unit.replace("MPN/100m L", "MPN/100mL").replace(" ", "") if "MPN" in unit else unit
        out.append(KqRow(stt=stt, name=name, unit=unit, method=meth, result=res.replace("＜", "<"), lod=lod,
                         level=stt.count(".") - 1, marks=marks))
    info["rows"] = out
    return info


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    for p in sys.argv[1:]:
        for k in read_kqtn_all(p):
            print(k["so_phieu"], k["ngay"], k["ten_mau"], "nhận", k["ngay_nhan_mau"], "lấy", k["ngay_lay_mau"])
            for r in k["rows"]:
                print(f"  {'  ' * r.level}{r.stt:10} {r.name:50} | {r.unit:10} | {r.method:35} | {r.result:15} | {r.lod}")
