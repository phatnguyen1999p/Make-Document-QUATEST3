# -*- coding: utf-8 -*-
"""Đọc Phiếu yêu cầu & nhận thử nghiệm (BM101.01, PDF có lớp chữ — `PYC-Q326D<6 số>.N2.n.pdf`) -> danh sách mẫu,
chỉ tiêu, phương pháp; ngày nhận, ngày dự kiến trả kết quả (= "Ngày yêu cầu cấp KQTN", AGENTS mục 13).

    from Core.Tools.pyc_reader import read_pyc
    p = read_pyc(path)  # dict: ngay_nhan, ngay_tra, ma_rod, samples=[{ten, ma_qt, params=[(tên, phương pháp)]}]

Dùng khi chưa có KQTN để dựng khung chứng thư (tên chỉ tiêu / phương pháp theo PYC; khi KQTN về thì ghi theo KQTN).
Xem nhanh: ``python Core/Tools/pyc_reader.py <file.pdf>``.
"""
import re
import sys
import unicodedata

import pymupdf

METHOD_START = ("SMEWW", "TCVN", "US EPA", "EPA", "ISO", "ASTM", "QCVN", "Đo ", "Cảm quan", "HD", "AOAC", "APHA",
                "ISO/", "BS ", "EN ", "DIN", "JIS", "NIOSH", "OSHA", "TK ", "H ")
PRICE = re.compile(r"^\d{1,3}(\.\d{3})+\s*(₫|đ)?$")


def _norm(s):
    return re.sub(r"\s+", " ", unicodedata.normalize("NFC", s or "")).strip()


def read_pyc(path):
    doc = pymupdf.open(path)
    text = "\n".join(p.get_text() for p in doc)
    info = {"path": path}
    m = re.search(r"Ngày nhận:\s*(\d{2}/\d{2}/\d{4})", text)
    info["ngay_nhan"] = m.group(1) if m else ""
    m = re.search(r"Ngày dự kiến trả kết quả:\s*(\d{2}/\d{2}/\d{4})", text)
    info["ngay_tra"] = m.group(1) if m else ""
    m = re.search(r"Mã ROD Main:\s*(\w+)", text)
    info["ma_rod"] = m.group(1) if m else ""
    lines = [_norm(x) for x in text.split("\n")]
    hdr = ("TRUNG TÂM KỸ THUẬT", "Trụ sở chính", "Phòng Dịch vụ Khách hàng", "Lô C5", "1196 Đường", "Số 31 đường",
           "BM101", "HL:", "Trang:")
    lines = [x for x in lines if x and not x.startswith(hdr)]
    samples, cur, i = [], None, 0
    while i < len(lines):
        ln = lines[i]
        if ln.startswith("Tên mẫu(1)"):
            j = i + 1
            name = lines[j] if j < len(lines) else ""
            cur = {"ten": name, "ma_qt": "", "params": []}
            samples.append(cur)
            i = j + 1
            continue
        if cur is not None and "Mã QT mẫu:" in ln:
            m = re.search(r"Mã QT mẫu:\s*(\S+?)\.*$", ln) or re.search(r"Mã QT mẫu:\s*([\w.]+)", ln)
            cur["ma_qt"] = re.sub(r"\.+$", "", m.group(1)) if m else ""
        if cur is not None and re.fullmatch(r"\d{1,2}", ln):
            # một chỉ tiêu: STT, tên (1+ dòng), phương pháp (1+ dòng), đơn giá
            k, chunk = i + 1, []
            while k < len(lines) and not PRICE.match(lines[k]) and k - i < 12:
                if re.fullmatch(r"\d{1,2}", lines[k]):
                    break
                chunk.append(lines[k])
                k += 1
            if k < len(lines) and PRICE.match(lines[k]) and chunk:
                mi = next((n for n, x in enumerate(chunk) if n > 0 and x.startswith(METHOD_START)), len(chunk))
                name = " ".join(chunk[:mi])
                meth = " ".join(chunk[mi:])
                meth = re.sub(r"3500 Cr B", "3500-Cr B", meth)
                meth = re.sub(r"4500 ?S²⁻ ?([A-Z])", r"4500-S²⁻ \1", meth)
                meth = re.sub(r"(\d)\s*-\s*(\d)", r"\1-\2", meth)
                cur["params"].append((name, meth))
                i = k + 1
                continue
        i += 1
    info["samples"] = samples
    return info


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    for p in sys.argv[1:]:
        k = read_pyc(p)
        print(k["ma_rod"], "nhận", k["ngay_nhan"], "dự kiến trả", k["ngay_tra"])
        for s in k["samples"]:
            print(f"  [{s['ten']}] {s['ma_qt']} — {len(s['params'])} chỉ tiêu")
            for n, m in s["params"]:
                print(f"      {n:60} | {m}")
