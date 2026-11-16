# -*- coding: utf-8 -*-
"""
Cleanroom Document Handler (ISO 14644-1:2015) for QUATEST 3.

Sinh TBKQ phòng sạch (VN/EN) từ template dạng TAG "TBKQ - viet/en - phong sach.docx" (form tổng hợp đầy đủ).
Mục kết quả nào không có dữ liệu trong Dossier.results_data thì bị xóa cả mục (kèm bảng ghi chú, đánh giá)
theo AGENTS.md mục 3 "Chỉ tiêu không thực hiện".

Cấu trúc results_data (chi tiết + ví dụ: Core/Templates/TAGS.md, mục "TBKQ phòng sạch"):
    bui_3kenh / bui_6kenh : [ {cap, gioi_han[], trang_so_do, diem: [{khu_vuc, vi_tri, gia_tri[]}]} ]
    vi_sinh               : [ {cap, gioi_han, trang_so_do, diem: [{khu_vuc, vi_tri, gia_tri}]} ]
    nhiet_am              : {dong: [{khu_vuc, nhiet_do, do_am, do_on}], muc_nhiet_do, muc_do_am, muc_do_on}
    chenh_ap              : {dong: [{khu_vuc_1, khu_vuc_2, gia_tri, muc}]}
    hepa                  : [ {trang_so_do, cap_loc ("H13"/"H14", tùy chọn), muc ("≤ 0,1", tùy chọn),
                               diem: [{khu_vuc, vi_tri, gia_tri}]} ]   mặc định ≤ 0,01 %, H13 ≤ 0,1 %
    scan                  : {khu_vuc, so_do (đường dẫn ảnh), trang_so_do, diem: [{khu_vuc, vi_tri, hepa, gasket}]}
    luu_luong             : [ {muc_ach, trang_so_do, diem: [{khu_vuc, vi_tri, luu_luong}],
                               phong: {<khu_vuc>: {tong, the_tich, ach}}} ]
Mọi nhóm có thể kèm ket_luan_vn / ket_luan_en để ghi đè kết luận tự tính.
"""

import os

from ...Models.dossier import Dossier
from ...Styling import strip_all_highlights, add_document_ending_line, drop_section_gaps, strip_image_crops
from ...Tools.tag_engine import TagImage
from ..tag_context import (
    PLACEHOLDER,
    cover_context,
    dec,
    thousands,
    quoted,
    photo_grid,
    format_assessments,
    keep_intros_with_next,
    style_note_tables,
    ensure_result_tables,
    within,
    all_ok,
    verdict,
    find_tagged_tables,
    clear_table_tags,
    merge_by_groups,
    open_tag_template,
    save_outputs,
    render,
)


def _as_list(v):
    if not v:
        return []
    return v if isinstance(v, list) else [v]


HEPA_LEAK_LIMIT_DEFAULT = "≤ 0,01"                 # % — ISO 14644-3:2019, mặc định quy trình phòng sạch
HEPA_LEAK_LIMIT_BY_CLASS = {"H13": "≤ 0,1", "H14": "≤ 0,01"}


def hepa_leak_limit(group: dict) -> str:
    """Mức rò rỉ HEPA (%, dạng '≤ x' dấu phẩy) cho một nhóm kết quả: 'muc' ghi thẳng > 'cap_loc' > mặc định."""
    muc = str(group.get("muc") or "").strip()
    if muc:
        return muc if muc[:1] in "≤<≥>" else "≤ " + muc
    cap = str(group.get("cap_loc") or "").strip().upper().replace(" ", "")
    return HEPA_LEAK_LIMIT_BY_CLASS.get(cap, HEPA_LEAK_LIMIT_DEFAULT)


def _num_text(v, is_en: bool) -> str:
    """Số đếm hạt -> '5 368'; số thập phân -> dấu theo ngôn ngữ; chữ giữ nguyên (đổi dấu thập phân)."""
    if isinstance(v, int) or (isinstance(v, float) and v.is_integer() and abs(v) >= 1000):
        return thousands(v)
    return dec(v, is_en)


class CleanroomHandler:
    """Builder handler for ISO 14644-1 Cleanroom inspection reports (template dạng TAG)."""

    def __init__(self, dossier: Dossier, template_path_vn: str = None, template_path_en: str = None):
        self.dossier = dossier
        templates_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "Templates")
        vn_cand = os.path.join(templates_dir, "TBKQ - viet - phong sach.docx")
        if not os.path.isfile(vn_cand):
            vn_cand = os.path.join(templates_dir, "TBKQ - vn - phong sach.docx")
        self.template_vn = template_path_vn or vn_cand
        self.template_en = template_path_en or os.path.join(templates_dir, "TBKQ - en - phong sach.docx")

    # ------------------------------------------------------------------
    def build_context(self, is_en: bool) -> dict:
        d = self.dossier
        ctx = cover_context(d, is_en)
        rd = d.results_data or {}
        name_map = {loc.name_vn.strip(): (loc.name_en or loc.name_vn).strip() for loc in d.locations if loc.name_vn}

        def kv(name, en_name=None):
            name = (name or "").strip()
            if is_en:
                name = (en_name or name_map.get(name, name)).strip()
            return quoted(name)

        def kl(group, ok):
            return verdict(ok, is_en, group.get("ket_luan_en" if is_en else "ket_luan_vn"))

        def ds(names):
            seen = []
            for n in names:
                if n and n not in seen:
                    seen.append(n)
            return "; ".join(seen)

        # --- Nội dung thực hiện ---
        ctx["ds_khu_vuc"] = ds(quoted(loc.name_en if is_en else loc.name_vn) for loc in d.locations) or PLACEHOLDER
        first = d.locations[0] if d.locations else None
        ctx["trang_thai"] = ((first.state_en if is_en else first.state_vn) if first else "") or PLACEHOLDER
        ctx["khu_vuc"] = []
        for loc in d.locations:
            area = (loc.area_m2 or "").replace("m²", "").replace("m2", "").strip()
            cap = (loc.clean_class or "").strip()
            cap = cap[4:].strip() if cap.upper().startswith("ISO ") else cap
            lines = [ln if ln.strip().startswith("-") else f"- {ln.strip()}"
                     for ln in (loc.points_summary_en if is_en else loc.points_summary_vn) if ln.strip()]
            ctx["khu_vuc"].append({
                "ten": (loc.name_en if is_en else loc.name_vn) or "",
                "dien_tich": dec(area, is_en) or PLACEHOLDER,
                "cap": cap or PLACEHOLDER,
                "chi_tieu": lines or [PLACEHOLDER],
            })
        photos = [p for p in (d.photos or []) if os.path.isfile(p.get("path", "") if isinstance(p, dict) else str(p))]
        layouts = [p for p in (d.layouts or []) if os.path.isfile(p.get("path", "") if isinstance(p, dict) else str(p))]
        ctx["co_anh"], ctx["bang_anh"] = bool(photos), photo_grid(photos, is_en, captions=False,  # ảnh hiện trường không chú thích
                                                         # photo_cols: ép số cột (vd 3 ảnh xếp 1 hàng 3 cột để gọn trang)
                                                         cols=(d.results_data or {}).get("photo_cols"),
                                                         width_cm=(d.results_data or {}).get("photo_width_cm"))
        ctx["co_so_do"], ctx["bang_so_do"] = bool(layouts), photo_grid(
            # sơ đồ mặt bằng rộng -> results_data["layout_cols"] = 1 (mỗi sơ đồ một dòng, rộng layout_width_cm)
            layouts, is_en, cols=(d.results_data or {}).get("layout_cols") or (2 if len(layouts) >= 2 else 1),
            single_width_cm=float((d.results_data or {}).get("layout_width_cm") or 13.0))   # cỡ sơ đồ (AGENTS mục 3)

        # --- Nồng độ hạt bụi 3/6 kênh ---
        for key_rd, key_ctx, n in (("bui_3kenh", "bui3", 3), ("bui_6kenh", "bui6", 6)):
            groups = []
            for g in _as_list(rd.get(key_rd)):
                lim = list(g.get("gioi_han") or []) + [""] * n
                lim = [_num_text(x, is_en) for x in lim[:n]]
                diem, oks = [], []
                for p in g.get("diem", []):
                    vals = list(p.get("gia_tri") or []) + [""] * n
                    vals = vals[:n]
                    diem.append({"khu_vuc": kv(p.get("khu_vuc"), p.get("khu_vuc_en")), "vi_tri": str(p.get("vi_tri", "")),
                                 "v": [_num_text(v, is_en) for v in vals]})
                    oks += [within(v, "≤ " + l) if l else None for v, l in zip(vals, lim)]
                cap = str(g.get("cap", "")).strip()
                cap = cap[4:].strip() if cap.upper().startswith("ISO ") else cap
                groups.append({"cap": cap or PLACEHOLDER, "gioi_han": lim, "diem": diem,
                               "trang_so_do": str(g.get("trang_so_do") or PLACEHOLDER),
                               "ds_khu_vuc": ds(r["khu_vuc"] for r in diem), "ket_luan": kl(g, all_ok(oks))})
            ctx[key_ctx] = groups

        # --- Vi sinh ---
        ctx["vi_sinh"] = []
        for g in _as_list(rd.get("vi_sinh")):
            diem = [{"khu_vuc": kv(p.get("khu_vuc"), p.get("khu_vuc_en")), "vi_tri": str(p.get("vi_tri", "")),
                     "gia_tri": _num_text(p.get("gia_tri", ""), is_en)} for p in g.get("diem", [])]
            ctx["vi_sinh"].append({"cap": str(g.get("cap") or PLACEHOLDER), "gioi_han": _num_text(g.get("gioi_han") or PLACEHOLDER, is_en),
                                   "trang_so_do": str(g.get("trang_so_do") or PLACEHOLDER), "diem": diem})

        # --- Nhiệt độ - độ ẩm - độ ồn ---
        na = rd.get("nhiet_am")
        if na:
            muc = {k: dec(na.get(k, "-") or "-", is_en) for k in ("muc_nhiet_do", "muc_do_am", "muc_do_on")}
            dong, oks = [], []
            for r in na.get("dong", []):
                row = {"khu_vuc": kv(r.get("khu_vuc"), r.get("khu_vuc_en"))}
                for k, mk in (("nhiet_do", "muc_nhiet_do"), ("do_am", "muc_do_am"), ("do_on", "muc_do_on")):
                    row[k] = dec(r.get(k, "-"), is_en) if r.get(k) not in (None, "") else "-"
                    if row[k] != "-":
                        oks.append(within(row[k], muc[mk]))
                dong.append(row)
            ctx["nhiet_am"] = dict(muc, dong=dong, ds_khu_vuc=ds(r["khu_vuc"] for r in dong), ket_luan=kl(na, all_ok(oks)))
        else:
            ctx["nhiet_am"] = None

        # --- Chênh áp ---
        ca = rd.get("chenh_ap")
        if ca:
            dong, oks = [], []
            for r in ca.get("dong", []):
                row = {"khu_vuc_1": kv(r.get("khu_vuc_1"), r.get("khu_vuc_1_en")), "khu_vuc_2": kv(r.get("khu_vuc_2"), r.get("khu_vuc_2_en")),
                       "gia_tri": dec(r.get("gia_tri", ""), is_en), "muc": dec(r.get("muc", "-"), is_en)}
                oks.append(within(row["gia_tri"], row["muc"]))
                dong.append(row)
            ctx["chenh_ap"] = {"dong": dong, "ket_luan": kl(ca, all_ok(oks))}
        else:
            ctx["chenh_ap"] = None

        # --- Độ rò rỉ lọc HEPA ---
        # Mức (ISO 14644-3:2019): mặc định ≤ 0,01 %; lọc H13 → ≤ 0,1 % (người dùng chốt 30/10/2026).
        # Mỗi nhóm có thể ghi "cap_loc": "H13"/"H14" hoặc ghi thẳng "muc": "≤ 0,1".
        ctx["hepa"] = []
        for g in _as_list(rd.get("hepa")):
            diem = [{"khu_vuc": kv(p.get("khu_vuc"), p.get("khu_vuc_en")), "vi_tri": str(p.get("vi_tri", "")),
                     "gia_tri": dec(p.get("gia_tri", ""), is_en)} for p in g.get("diem", [])]
            muc = dec(hepa_leak_limit(g), is_en)
            ok = all_ok(within(p["gia_tri"], muc) for p in diem)
            ctx["hepa"].append({"diem": diem, "muc": muc, "trang_so_do": str(g.get("trang_so_do") or PLACEHOLDER),
                                "ds_khu_vuc": ds(p["khu_vuc"] for p in diem), "ket_luan": kl(g, ok)})

        # --- Quét rò rỉ HEPA (Scanloc) ---
        sc = rd.get("scan")
        if sc:
            ctx["scan"] = {
                "khu_vuc": kv(sc.get("khu_vuc"), sc.get("khu_vuc_en")).strip("“”"),
                "so_do": TagImage(sc.get("so_do", ""), 15.0) if sc.get("so_do") else "",
                "trang_so_do": str(sc.get("trang_so_do") or PLACEHOLDER),
                "diem": [{"khu_vuc": kv(p.get("khu_vuc"), p.get("khu_vuc_en")), "vi_tri": str(p.get("vi_tri", "")),
                          "hepa": dec(p.get("hepa", ""), is_en), "gasket": dec(p.get("gasket", ""), is_en)}
                         for p in sc.get("diem", [])],
            }
        else:
            ctx["scan"] = None

        # --- Lưu lượng / ACH ---
        ctx["luu_luong"] = []
        for g in _as_list(rd.get("luu_luong")):
            phong = g.get("phong", {}) or {}
            muc = _num_text(g.get("muc_ach") or PLACEHOLDER, is_en)
            diem, oks = [], []
            for p in g.get("diem", []):
                info = phong.get(p.get("khu_vuc", ""), {})
                diem.append({"khu_vuc": kv(p.get("khu_vuc"), p.get("khu_vuc_en")), "vi_tri": str(p.get("vi_tri", "")),
                             "luu_luong": _num_text(p.get("luu_luong", ""), is_en),
                             "tong": _num_text(info.get("tong", ""), is_en),
                             "the_tich": _num_text(info.get("the_tich", ""), is_en),
                             "ach": _num_text(info.get("ach", ""), is_en)})
            for name, info in phong.items():
                oks.append(within(info.get("ach", ""), "≥ " + str(g.get("muc_ach", ""))) if g.get("muc_ach") else None)
            ctx["luu_luong"].append({"diem": diem, "muc_ach": muc, "trang_so_do": str(g.get("trang_so_do") or PLACEHOLDER),
                                     "ds_khu_vuc": ds(p["khu_vuc"] for p in diem), "ket_luan": kl(g, all_ok(oks))})
        return ctx

    # ------------------------------------------------------------------
    @staticmethod
    def _merge_rooms(doc):
        """Gộp dọc ô 'Khu vực' (và tổng lưu lượng/thể tích/ACH) cho các điểm đo cùng phòng."""
        spec = {  # tên bảng: (số dòng tiêu đề, có dòng giới hạn ở cuối, các cột gộp)
            "tag:bui3": (2, True, (0,)), "tag:bui6": (2, True, (0,)), "tag:vi_sinh": (2, True, (0,)),
            "tag:hepa": (1, True, (0,)), "tag:scan": (1, False, (0,)), "tag:luu_luong": (1, True, (0, 3, 4, 5)),
        }
        for name, (hdr, has_lim, cols) in spec.items():
            for t in find_tagged_tables(doc, name):
                end = len(t.rows) - (2 if has_lim else 1)
                if end > hdr:
                    merge_by_groups(t, hdr, end, key_col=0, cols=cols)
        clear_table_tags(doc)

    def build_tbkq(self, is_en: bool = False, **kwargs) -> str:
        """Builds complete TBKQ document (VN or EN) for Cleanroom from the tagged template."""
        d = self.dossier
        doc = open_tag_template(self.template_en if is_en else self.template_vn)
        render(doc, self.build_context(is_en), label=f"TBKQ phòng sạch {'EN' if is_en else 'VN'}")
        self._merge_rooms(doc)
        ensure_result_tables(doc)
        format_assessments(doc)
        keep_intros_with_next(doc)
        style_note_tables(doc)          # bảng Ghi chú 11pt in nghiêng (09/11/2026)
        drop_section_gaps(doc)          # đoạn trống sót giữa các mục (mục bị xóa khỏi template) — AGENTS mục 3
        add_document_ending_line(doc)
        strip_all_highlights(doc)
        strip_image_crops(doc)

        out_name = f"TBKQ - {'en' if is_en else 'vn'} - {d.sub_no or 1} - phong sach - {d.code}.docx"
        target = save_outputs(doc, d, d.tbkq_dir, out_name, label="cleanroom.py")
        print(f"[CleanroomHandler] Successfully generated TBKQ {'EN' if is_en else 'VN'} -> {target}")
        return target
