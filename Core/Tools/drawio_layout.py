# -*- coding: utf-8 -*-
"""
Sơ đồ vị trí đo dạng draw.io — người dùng kéo thả lại điểm đo khi code đặt chưa đúng.

Dùng công cụ AutoAddImage của người dùng (`DrawIOBuilder.add_image_tab` thêm ảnh thành tab,
`DrawIOExporter.export_all_tabs` xuất mỗi tab một ảnh) — thư mục mặc định
`C:\\Users\\phat.nt\\Desktop\\Source\\Source Code\\DrawIO Tool\\AutoAddImage`, đổi bằng biến môi trường `DRAWIO_TOOL_DIR`.

File `<thư mục sơ đồ>/so_do_vi_tri.drawio` có 2 tab `so_do_vi_tri_vn`, `so_do_vi_tri_en`; mỗi tab 3 lớp:
  - "Mặt bằng (khóa)" : ảnh mặt bằng (đã xóa chữ viết tay) + khung viền — khóa, không kéo nhầm;
  - "Vị trí đo"       : ký hiệu điểm đo (+1, ②, PA1, HEPA 1…) — mỗi ký hiệu một ô kéo thả được;
  - "Chú thích"       : tên phòng + bảng chú thích theo ngôn ngữ của tab.

Quy trình:
  1. Lần dựng đầu code tạo file .drawio từ tọa độ trong script hồ sơ.
  2. Người dùng mở file bằng draw.io → **chỉ sửa tab `_vn`** (kéo điểm, sửa nhãn) → Ctrl+S.
  3. Chạy lại script hồ sơ: file đã có thì **không dựng lại**; lớp "Vị trí đo" của tab `_en` được đồng bộ theo
     tab `_vn` (vị trí, chữ, kiểu), rồi xuất `so_do_vi_tri_vn.png` / `so_do_vi_tri_en.png` (tên ảnh = tên tab).
     Lớp "Chú thích" của mỗi tab giữ riêng (tên phòng/chú thích EN sửa trực tiếp trên tab `_en`).
  Dựng lại từ tọa độ trong code (bỏ hết chỉnh tay): `overwrite=True` (cờ `--tao-lai-so-do` của script hồ sơ).
Chữ trong hình Times New Roman (AGENTS 8C).
"""

from __future__ import annotations

import base64
import os
import sys
import tempfile
import urllib.parse
import xml.etree.ElementTree as ET
import zlib
from xml.sax.saxutils import quoteattr

def _find_tool_dir() -> str:
    """Biến DRAWIO_TOOL_DIR → bản gốc trên máy → bản chép kèm `CongCu\\AutoAddImage` (11/11/2026)."""
    env = os.environ.get("DRAWIO_TOOL_DIR")
    if env:
        return env
    workspace = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    for c in (os.path.join(os.path.expanduser("~"), "Desktop", "Source", "Source Code", "DrawIO Tool", "AutoAddImage"),
              os.path.join(workspace, "CongCu", "AutoAddImage")):
        if os.path.isfile(os.path.join(c, "drawio_builder.py")):
            return c
    return os.path.join(workspace, "CongCu", "AutoAddImage")


TOOL_DIR = _find_tool_dir()
FONT = "Times New Roman"
BASENAME = "so_do_vi_tri"
LAYER_POINTS, LAYER_TEXT = "vitri", "chuthich"
LAYER_LABELS = {"base": "Mặt bằng (khóa)", LAYER_POINTS: "Vị trí đo (kéo thả)", LAYER_TEXT: "Chú thích"}

# Kiểu ký hiệu — khớp màu/cỡ của bản PIL cũ (132809, 138143)
PRESETS = {
    "dust": dict(shape="text", color="#C80000", font_size=54, bold=True),          # số vị trí bụi, chữ đỏ
    "point": dict(shape="ellipse", color="#1428AA", stroke="#1428AA", stroke_width=6, font_size=44, bold=True,
                  w=68, h=68),                                                     # số trong vòng tròn xanh
    "dp": dict(shape="text", color="#006E28", font_size=40, bold=True, fill="#FFFFFF"),     # PA1, PA2
    "hepa": dict(shape="text", color="#282828", font_size=22, bold=True, fill="#FFFFFF"),   # HEPA 1…
    "room": dict(shape="text", color="#000000", font_size=40, bold=True, fill="#FFFFFF"),   # tên phòng
}


def _tool():
    """Nạp DrawIOBuilder, DrawIOExporter từ AutoAddImage của người dùng."""
    if not os.path.isfile(os.path.join(TOOL_DIR, "drawio_builder.py")):
        raise SystemExit(f"[FATAL] Không thấy công cụ AutoAddImage tại {TOOL_DIR} (đặt biến DRAWIO_TOOL_DIR).")
    if TOOL_DIR not in sys.path:
        sys.path.insert(0, TOOL_DIR)
    from drawio_builder import DrawIOBuilder
    from drawio_exporter import DrawIOExporter
    return DrawIOBuilder, DrawIOExporter


# ---------------------------------------------------------------------------
# Ô (cell) draw.io
# ---------------------------------------------------------------------------
def _text_w(text: str, size: float) -> int:
    return int(max(len(text), 1) * size * 0.62 + 16)


def _style(spec: dict) -> str:
    shape = spec.get("shape", "text")
    s = {"html": "0", "fontFamily": FONT, "fontSize": spec.get("font_size", 34),
         "fontColor": spec.get("color", "#000000"), "align": spec.get("align", "center"),
         "verticalAlign": "middle", "spacing": "0"}
    if spec.get("bold"):
        s["fontStyle"] = "1"
    if shape == "text":
        s = {"text": None, **s, "fillColor": spec.get("fill", "none"), "strokeColor": "none"}
    elif shape == "ellipse":
        s = {"ellipse": None, **s, "fillColor": spec.get("fill", "none"), "strokeColor": spec.get("stroke", "#000000"),
             "strokeWidth": spec.get("stroke_width", 4)}
    elif shape == "rect":
        s = {"rounded": "0", **s, "fillColor": spec.get("fill", "none"),
             "strokeColor": spec.get("stroke", "#000000"), "strokeWidth": spec.get("stroke_width", 4)}
    s.update(spec.get("extra", {}))
    return "".join(f"{k};" if v is None else f"{k}={v};" for k, v in s.items())


def _vertex(cid, parent, x, y, w, h, style, value=""):
    return ET.fromstring(
        f'<mxCell id="{cid}" value={quoteattr(value)} style={quoteattr(style)} vertex="1" parent="{parent}">'
        f'<mxGeometry x="{x:.1f}" y="{y:.1f}" width="{w:.1f}" height="{h:.1f}" as="geometry"/></mxCell>')


def _arrow(cid, parent, x0, y0, x1, y1, color, width=5):
    return ET.fromstring(
        f'<mxCell id="{cid}" value="" style="endArrow=block;endFill=1;html=0;strokeColor={color};'
        f'strokeWidth={width};" edge="1" parent="{parent}"><mxGeometry relative="1" as="geometry">'
        f'<mxPoint x="{x0}" y="{y0}" as="sourcePoint"/><mxPoint x="{x1}" y="{y1}" as="targetPoint"/>'
        f'</mxGeometry></mxCell>')


def _item(cid, parent, spec: dict, cx: float, cy: float, text: str = ""):
    """Ký hiệu có tâm (cx, cy). spec: preset + ghi đè (shape text/ellipse/rect/arrow, color, font_size, w, h…)."""
    spec = {**PRESETS.get(spec.get("preset", ""), {}), **spec}
    if spec.get("shape") == "arrow":
        L, c, sw = spec.get("w", 50), spec.get("color", "#000000"), spec.get("stroke_width", 5)
        if spec.get("dir", "right") == "down":
            return _arrow(cid, parent, cx, cy - L / 2, cx, cy + L / 2, c, sw)
        return _arrow(cid, parent, cx - L / 2, cy, cx + L / 2, cy, c, sw)
    fs = spec.get("font_size", 34)
    w = spec.get("w") or _text_w(text, fs)
    h = spec.get("h") or int(fs * 1.25 + 6)
    return _vertex(cid, parent, cx - w / 2, cy - h / 2, w, h, _style(spec), text)


def _layer(lid, label, locked=False):
    st = ' style="locked=1;"' if locked else ""
    return ET.fromstring(f'<mxCell id="{lid}" value={quoteattr(label)}{st} parent="0"/>')


# ---------------------------------------------------------------------------
# Dựng file .drawio
# ---------------------------------------------------------------------------
def build_layout_drawio(drawio_path: str, plan, markers: list[dict], rooms: dict | None = None,
                        legend: dict | None = None, legend_cols: dict | None = None,
                        legend_font_size: int = 34, row_h: int = 52, pad: int = 30, frame: bool = True):
    """Tạo (ghi đè) file .drawio 2 tab VN/EN.

    plan    : PIL.Image mặt bằng (đã xóa chữ viết tay) — mọi tọa độ tính theo pixel ảnh này.
    markers : [{"preset": "dust"|"point"|"dp"|"hepa", "text": "+1", "x": cx, "y": cy, ...ghi đè kiểu}]
    rooms   : {"vn": [(cx, cy, "PHÒNG SẠCH ISO 5"), ...], "en": [...]}
    legend  : {"vn": [(symbol_spec, "Vị trí đo ..."), ...], "en": [...]} — symbol_spec như marker (+ "text" mẫu).
    legend_cols: {"vn": 2, "en": 1}
    """
    DrawIOBuilder, _ = _tool()
    os.makedirs(os.path.dirname(drawio_path), exist_ok=True)
    rooms, legend, legend_cols = rooms or {}, legend or {}, legend_cols or {}
    W, H = plan.size
    with tempfile.TemporaryDirectory(prefix="drawio_") as tmp:
        plan_png = os.path.join(tmp, "mat_bang.png")
        plan.convert("RGB").save(plan_png, "PNG", optimize=True)
        builder = DrawIOBuilder()
        for lang in ("vn", "en"):
            builder.add_image_tab(plan_png, tab_name=f"{BASENAME}_{lang}", embed=True, margin=0)
    for diagram in builder.root.findall("diagram"):
        lang = diagram.get("name").rsplit("_", 1)[-1]
        model = diagram.find("mxGraphModel")
        model.set("grid", "0")
        model.set("background", "#FFFFFF")
        root = model.find("root")
        base = root.find("mxCell[@id='1']")                       # lớp mặc định chứa ảnh -> khóa
        base.set("value", LAYER_LABELS["base"])
        base.set("style", "locked=1;")
        img = next(c for c in root.findall("mxCell") if "shape=image" in (c.get("style") or ""))
        img.set("id", "matbang")
        if frame:
            root.append(_vertex("khung", "1", 1, 1, W - 2, H - 2,
                                "rounded=0;html=0;fillColor=none;strokeColor=#000000;strokeWidth=2;"))
        root.append(_layer(LAYER_POINTS, LAYER_LABELS[LAYER_POINTS]))
        for i, m in enumerate(markers, 1):
            root.append(_item(f"pt{i}", LAYER_POINTS, m, m["x"], m["y"], m.get("text", "")))
        root.append(_layer(LAYER_TEXT, LAYER_LABELS[LAYER_TEXT]))
        labels, cols = legend.get(lang, []), legend_cols.get(lang, 1)
        rows = (len(labels) + cols - 1) // cols
        if labels:                                   # nền trắng vùng chú thích
            root.append(_vertex("nen_cj", LAYER_TEXT, 0, H, W, pad + rows * row_h + pad,
                                "rounded=0;html=0;fillColor=#FFFFFF;strokeColor=none;"))
        for i, (cx, cy, t) in enumerate(rooms.get(lang, []), 1):
            root.append(_item(f"phong{i}", LAYER_TEXT, {"preset": "room"}, cx, cy, t))
        for i, (sym, lab) in enumerate(labels):
            x0 = 60 + (i % cols) * (W // cols)
            cy = H + pad + (i // cols) * row_h + row_h // 2
            root.append(_item(f"cj_kh{i + 1}", LAYER_TEXT, sym, x0 + 18, cy, sym.get("text", "")))
            root.append(_vertex(f"cj_chu{i + 1}", LAYER_TEXT, x0 + 80, cy - row_h / 2, W // cols - 150, row_h,
                                _style({"shape": "text", "font_size": legend_font_size, "align": "left"}), lab))
    builder.save(drawio_path)
    print(f"[drawio] Đã tạo {drawio_path} — mở bằng draw.io, kéo lại điểm ở tab _vn nếu sai, Ctrl+S, chạy lại.")


# ---------------------------------------------------------------------------
# Đồng bộ tab EN theo tab VN + xuất PNG
# ---------------------------------------------------------------------------
def _model(diagram: ET.Element) -> ET.Element:
    """mxGraphModel của tab; tab nén (deflate+base64+urlencode) thì giải nén và gắn lại dạng không nén."""
    model = diagram.find("mxGraphModel")
    if model is None:
        raw = zlib.decompress(base64.b64decode((diagram.text or "").strip()), -15)
        model = ET.fromstring(urllib.parse.unquote(raw.decode("utf-8")))
        diagram.text = None
        diagram.append(model)
    return model


def _layer_id(root: ET.Element, lid: str, label: str) -> str | None:
    """Id lớp theo id gốc, hoặc theo tên (người dùng có thể đã tạo lại lớp)."""
    for c in root.findall("mxCell"):
        if c.get("parent") == "0" and (c.get("id") == lid or (c.get("value") or "").strip() == label):
            return c.get("id")
    return None


def sync_points(drawio_path: str) -> bool:
    """Chép lớp "Vị trí đo" của tab _vn sang tab _en (thay toàn bộ). Trả True nếu file đổi."""
    tree = ET.parse(drawio_path)
    tabs = {d.get("name", "").rsplit("_", 1)[-1].lower(): d for d in tree.getroot().findall("diagram")}
    if "vn" not in tabs or "en" not in tabs:
        print("[drawio] Không thấy đủ 2 tab _vn/_en — bỏ qua đồng bộ.")
        return False
    vn_root, en_root = _model(tabs["vn"]).find("root"), _model(tabs["en"]).find("root")
    label = LAYER_LABELS[LAYER_POINTS]
    vn_layer, en_layer = _layer_id(vn_root, LAYER_POINTS, label), _layer_id(en_root, LAYER_POINTS, label)
    if not vn_layer or not en_layer:
        print("[drawio] Không thấy lớp 'Vị trí đo' ở cả 2 tab — bỏ qua đồng bộ.")
        return False

    def subtree(root, layer):
        ids, out = {layer}, []
        for c in root.findall("mxCell"):              # thứ tự trong file: cha đứng trước con
            if c.get("parent") in ids:
                ids.add(c.get("id"))
                out.append(c)
        return out

    vn_cells, en_cells = subtree(vn_root, vn_layer), subtree(en_root, en_layer)
    def sig(cells):                                    # so nội dung, bỏ khoảng trắng định dạng (tail)
        out = []
        for c in cells:
            c = ET.fromstring(ET.tostring(c))
            for e in c.iter():
                e.tail = None
                if e.text and not e.text.strip():
                    e.text = None
            out.append(ET.tostring(c))
        return out

    new = [ET.fromstring(ET.tostring(c)) for c in vn_cells]
    for c in new:
        c.tail = None
        if c.get("parent") == vn_layer:
            c.set("parent", en_layer)
    if sig(new) == sig(en_cells):
        return False
    for c in en_cells:
        en_root.remove(c)
    pos = list(en_root).index(en_root.find(f"mxCell[@id='{en_layer}']")) + 1
    for k, c in enumerate(new):
        en_root.insert(pos + k, c)
    tree.write(drawio_path, encoding="utf-8", xml_declaration=False)
    print(f"[drawio] Đồng bộ {len(new)} ký hiệu 'Vị trí đo' từ tab _vn sang tab _en.")
    return True


def export_layouts(drawio_path: str, out_dir: str | None = None, scale: float = 1.0) -> dict:
    """Xuất mọi tab bằng DrawIOExporter (tên ảnh = tên tab) -> {lang: png_path}."""
    _, DrawIOExporter = _tool()
    if not DrawIOExporter.find_drawio_executable():
        raise SystemExit("[FATAL] Không tìm thấy draw.io desktop (draw.io.exe).")
    out_dir = out_dir or os.path.dirname(os.path.abspath(drawio_path))
    targets = {lang: os.path.join(out_dir, f"{BASENAME}_{lang}.png") for lang in ("vn", "en")}
    before = {p: os.path.getmtime(p) if os.path.exists(p) else 0 for p in targets.values()}
    DrawIOExporter.export_all_tabs(drawio_path, output_dir=out_dir, format="png", scale=scale)
    for lang, p in targets.items():
        if not os.path.isfile(p) or os.path.getmtime(p) <= before[p]:
            raise SystemExit(f"[FATAL] draw.io không xuất được {p} (tab {BASENAME}_{lang}).")
        _flatten_white(p)
        print(f"[drawio] {os.path.basename(drawio_path)} -> {os.path.basename(p)}")
    return targets


def _flatten_white(path: str):
    from PIL import Image
    im = Image.open(path)
    if im.mode in ("RGBA", "LA", "P"):
        im = im.convert("RGBA")
        bg = Image.new("RGB", im.size, "white")
        bg.paste(im, mask=im.split()[-1])
        bg.save(path)


# ---------------------------------------------------------------------------
# Kiểm tra sơ đồ sau khi người dùng sửa tay (chỉ cảnh báo, không dừng build)
# ---------------------------------------------------------------------------
Z_ROOTS = [r"Z:\9.NV2 - MOI TRUONG\1.Ho So Moi Truong\2026",
           r"Z:\9.NV2 - MOI TRUONG\1.Ho So Moi Truong\2026\0.VIMCERTS078",
           r"Z:\9.NV2 - MOI TRUONG\1.Ho So Moi Truong\2026\0.1.Tân Hiệp"]


def _plain(value: str) -> str:
    """Chữ của ô (bỏ thẻ html nếu người dùng bật định dạng html khi sửa)."""
    import html
    import re
    return html.unescape(re.sub(r"<[^>]+>", "", value or "")).strip()


def check_layout(drawio_path: str, expected_labels: list[str] | None = None) -> list[str]:
    """[Layout Labels]: nhãn lớp "Vị trí đo" tab _vn khác danh sách code mong đợi (xóa/nhân đôi nhầm điểm);
    [Layout Out Of Plan]: điểm có tâm nằm ngoài ảnh mặt bằng (kéo lạc ra ngoài / vào vùng chú thích)."""
    from collections import Counter
    warns = []
    tabs = {d.get("name", "").rsplit("_", 1)[-1].lower(): d for d in ET.parse(drawio_path).getroot().findall("diagram")}
    if "vn" not in tabs:
        return [f"[Layout Labels] {os.path.basename(drawio_path)}: không thấy tab {BASENAME}_vn"]
    root = _model(tabs["vn"]).find("root")
    layer = _layer_id(root, LAYER_POINTS, LAYER_LABELS[LAYER_POINTS])
    if not layer:
        return [f"[Layout Labels] {os.path.basename(drawio_path)}: tab _vn không có lớp 'Vị trí đo'"]
    plan = root.find("mxCell[@id='matbang']")
    if plan is None:
        plan = next((c for c in root.findall("mxCell") if "shape=image" in (c.get("style") or "")), None)
    pts = [c for c in root.findall("mxCell") if c.get("parent") == layer and c.get("vertex") == "1"]
    labels = [_plain(c.get("value")) for c in pts]
    if expected_labels is not None:
        got, want = Counter(labels), Counter(expected_labels)
        missing, extra = list((want - got).elements()), list((got - want).elements())
        if missing or extra:
            warns.append(f"[Layout Labels] tab _vn: thiếu {missing or '-'}, thừa {extra or '-'} "
                         f"(code mong đợi {len(expected_labels)} ký hiệu, file có {len(labels)})")
    if plan is not None:
        g = plan.find("mxGeometry")
        px, py = float(g.get("x", 0)), float(g.get("y", 0))
        pw, ph = float(g.get("width", 0)), float(g.get("height", 0))
        for c, lab in zip(pts, labels):
            g = c.find("mxGeometry")
            if g is None:
                continue
            cx = float(g.get("x", 0)) + float(g.get("width", 0)) / 2
            cy = float(g.get("y", 0)) + float(g.get("height", 0)) / 2
            if not (px <= cx <= px + pw and py <= cy <= py + ph):
                warns.append(f"[Layout Out Of Plan] tab _vn: ký hiệu '{lab or c.get('id')}' tâm ({cx:.0f}, {cy:.0f}) "
                             f"nằm ngoài mặt bằng ({pw:.0f} × {ph:.0f})")
    return warns


def check_z_newer(drawio_path: str) -> list[str]:
    """[Layout Z Newer]: bản .drawio trên Z mới hơn bản Input (người dùng sửa trên Z) — chạy tiếp sẽ dùng bản cũ,
    và robocopy (mặc định chép cả file cũ hơn) sẽ đè mất chỉnh sửa trên Z (AGENTS mục 0)."""
    parts = os.path.abspath(drawio_path).split(os.sep)
    try:
        code6 = parts[parts.index("TBKQ") - 1]
    except ValueError:
        return []
    if not (code6.isdigit() and len(code6) == 6):
        return []
    import glob
    for z_root in Z_ROOTS:
        if not os.path.isdir(z_root):
            continue
        hits = [d for d in glob.glob(os.path.join(z_root, f"*{code6}*")) if os.path.isdir(d)]
        if len(hits) != 1:
            continue
        z_file = os.path.join(hits[0], "TBKQ", "_diagrams", os.path.basename(drawio_path))
        if os.path.isfile(z_file) and os.path.getmtime(z_file) > os.path.getmtime(drawio_path) + 2:
            return [f"[Layout Z Newer] {z_file} mới hơn bản Input — chép về Input\\{code6}\\TBKQ\\_diagrams\\ "
                    f"rồi chạy lại; KHÔNG robocopy lên Z trước khi đồng bộ."]
        return []
    return []


def _report(warns: list[str]):
    for w in warns:
        lvl = "[CRITICAL WARNING]" if w.startswith("[Layout Z Newer]") else "[WARNING]"
        print(f"{lvl} {w}")


def ensure_layout(diagram_dir: str, make_plan, markers, rooms=None, legend=None, legend_cols=None,
                  overwrite: bool = False, expected_labels: list[str] | None = None, **kw) -> dict:
    """Tạo .drawio nếu chưa có (make_plan() chỉ gọi khi phải dựng), đồng bộ EN theo VN, kiểm tra, xuất PNG.
    expected_labels: nhãn điểm đo mong đợi (mặc định = chữ của `markers`)."""
    path = os.path.join(diagram_dir, f"{BASENAME}.drawio")
    if os.path.exists(path):
        _report(check_z_newer(path))
    if overwrite or not os.path.exists(path):
        build_layout_drawio(path, make_plan(), markers, rooms, legend, legend_cols, **kw)
    else:
        print(f"[drawio] Dùng sơ đồ đã có (giữ chỉnh tay của người dùng): {path}")
        sync_points(path)
    if expected_labels is None:
        expected_labels = [m.get("text", "") for m in markers]
    _report(check_layout(path, expected_labels))
    return export_layouts(path, diagram_dir)
