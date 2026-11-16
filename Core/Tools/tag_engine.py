# -*- coding: utf-8 -*-
"""
Bộ máy điền TAG cho template Word (.docx) và Excel (.xlsx) — QUATEST 3 / Phòng NV2.

Cú pháp TAG (tương thích docxtpl / Jinja ở mức cơ bản, KHÔNG cần cài thêm thư viện):

    {{ ten_bien }}                  -> thay bằng giá trị (chuỗi, số). "\n" -> xuống dòng (w:br), "\t" -> tab.
    {{ vt.ten }} / {{ ds[0] }}      -> truy cập thuộc tính / phần tử.
    {{ 'a' if dk else 'b' }}        -> biểu thức Python đơn giản (không có hàm hệ thống, chỉ len/str/int/...).

    {%p if dieu_kien %}   ...  {%p elif dk2 %} ... {%p else %} ... {%p endif %}
    {%p for x in danh_sach %} ...  {%p endfor %}
        -> TAG khối mức ĐOẠN VĂN: đoạn chứa TAG bị xóa; các đoạn/bảng nằm giữa được giữ, xóa hoặc lặp.
           Dùng được ở thân văn bản, trong ô bảng, header/footer, textbox.
    {%tr if dk %} ... {%tr endif %} / {%tr for x in ds %} ... {%tr endfor %}
        -> TAG khối mức DÒNG BẢNG: dòng chứa TAG bị xóa; các dòng giữa được giữ, xóa hoặc lặp.

    Trong vòng lặp có biến `loop`: loop.index (1, 2, ...), loop.index0, loop.first, loop.last, loop.length.

Giá trị đặc biệt:
    TagImage(path, width_cm)  -> chèn ảnh (chỉ đặt chiều rộng, giữ tỉ lệ gốc, không crop — AGENTS.md mục 3).
    TagBlock(func)            -> đoạn văn chỉ chứa {{ ten }} sẽ được thay bằng các phần tử do func dựng
                                  (vd lưới ảnh 2/3 cột). func(doc, anchor_p, ctx) -> list phần tử XML (hoặc []).

Biến không có trong dữ liệu -> chuỗi rỗng và được ghi vào `TagRenderer.missing` để cảnh báo.
"""

import copy
import re
import zipfile
import shutil
import os
from dataclasses import is_dataclass, fields as dc_fields
from typing import Any, Callable, Dict, List, Optional

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

TAG_RE = re.compile(r"\{\{(.*?)\}\}", re.S)
BLOCK_RE = re.compile(r"^\s*\{%\s*(p|tr)\s+(.*?)\s*%\}\s*$", re.S)
ANY_BLOCK_RE = re.compile(r"\{%.*?%\}", re.S)
FOR_RE = re.compile(r"^for\s+(.+?)\s+in\s+(.+)$", re.S)

# Thẻ "thuộc tính" luôn đứng yên trong container, không tham gia khối TAG
_PROPERTY_TAGS = {
    qn("w:sectPr"), qn("w:tblPr"), qn("w:tblGrid"), qn("w:tcPr"), qn("w:trPr"),
    qn("w:sdtPr"), qn("w:sdtEndPr"), qn("w:tblPrEx"),
}
_P, _TBL, _TR, _TC, _SDT = qn("w:p"), qn("w:tbl"), qn("w:tr"), qn("w:tc"), qn("w:sdt")
_T, _R = qn("w:t"), qn("w:r")

_IMG_MARK = "\ue000IMG{}\ue001"  # ký tự vùng riêng (Private Use) đánh dấu chỗ chèn ảnh
_IMG_MARK_RE = re.compile("\ue000IMG(\\d+)\ue001")


class TemplateTagError(Exception):
    """Lỗi cú pháp TAG trong template (khối if/for không đóng, ...)."""


# ---------------------------------------------------------------------------
# Giá trị đặc biệt
# ---------------------------------------------------------------------------
class TagImage:
    """Ảnh chèn vào vị trí TAG. Chỉ đặt chiều rộng (không đặt chiều cao, không crop)."""

    def __init__(self, path: str, width_cm: float = 7.5):
        self.path = path
        self.width_cm = width_cm

    def __bool__(self):
        return bool(self.path) and os.path.isfile(self.path)

    def __str__(self):
        return ""


class TagBlock:
    """Khối dựng bằng code thay cho một đoạn văn chỉ chứa {{ ten }} (vd lưới ảnh)."""

    def __init__(self, func: Callable, *args, **kwargs):
        self.func = func
        self.args = args
        self.kwargs = kwargs

    def build(self, doc, anchor_p, ctx) -> List:
        return self.func(doc, anchor_p, ctx, *self.args, **self.kwargs) or []

    def __bool__(self):
        return True

    def __str__(self):
        return ""


class Undefined:
    """Giá trị của biến không có trong dữ liệu: rỗng, False, lặp rỗng."""

    def __init__(self, name: str = ""):
        self._name = name

    def __str__(self):
        return ""

    def __repr__(self):
        return f"Undefined({self._name!r})"

    def __bool__(self):
        return False

    def __iter__(self):
        return iter(())

    def __len__(self):
        return 0

    def __getattr__(self, item):
        if item.startswith("__"):
            raise AttributeError(item)
        return Undefined(f"{self._name}.{item}")

    def __getitem__(self, item):
        return Undefined(f"{self._name}[{item!r}]")

    def __call__(self, *a, **k):
        return self

    def __eq__(self, other):
        return isinstance(other, Undefined) or other in ("", None)

    def __hash__(self):
        return 0

    def __add__(self, other):
        return other

    def __radd__(self, other):
        return other


_ACTIVE_MISSING: Optional[set] = None  # tập ghi nhận TAG thiếu của lần điền đang chạy


class AttrDict(dict):
    """dict cho phép truy cập kiểu thuộc tính (vt.ten); khóa thiếu -> Undefined (và được ghi nhận)."""

    def __getattr__(self, item):
        if item.startswith("__"):
            raise AttributeError(item)
        if item in self:
            return self[item]
        if _ACTIVE_MISSING is not None:
            _ACTIVE_MISSING.add(f"*.{item}")
        return Undefined(item)

    def __missing__(self, key):
        if _ACTIVE_MISSING is not None:
            _ACTIVE_MISSING.add(f"*[{key!r}]")
        return Undefined(str(key))


class Loop:
    def __init__(self, index0: int, length: int):
        self.index0 = index0
        self.index = index0 + 1
        self.length = length
        self.first = index0 == 0
        self.last = index0 == length - 1
        self.revindex = length - index0


def to_context(value: Any) -> Any:
    """Chuyển dữ liệu (dict, list, dataclass) sang dạng dùng được trong TAG."""
    if isinstance(value, (TagImage, TagBlock, Undefined, Loop, str, int, float, bool)) or value is None:
        return value
    if isinstance(value, dict):
        return AttrDict({str(k): to_context(v) for k, v in value.items()})
    if isinstance(value, (list, tuple)):
        return [to_context(v) for v in value]
    if is_dataclass(value) and not isinstance(value, type):
        out = AttrDict()
        for f in dc_fields(value):
            out[f.name] = to_context(getattr(value, f.name))
        for k, v in vars(value).items():  # thuộc tính gán thêm sau __init__ (vd cert_no_vn)
            if k not in out:
                out[k] = to_context(v)
        return out
    return value


_SAFE_BUILTINS = {
    "len": len, "str": str, "int": int, "float": float, "round": round, "abs": abs,
    "min": min, "max": max, "sum": sum, "any": any, "all": all, "sorted": sorted,
    "enumerate": enumerate, "range": range, "zip": zip, "list": list, "dict": dict,
    "True": True, "False": False, "None": None,
}


class _Scope(dict):
    """Không gian tên cho eval: biến thiếu -> Undefined (và được ghi nhận)."""

    def __init__(self, data: Dict[str, Any], parent: Optional["_Scope"] = None, missing: Optional[set] = None):
        super().__init__(data)
        self.parent = parent
        self.missing = missing if missing is not None else (parent.missing if parent else set())

    def __missing__(self, key):
        if self.parent is not None:
            return self.parent[key]
        if key in _SAFE_BUILTINS:
            return _SAFE_BUILTINS[key]
        self.missing.add(key)
        return Undefined(key)

    def child(self, data: Dict[str, Any]) -> "_Scope":
        return _Scope(data, parent=self)


_CODE_CACHE: Dict[str, Any] = {}


def _eval(expr: str, scope: _Scope) -> Any:
    expr = expr.strip()
    if not expr:
        return ""
    code = _CODE_CACHE.get(expr)
    if code is None:
        try:
            code = compile(expr, "<tag>", "eval")
        except SyntaxError as e:
            raise TemplateTagError(f"TAG sai cú pháp: {{{{ {expr} }}}} ({e.msg})") from None
        _CODE_CACHE[expr] = code
    try:
        return eval(code, {"__builtins__": {}}, scope)
    except TemplateTagError:
        raise
    except Exception as e:  # lỗi lúc tính (vd chia 0, None.attr) -> rỗng, ghi nhận
        scope.missing.add(f"{expr} ({type(e).__name__})")
        return Undefined(expr)


def _single_tag(text: str):
    """Trả về match nếu chuỗi (bỏ khoảng trắng 2 đầu) chỉ gồm đúng 1 TAG {{ }}."""
    t = (text or "").strip()
    if not t.startswith("{{"):
        return None
    ms = list(TAG_RE.finditer(t))
    if len(ms) == 1 and ms[0].start() == 0 and ms[0].end() == len(t):
        return ms[0]
    return None


def _to_text(value: Any) -> str:
    if value is None or isinstance(value, Undefined):
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


# ---------------------------------------------------------------------------
# Tiện ích XML
# ---------------------------------------------------------------------------
def _owner_p(node):
    el = node.getparent()
    while el is not None and el.tag != _P:
        el = el.getparent()
    return el


def _p_text_nodes(p) -> List:
    """Các w:t thuộc trực tiếp đoạn p (không lấy chữ trong textbox lồng bên trong)."""
    return [t for t in p.iter(_T) if _owner_p(t) is p]


def _p_text(p) -> str:
    return "".join(t.text or "" for t in _p_text_nodes(p))


def _row_text(tr) -> str:
    return "".join(t.text or "" for t in tr.iter(_T))


def _owner_run(t):
    el = t.getparent()
    while el is not None and el.tag != _R:
        el = el.getparent()
    return el


# ---------------------------------------------------------------------------
# Bộ máy điền DOCX
# ---------------------------------------------------------------------------
class TagRenderer:
    """Điền TAG vào một python-docx Document."""

    def __init__(self, doc, context: Dict[str, Any], strict: bool = False):
        self.doc = doc
        self.strict = strict
        self.missing: set = set()
        self.scope = _Scope({k: to_context(v) for k, v in (context or {}).items()}, missing=self.missing)
        self._images: List[TagImage] = []
        self._part = None

    # ---- công khai ----
    def render(self):
        global _ACTIVE_MISSING
        _ACTIVE_MISSING = self.missing
        try:
            return self._render_all()
        finally:
            _ACTIVE_MISSING = None

    def _render_all(self):
        doc = self.doc
        self._part = doc.part
        self._render_container(doc.element.body, self.scope, "p")
        self._explode_specials(doc.element.body)
        seen = set()
        for rel in list(doc.part.rels.values()):
            part = getattr(rel, "target_part", None) if not rel.is_external else None
            if part is None or id(part) in seen:
                continue
            ct = getattr(part, "content_type", "")
            if ct.endswith("header+xml") or ct.endswith("footer+xml"):
                seen.add(id(part))
                self._part = part
                self._render_container(part.element, self.scope, "p")
                self._explode_specials(part.element)
        self._part = doc.part
        if self.strict and self.missing:
            raise TemplateTagError("Thiếu dữ liệu cho TAG: " + ", ".join(sorted(self.missing)))
        return self.missing

    # ---- khối ----
    def _marker(self, el, kind: str):
        if el.tag == _P and kind == "p":
            txt = _p_text(el)
        elif el.tag == _TR and kind == "tr":
            txt = _row_text(el)
        else:
            return None
        if "{%" not in txt:
            return None
        if kind == "tr" and "{%tr" not in txt.replace("{% tr", "{%tr"):
            return None  # chỉ có TAG khối {%p %} trong ô -> xử lý ở cấp ô
        m = BLOCK_RE.match(txt)
        if not m:
            raise TemplateTagError(f"Đoạn chứa TAG khối phải chỉ có duy nhất TAG đó: {txt.strip()[:120]!r}")
        if m.group(1) != kind:
            raise TemplateTagError(
                f"TAG {{%{m.group(1)} ...%}} đặt sai chỗ (cần {{%{kind} ...%}}): {txt.strip()[:120]!r}")
        stmt = m.group(2).strip()
        kw = stmt.split(None, 1)[0]
        arg = stmt[len(kw):].strip()
        return kw, arg, stmt

    def _parse(self, elements: List, kind: str) -> List:
        root: List = []
        stack: List = [(root, None)]  # (danh sách đích, nút if/for)
        for el in elements:
            mk = self._marker(el, kind)
            if mk is None:
                stack[-1][0].append(("el", el))
                continue
            kw, arg, stmt = mk
            if kw == "for":
                node = {"type": "for", "stmt": stmt, "body": []}
                stack[-1][0].append(("node", node))
                stack.append((node["body"], node))
            elif kw == "if":
                node = {"type": "if", "branches": [[arg, []]]}
                stack[-1][0].append(("node", node))
                stack.append((node["branches"][0][1], node))
            elif kw in ("elif", "else"):
                if len(stack) < 2 or stack[-1][1]["type"] != "if":
                    raise TemplateTagError(f"{{%{kind} {stmt} %}} không nằm trong khối if")
                node = stack[-1][1]
                branch = [arg if kw == "elif" else "True", []]
                node["branches"].append(branch)
                stack[-1] = (branch[1], node)
            elif kw in ("endfor", "endif"):
                want = "for" if kw == "endfor" else "if"
                if len(stack) < 2 or stack[-1][1]["type"] != want:
                    raise TemplateTagError(f"{{%{kind} {kw} %}} không khớp khối mở")
                stack.pop()
            else:
                raise TemplateTagError(f"Không hiểu TAG khối: {{%{kind} {stmt} %}}")
        if len(stack) != 1:
            raise TemplateTagError(f"Khối {{%{kind} {stack[-1][1]['type']} ... %}} chưa được đóng")
        return root

    def _render_nodes(self, nodes: List, scope: _Scope, copy_mode: bool) -> List:
        out = []
        for typ, item in nodes:
            if typ == "el":
                el = copy.deepcopy(item) if copy_mode else item
                out.extend(self._render_element(el, scope))
            elif item["type"] == "if":
                for cond, body in item["branches"]:
                    if _eval(cond, scope):
                        out.extend(self._render_nodes(body, scope, copy_mode))
                        break
            else:  # for
                m = FOR_RE.match(item["stmt"])
                if not m:
                    raise TemplateTagError(f"Vòng lặp sai cú pháp: {{% {item['stmt']} %}}")
                names = [n.strip() for n in m.group(1).split(",")]
                seq = _eval(m.group(2), scope)
                seq = list(seq) if seq else []
                for i, val in enumerate(seq):
                    if len(names) == 1:
                        data = {names[0]: to_context(val)}
                    else:
                        vals = list(val)
                        data = {n: to_context(vals[j] if j < len(vals) else None) for j, n in enumerate(names)}
                    data["loop"] = Loop(i, len(seq))
                    out.extend(self._render_nodes(item["body"], scope.child(data), True))
        return out

    def _render_container(self, parent, scope: _Scope, kind: str):
        children = list(parent)
        content = [c for c in children if c.tag not in _PROPERTY_TAGS]
        if not content:
            return
        # Không có TAG nào trong container -> chỉ điền biến, không dựng lại (nhanh, giữ nguyên)
        nodes = self._parse(content, kind)
        rendered = self._render_nodes(nodes, scope, copy_mode=False)
        first_idx = children.index(content[0])
        for c in content:
            parent.remove(c)
        # chèn lại đúng vị trí (sau các thẻ thuộc tính đứng đầu như tcPr/tblPr/tblGrid)
        for offset, el in enumerate(rendered):
            parent.insert(first_idx + offset, el)
        if parent.tag == _TC and parent.find(_P) is None:
            parent.append(OxmlElement("w:p"))

    def _render_element(self, el, scope: _Scope) -> List:
        tag = el.tag
        if tag == _P:
            # textbox lồng trong đoạn
            for txbx in [x for x in el.iter(qn("w:txbxContent")) if _owner_p(x) is el]:
                self._render_container(txbx, scope, "p")
            # đoạn chỉ chứa 1 TAG TagBlock -> thay bằng khối dựng sẵn
            m = _single_tag(_p_text(el))
            if m:
                val = _eval(m.group(1), scope)
                if isinstance(val, TagBlock):
                    return list(val.build(self.doc, el, scope))
            self._substitute(el, scope)
            return [el]
        if tag == _TBL:
            self._render_container(el, scope, "tr")
            return [el]
        if tag == _TR:
            for tc in el.findall(_TC):
                self._render_container(tc, scope, "p")
            # ô lồng trong w:sdt của dòng
            for sdt in el.findall(_SDT):
                content = sdt.find(qn("w:sdtContent"))
                if content is not None:
                    for tc in content.findall(_TC):
                        self._render_container(tc, scope, "p")
            return [el]
        if tag == _SDT:
            content = el.find(qn("w:sdtContent"))
            if content is not None:
                self._render_container(content, scope, "p")
            return [el]
        # các thẻ khác (bookmark, ...) giữ nguyên, nhưng vẫn điền biến nếu có chữ
        for p in el.iter(_P):
            self._substitute(p, scope)
        return [el]

    # ---- thay biến trong đoạn ----
    def _substitute(self, p, scope: _Scope):
        nodes = _p_text_nodes(p)
        if not nodes:
            return
        texts = [n.text or "" for n in nodes]
        full = "".join(texts)
        if "{{" not in full:
            return
        if ANY_BLOCK_RE.search(full):
            raise TemplateTagError(f"TAG khối {{% %}} phải nằm riêng một đoạn/dòng: {full.strip()[:120]!r}")
        # vị trí bắt đầu của từng node trong chuỗi ghép
        starts = []
        pos = 0
        for t in texts:
            starts.append(pos)
            pos += len(t)

        def locate(ch):
            for i in range(len(nodes)):
                if starts[i] <= ch < starts[i] + len(texts[i]):
                    return i
            return len(nodes) - 1

        matches = list(TAG_RE.finditer(full))
        for m in reversed(matches):
            val = _eval(m.group(1), scope)
            if isinstance(val, TagImage):
                self._images.append(val)
                rep = _IMG_MARK.format(len(self._images) - 1) if val else ""
            elif isinstance(val, TagBlock):
                rep = ""
                scope.missing.add(f"{m.group(1).strip()} (TagBlock phải nằm riêng một đoạn)")
            else:
                rep = _to_text(val)
            s, e = m.start(), m.end() - 1
            i_s, i_e = locate(s), locate(e)
            off_s, off_e = s - starts[i_s], e - starts[i_e]
            if i_s == i_e:
                cur = nodes[i_s].text or ""
                nodes[i_s].text = cur[:off_s] + rep + cur[off_e + 1:]
            else:
                cur_s = nodes[i_s].text or ""
                nodes[i_s].text = cur_s[:off_s] + rep
                for k in range(i_s + 1, i_e):
                    nodes[k].text = ""
                cur_e = nodes[i_e].text or ""
                nodes[i_e].text = cur_e[off_e + 1:]
            for k in range(i_s, i_e + 1):
                nodes[k].set(XML_SPACE, "preserve")

    # ---- \n, \t, ảnh ----
    def _explode_specials(self, root):
        for t in list(root.iter(_T)):
            txt = t.text or ""
            if "\n" not in txt and "\t" not in txt and "\ue000IMG" not in txt:
                continue
            run = _owner_run(t)
            if run is None or t.getparent() is not run:
                t.text = txt.replace("\n", " ").replace("\t", " ")
                continue
            rpr = run.find(qn("w:rPr"))
            # tách thành các đoạn: chữ / br / tab / ảnh
            pieces = re.split("(\n|\t|\ue000IMG\\d+\ue001)", txt)
            new_children = []
            img_runs_after = []
            for piece in pieces:
                if piece == "":
                    continue
                if piece == "\n":
                    new_children.append(OxmlElement("w:br"))
                elif piece == "\t":
                    new_children.append(OxmlElement("w:tab"))
                elif _IMG_MARK_RE.fullmatch(piece):
                    new_children.append(("IMG", int(_IMG_MARK_RE.fullmatch(piece).group(1))))
                else:
                    nt = OxmlElement("w:t")
                    nt.text = piece
                    nt.set(XML_SPACE, "preserve")
                    new_children.append(nt)
            if not any(isinstance(c, tuple) for c in new_children):
                idx = list(run).index(t)
                run.remove(t)
                for off, c in enumerate(new_children):
                    run.insert(idx + off, c)
                continue
            # Có ảnh: tách run thành nhiều run (chữ trước, ảnh, chữ sau)
            parent = run.getparent()
            r_idx = list(parent).index(run)
            idx = list(run).index(t)
            before = list(run)[:idx]
            after = list(run)[idx + 1:]
            groups: List = []
            cur: List = []
            for c in new_children:
                if isinstance(c, tuple):
                    if cur:
                        groups.append(("txt", cur))
                        cur = []
                    groups.append(("img", c[1]))
                else:
                    cur.append(c)
            if cur:
                groups.append(("txt", cur))
            new_runs = []
            for gi, (kind, payload) in enumerate(groups):
                nr = OxmlElement("w:r")
                if rpr is not None:
                    nr.append(copy.deepcopy(rpr))
                if gi == 0:
                    for b in before:
                        if b.tag != qn("w:rPr"):
                            nr.append(copy.deepcopy(b))
                if kind == "txt":
                    for c in payload:
                        nr.append(c)
                else:
                    img = self._images[payload]
                    try:
                        inline = self._part.new_pic_inline(img.path, Cm(img.width_cm), None)
                        drawing = OxmlElement("w:drawing")
                        drawing.append(inline)
                        nr.append(drawing)
                        for sr in drawing.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}srcRect"):
                            sr.attrib.clear()
                    except Exception as _err:
                        print(f"[Warning] tag_engine: không chèn được ảnh {img.path}: {type(_err).__name__}: {_err}")
                if gi == len(groups) - 1:
                    for a in after:
                        nr.append(copy.deepcopy(a))
                new_runs.append(nr)
            parent.remove(run)
            for off, nr in enumerate(new_runs):
                parent.insert(r_idx + off, nr)


def render_docx(doc, context: Dict[str, Any], strict: bool = False, warn: bool = True) -> set:
    """Điền TAG vào Document (sửa tại chỗ). Trả về tập TAG thiếu dữ liệu."""
    r = TagRenderer(doc, context, strict=strict)
    missing = r.render()
    if warn and missing:
        print("[TAG] Thiếu dữ liệu (để trống): " + ", ".join(sorted(missing)))
    return missing


def render_docx_file(template_path: str, context: Dict[str, Any], out_path: str, strict: bool = False):
    import docx
    doc = docx.Document(template_path)
    render_docx(doc, context, strict=strict)
    doc.save(out_path)
    return out_path


def has_tags(doc) -> bool:
    """Template có chứa TAG {{ }} / {% %} không (body + header/footer)."""
    xmls = [doc.element.body]
    for rel in doc.part.rels.values():
        if not rel.is_external:
            ct = getattr(rel.target_part, "content_type", "")
            if ct.endswith("header+xml") or ct.endswith("footer+xml"):
                xmls.append(rel.target_part.element)
    for x in xmls:
        txt = "".join(t.text or "" for t in x.iter(_T))
        if "{{" in txt or "{%" in txt:
            return True
    return False


def find_leftover_tags(doc) -> List[str]:
    """Liệt kê TAG còn sót trong tài liệu đã điền (dùng cho QA)."""
    found = []
    xmls = [doc.element.body]
    for rel in doc.part.rels.values():
        if not rel.is_external:
            ct = getattr(rel.target_part, "content_type", "")
            if ct.endswith("header+xml") or ct.endswith("footer+xml"):
                xmls.append(rel.target_part.element)
    for x in xmls:
        for p in x.iter(_P):
            txt = _p_text(p)
            for m in re.finditer(r"\{\{.*?\}\}|\{%.*?%\}", txt):
                found.append(m.group(0))
    return found


def list_tags(doc) -> List[str]:
    """Liệt kê mọi TAG trong template (để kiểm tra/ghi tài liệu)."""
    return find_leftover_tags(doc)


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------
def render_xlsx(wb, context: Dict[str, Any], warn: bool = True) -> set:
    """
    Điền TAG {{ }} trong các ô chữ (không đụng ô công thức) và header/footer trang in của openpyxl Workbook.
    Ô chỉ chứa đúng 1 TAG mà giá trị là số/ngày -> ghi đúng kiểu số/ngày (để công thức Excel tính được).
    """
    import datetime as _dt
    global _ACTIVE_MISSING
    missing: set = set()
    scope = _Scope({k: to_context(v) for k, v in (context or {}).items()}, missing=missing)
    _ACTIVE_MISSING = missing

    def sub(text: str) -> str:
        return TAG_RE.sub(lambda m: _to_text(_eval(m.group(1), scope)), text)

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if not isinstance(v, str) or "{{" not in v or v.startswith("="):
                    continue
                m = _single_tag(v)
                if m:
                    val = _eval(m.group(1), scope)
                    if isinstance(val, (int, float, _dt.date, _dt.datetime, _dt.time)) and not isinstance(val, bool):
                        c.value = val
                    else:
                        txt = _to_text(val)
                        c.value = txt if txt != "" else None
                else:
                    c.value = sub(v)
        for hf_name in ("oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter"):
            hf = getattr(ws, hf_name, None)
            if hf is None:
                continue
            for part_name in ("left", "center", "right"):
                part = getattr(hf, part_name, None)
                if part is not None and part.text and "{{" in part.text:
                    part.text = sub(part.text)
    _ACTIVE_MISSING = None
    if warn and missing:
        print("[TAG] Excel thiếu dữ liệu (để trống): " + ", ".join(sorted(missing)))
    return missing


def xlsx_set_cells_raw(src_path: str, dst_path: str, sheet_cells: Dict[str, Dict[str, str]],
                       header_subs: Optional[List] = None):
    """
    Ghi chuỗi (TAG) vào ô Excel bằng cách sửa XML trực tiếp, KHÔNG qua openpyxl
    (để không làm mất ảnh, hình vẽ, textbox... của template). Chỉ dùng khi dựng template.
    sheet_cells = {"Tên sheet": {"C9": "{{ kh_ten }}", ...}}
    header_subs = [(regex, thay_bang), ...] áp cho chữ header/footer trang in (oddHeader...) của mọi sheet.
    """
    from lxml import etree
    NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
          "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
    M = "{%s}" % NS["m"]
    tmp = dst_path + ".tmp"
    with zipfile.ZipFile(src_path) as zin:
        wb_xml = etree.fromstring(zin.read("xl/workbook.xml"))
        rels = etree.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
        rel_map = {r.get("Id"): r.get("Target") for r in rels}
        sheet_files = {}
        for s in wb_xml.find("m:sheets", NS):
            rid = s.get("{%s}id" % NS["r"])
            target = rel_map[rid]
            target = target.lstrip("/")
            if not target.startswith("xl/"):
                target = "xl/" + target
            sheet_files[s.get("name")] = target
        sst_name = "xl/sharedStrings.xml"
        sst = etree.fromstring(zin.read(sst_name)) if sst_name in zin.namelist() else None
        modified = {}
        for sheet, cells in sheet_cells.items():
            if sheet not in sheet_files:
                raise KeyError(f"Không có sheet {sheet!r}")
            fname = sheet_files[sheet]
            root = etree.fromstring(modified.get(fname) or zin.read(fname))
            sd = root.find("m:sheetData", NS)
            for ref, text in cells.items():
                col = re.match(r"[A-Z]+", ref).group(0)
                rown = int(ref[len(col):])
                row_el = None
                for r_el in sd.findall("m:row", NS):
                    if int(r_el.get("r")) == rown:
                        row_el = r_el
                        break
                if row_el is None:
                    raise KeyError(f"{sheet}!{ref}: không có dòng {rown}")
                c_el = None
                for c in row_el.findall("m:c", NS):
                    if c.get("r") == ref:
                        c_el = c
                        break
                if c_el is None:
                    raise KeyError(f"{sheet}!{ref}: ô chưa tồn tại trong template")
                for child in list(c_el):
                    if child.tag in (M + "v", M + "f", M + "is"):
                        c_el.remove(child)
                if text is None:
                    c_el.attrib.pop("t", None)
                    continue
                c_el.set("t", "inlineStr")
                is_el = etree.SubElement(c_el, M + "is")
                t_el = etree.SubElement(is_el, M + "t")
                t_el.text = text
                t_el.set(XML_SPACE, "preserve")
            modified[fname] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        if header_subs:
            for fname in sheet_files.values():
                root = etree.fromstring(modified.get(fname) or zin.read(fname))
                changed = False
                for hf in root.iter(M + "oddHeader", M + "evenHeader", M + "firstHeader",
                                    M + "oddFooter", M + "evenFooter", M + "firstFooter"):
                    txt = hf.text or ""
                    new = txt
                    for pat, rep in header_subs:
                        new = re.sub(pat, rep, new)
                    if new != txt:
                        hf.text = new
                        changed = True
                if changed:
                    modified[fname] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = modified.get(item.filename)
                if data is None:
                    data = zin.read(item.filename)
                zout.writestr(item, data)
    shutil.move(tmp, dst_path)
    return dst_path
