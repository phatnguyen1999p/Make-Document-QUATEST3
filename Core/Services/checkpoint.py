# -*- coding: utf-8 -*-
"""
Checkpoint (phân đoạn tiến độ) cho pipeline hồ sơ QUATEST 3.

Mục tiêu: mỗi giai đoạn tốn thời gian/token (đặc biệt là OCR file scan) chỉ chạy
MỘT lần; kết quả được ghi ra file JSON (bản chính) + Excel (bản xem/sửa tay).
Chạy lại thì đọc checkpoint thay vì OCR lại.

Vị trí lưu:  Input\\<hồ sơ>\\ThucHien\\_checkpoint\\
    manifest.json     Tiến độ từng giai đoạn (done / failed / running), thời gian, file đầu ra
    ocr.json          Văn bản OCR theo từng file scan, từng trang (kèm mã băm file gốc)
    ocr.xlsx          Bản Excel của ocr.json. Sửa tay chỗ OCR sai rồi lưu -> lần chạy sau tự nạp lại
    header_ocr.json   OCR phần đầu trang dùng để tách BBLM / Ghi chép (scan_splitter)
    metadata.json     Metadata đã trích từ văn bản OCR (intake_helper)

Nguyên tắc: lỗi đọc/ghi checkpoint KHÔNG bao giờ làm hỏng pipeline chính; chỉ in cảnh báo
và rơi về hành vi cũ (chạy lại từ đầu).
"""

import hashlib
import json
import os
import shutil
import time
from contextlib import contextmanager
from datetime import datetime
from typing import Any, Dict, Iterable, List, Optional

SCHEMA_VERSION = 1
CHECKPOINT_DIRNAME = "_checkpoint"
OCR_ENGINE = "windows-media-ocr"

# Thứ tự giai đoạn. Làm mới một giai đoạn sẽ làm mất hiệu lực các giai đoạn đứng sau nó.
STAGES = ("ocr", "intake", "split_scan", "bckq", "tbkq", "pdf", "qa")

STAGE_LABELS = {
    "ocr": "OCR file scan",
    "intake": "Trích metadata / tạo profile.json",
    "split_scan": "Tách scan -> BBLM / Ghi chép (TBKQ/Edoc)",
    "bckq": "Dựng BCKQ",
    "tbkq": "Dựng TBKQ + tem tủ + Excel tính toán",
    "pdf": "Xuất PDF",
    "qa": "Kiểm tra QA",
}

_EXCEL_CELL_LIMIT = 32000  # Excel giới hạn 32767 ký tự/ô


# ---------------------------------------------------------------- tiện ích băm

def sha256_file(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def sha256_text(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8", errors="replace")).hexdigest()


def combined_files_hash(paths: Iterable[str], extra: str = "") -> str:
    """Mã băm gộp của nhiều file (theo tên + nội dung) + chuỗi phụ (vd. dpi, mã hồ sơ)."""
    h = hashlib.sha256()
    for p in sorted(paths):
        h.update(os.path.basename(p).encode("utf-8", errors="replace"))
        h.update(sha256_file(p).encode())
    h.update(extra.encode("utf-8", errors="replace"))
    return h.hexdigest()


def _now() -> str:
    return datetime.now().strftime("%Y-%m-%dT%H:%M:%S")


def _warn(msg: str) -> None:
    print(f"[Checkpoint] {msg}")


# ------------------------------------------------------------------ CheckpointStore

class CheckpointStore:
    """Kho checkpoint của MỘT hồ sơ."""

    def __init__(self, dossier_dir: Optional[str] = None, thuchien_dir: Optional[str] = None,
                 force: Iterable[str] = ()):
        if thuchien_dir:
            self.thuchien_dir = os.path.abspath(thuchien_dir)
            self.base_dir = os.path.dirname(self.thuchien_dir)
        elif dossier_dir:
            self.base_dir = os.path.abspath(dossier_dir)
            self.thuchien_dir = os.path.join(self.base_dir, "ThucHien")
        else:
            raise ValueError("Cần dossier_dir hoặc thuchien_dir")
        self.dir = os.path.join(self.thuchien_dir, CHECKPOINT_DIRNAME)
        self.force = set(force)
        self._backup_done = False
        self._excel_synced = False

    # ------------------------------------------------------------ file JSON
    def path(self, name: str) -> str:
        return os.path.join(self.dir, name)

    def read_json(self, name: str, default: Any = None) -> Any:
        p = self.path(name)
        if not os.path.isfile(p):
            return default
        try:
            with open(p, "r", encoding="utf-8-sig") as f:
                return json.load(f)
        except Exception as e:
            _warn(f"Không đọc được {name} ({type(e).__name__}: {e}) - bỏ qua checkpoint này.")
            return default

    def write_json(self, name: str, data: Any) -> bool:
        try:
            os.makedirs(self.dir, exist_ok=True)
            p = self.path(name)
            tmp = p + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=1, default=str)
            os.replace(tmp, p)
            return True
        except Exception as e:
            _warn(f"Không ghi được {name} ({type(e).__name__}: {e}).")
            return False

    # ------------------------------------------------------------ cưỡng bức làm mới
    def is_forced(self, stage: str) -> bool:
        return stage in self.force or "all" in self.force

    # ------------------------------------------------------------ manifest
    def _manifest(self) -> Dict[str, Any]:
        m = self.read_json("manifest.json", None)
        if not isinstance(m, dict) or "stages" not in m:
            m = {"schema": SCHEMA_VERSION, "stages": {}}
        return m

    def _rel(self, p: str) -> str:
        try:
            return os.path.relpath(os.path.abspath(p), self.base_dir)
        except ValueError:  # khác ổ đĩa
            return os.path.abspath(p)

    def _abs(self, p: str) -> str:
        return p if os.path.isabs(p) else os.path.normpath(os.path.join(self.base_dir, p))

    def get_stage(self, stage: str) -> Optional[Dict[str, Any]]:
        return self._manifest()["stages"].get(stage)

    def invalidate_from(self, stage: str, include_self: bool = False) -> None:
        """Xóa trạng thái của các giai đoạn đứng SAU `stage` (vì chúng dựa trên dữ liệu cũ)."""
        if stage not in STAGES:
            return
        m = self._manifest()
        idx = STAGES.index(stage)
        drop = STAGES[idx if include_self else idx + 1:]
        changed = False
        for s in drop:
            if s in m["stages"]:
                del m["stages"][s]
                changed = True
        if changed:
            self.write_json("manifest.json", m)

    def should_skip(self, stage: str, inputs_hash: Optional[str] = None, check_outputs: bool = True) -> bool:
        """True nếu giai đoạn đã xong, dữ liệu đầu vào không đổi, file đầu ra còn đủ, và không bị ép chạy lại."""
        if self.is_forced(stage):
            return False
        return self.entry_reusable(self.get_stage(stage), inputs_hash, check_outputs)

    def entry_reusable(self, st: Optional[Dict[str, Any]], inputs_hash: Optional[str] = None,
                       check_outputs: bool = True) -> bool:
        """Như should_skip nhưng xét một bản ghi giai đoạn đã chụp lại trước đó (vì begin() của giai đoạn
        đứng trước sẽ xóa bản ghi các giai đoạn sau)."""
        if not st or st.get("status") != "done":
            return False
        if inputs_hash and st.get("inputs_hash") != inputs_hash:
            return False
        if check_outputs:
            for o in st.get("outputs", []):
                p = self._abs(o.get("path", ""))
                if not os.path.isfile(p) or os.path.getsize(p) == 0:
                    return False
        return True

    def restore_stage(self, stage: str, entry: Dict[str, Any]) -> None:
        """Ghi lại nguyên bản ghi của một giai đoạn được bỏ qua (do begin() giai đoạn trước đã xóa nó)."""
        m = self._manifest()
        m["stages"][stage] = entry
        self.write_json("manifest.json", m)

    def begin(self, stage: str, inputs_hash: Optional[str] = None) -> None:
        m = self._manifest()
        # Bắt đầu làm lại một giai đoạn => kết quả các giai đoạn sau nó (dựng từ dữ liệu cũ) hết hiệu lực.
        if stage in STAGES:
            for s in STAGES[STAGES.index(stage) + 1:]:
                m["stages"].pop(s, None)
        m["stages"][stage] = {
            "status": "running",
            "started_at": _now(),
            "_t0": time.time(),
            "inputs_hash": inputs_hash,
            "outputs": [],
        }
        self.write_json("manifest.json", m)

    def finish(self, stage: str, outputs: Iterable[Optional[str]] = (), data: Optional[Dict[str, Any]] = None,
               note: str = "", status: str = "done") -> None:
        m = self._manifest()
        st = m["stages"].get(stage) or {"started_at": _now(), "_t0": time.time()}
        t0 = st.pop("_t0", None)
        outs = []
        for p in outputs:
            if p and os.path.isfile(p):
                outs.append({"path": self._rel(p), "size": os.path.getsize(p)})
        st.update({
            "status": status,
            "finished_at": _now(),
            "seconds": round(time.time() - t0, 1) if t0 else None,
            "outputs": outs,
            "data": data or {},
            "note": note,
            "error": None,
        })
        m["stages"][stage] = st
        self.write_json("manifest.json", m)

    def fail(self, stage: str, error: BaseException) -> None:
        m = self._manifest()
        st = m["stages"].get(stage) or {"started_at": _now()}
        t0 = st.pop("_t0", None)
        st.update({
            "status": "failed",
            "finished_at": _now(),
            "seconds": round(time.time() - t0, 1) if t0 else None,
            "error": f"{type(error).__name__}: {error}",
        })
        m["stages"][stage] = st
        self.write_json("manifest.json", m)

    @contextmanager
    def stage(self, name: str, inputs_hash: Optional[str] = None, outputs: Iterable[Optional[str]] = ()):
        """with store.stage("ocr", h, outputs=[...]): ...  (tự ghi done/failed)."""
        self.begin(name, inputs_hash)
        try:
            yield
        except BaseException as e:
            self.fail(name, e)
            raise
        else:
            self.finish(name, outputs=list(outputs))

    def force_refresh(self, stages: Iterable[str]) -> None:
        """Dùng cho --force-stage: bỏ trạng thái của giai đoạn chỉ định và mọi giai đoạn sau nó."""
        names = [s for s in stages if s in STAGES]
        if names:
            self.invalidate_from(min(names, key=STAGES.index), include_self=True)

    def format_status(self) -> str:
        m = self._manifest()
        icons = {"done": "[xong]   ", "failed": "[LỖI]    ", "running": "[dở dang]", "skipped": "[bỏ qua] "}
        lines = [f"Tiến độ hồ sơ: {self.base_dir}", f"Checkpoint:    {self.dir}", ""]
        # Việc tiếp theo: giai đoạn lỗi/dở dang đầu tiên; nếu không có thì giai đoạn chưa chạy đứng sau giai đoạn xong cuối cùng.
        finished = [i for i, s in enumerate(STAGES) if (m["stages"].get(s) or {}).get("status") in ("done", "skipped")]
        last_done = max(finished) if finished else -1
        next_todo = next((s for s in STAGES if (m["stages"].get(s) or {}).get("status") in ("failed", "running")), None)
        if next_todo is None:
            next_todo = next((s for i, s in enumerate(STAGES) if i > last_done and s not in m["stages"]), None)
        for s in STAGES:
            st = m["stages"].get(s)
            label = STAGE_LABELS.get(s, s)
            if not st:
                lines.append(f"  [chưa]    {s:<11} {label}")
                continue
            status = st.get("status", "?")
            extra = []
            if st.get("finished_at"):
                extra.append(st["finished_at"].replace("T", " "))
            if st.get("seconds") is not None:
                extra.append(f"{st['seconds']}s")
            if st.get("note"):
                extra.append(st["note"])
            if status == "failed" and st.get("error"):
                extra.append(st["error"])
            lines.append(f"  {icons.get(status, '[?]')} {s:<11} {label}" + (f"  ({'; '.join(extra)})" if extra else ""))
        lines.append("")
        lines.append(f"Việc tiếp theo: {next_todo}" if next_todo else "Mọi giai đoạn đã xong.")
        return "\n".join(lines)

    # ------------------------------------------------------------ cache OCR file scan
    def _sync_excel_edits(self) -> None:
        """Nếu ocr.xlsx được lưu SAU ocr.json (người dùng sửa tay) thì nạp phần sửa vào ocr.json."""
        if self._excel_synced:
            return
        self._excel_synced = True
        xp, jp = self.path("ocr.xlsx"), self.path("ocr.json")
        if not (os.path.isfile(xp) and os.path.isfile(jp)):
            return
        if os.stat(xp).st_mtime_ns <= os.stat(jp).st_mtime_ns:
            return
        try:
            import openpyxl
            data = self.read_json("ocr.json", None)
            if not data:
                return
            wb = openpyxl.load_workbook(xp, read_only=True, data_only=True)
            ws = wb["OCR"]
            n_changed = 0
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or row[0] is None or row[1] is None:
                    continue
                fname, page, text = str(row[0]), int(row[1]), ("" if row[3] is None else str(row[3]))
                entry = data.get("files", {}).get(fname)
                if not entry:
                    continue
                for pg in entry["pages"]:
                    if pg["page"] == page:
                        old = pg["text"]
                        if len(old) <= _EXCEL_CELL_LIMIT and text.strip() != old.strip():
                            pg["text"] = text
                            pg["edited"] = True
                            n_changed += 1
            wb.close()
            if n_changed:
                self._save_ocr_data(data)
                _warn(f"Đã nạp {n_changed} trang OCR bạn sửa tay trong ocr.xlsx.")
        except Exception as e:
            _warn(f"Không nạp được chỉnh sửa từ ocr.xlsx ({type(e).__name__}: {e}).")

    def get_ocr(self, file_path: str, dpi: int, force: bool = False) -> Optional[List[str]]:
        """Trả về danh sách văn bản theo trang nếu có cache hợp lệ (file chưa đổi, cùng dpi); ngược lại None."""
        if force or self.is_forced("ocr"):
            return None
        self._sync_excel_edits()
        data = self.read_json("ocr.json", None)
        if not data or data.get("schema") != SCHEMA_VERSION:
            return None
        entry = data.get("files", {}).get(os.path.basename(file_path))
        if not entry or entry.get("dpi") != dpi:
            return None
        try:
            if entry.get("sha256") != sha256_file(file_path):
                _warn(f"{os.path.basename(file_path)} đã thay đổi so với lúc OCR -> OCR lại.")
                return None
        except OSError:
            return None
        pages = sorted(entry.get("pages", []), key=lambda x: x["page"])
        if not pages or [p["page"] for p in pages] != list(range(1, len(pages) + 1)):
            return None
        return [p["text"] for p in pages]

    def put_ocr(self, file_path: str, pages: List[str], dpi: int, engine: str = OCR_ENGINE) -> None:
        """Lưu kết quả OCR của một file. Không lưu nếu kết quả trông như OCR hỏng."""
        failed = [t for t in pages if t.startswith("ERROR:") or t == "FILE_NOT_FOUND"]
        if failed or not any(t.strip() for t in pages):
            _warn(f"OCR {os.path.basename(file_path)} có trang lỗi/rỗng toàn bộ -> KHÔNG lưu checkpoint.")
            return
        data = self.read_json("ocr.json", None)
        if not data or data.get("schema") != SCHEMA_VERSION:
            data = {"schema": SCHEMA_VERSION, "files": {}}
        name = os.path.basename(file_path)
        old = data["files"].get(name)
        if old and any(p.get("edited") for p in old.get("pages", [])) and not self._backup_done:
            try:
                shutil.copy2(self.path("ocr.json"), self.path("ocr.before_reocr.json"))
                _warn("OCR lại sẽ ghi đè phần bạn sửa tay; bản cũ lưu ở ocr.before_reocr.json.")
            except OSError:
                pass
            self._backup_done = True
        data["files"][name] = {
            "sha256": sha256_file(file_path),
            "size": os.path.getsize(file_path),
            "dpi": dpi,
            "engine": engine,
            "ocr_at": _now(),
            "pages": [{"page": i + 1, "text": t, "edited": False} for i, t in enumerate(pages)],
        }
        self._save_ocr_data(data)

    def _save_ocr_data(self, data: Dict[str, Any]) -> None:
        # Ghi Excel trước, JSON sau -> mtime(json) > mtime(xlsx) nghĩa là Excel chưa bị sửa tay.
        self.export_ocr_excel(data)
        self.write_json("ocr.json", data)

    def export_ocr_excel(self, data: Optional[Dict[str, Any]] = None) -> Optional[str]:
        data = data or self.read_json("ocr.json", None)
        if not data:
            return None
        try:
            import openpyxl
            from openpyxl.styles import Alignment, Font, PatternFill
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "OCR"
            ws.append(["File", "Trang", "Số ký tự", "Văn bản OCR (sửa ở cột này)", "Đã sửa tay"])
            for c in ws[1]:
                c.font = Font(bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor="305496")
            for fname, entry in data.get("files", {}).items():
                for pg in entry.get("pages", []):
                    t = pg["text"]
                    ws.append([fname, pg["page"], len(t), t[:_EXCEL_CELL_LIMIT], "x" if pg.get("edited") else ""])
            ws.column_dimensions["A"].width = 34
            ws.column_dimensions["B"].width = 7
            ws.column_dimensions["C"].width = 10
            ws.column_dimensions["D"].width = 110
            ws.column_dimensions["E"].width = 11
            ws.freeze_panes = "A2"
            for row in ws.iter_rows(min_row=2, min_col=4, max_col=4):
                row[0].alignment = Alignment(wrap_text=True, vertical="top")
            info = wb.create_sheet("Thông tin")
            info.append(["File", "SHA-256", "DPI", "Công cụ", "OCR lúc", "Số trang"])
            for c in info[1]:
                c.font = Font(bold=True)
            for fname, e in data.get("files", {}).items():
                info.append([fname, e.get("sha256"), e.get("dpi"), e.get("engine"), e.get("ocr_at"), len(e.get("pages", []))])
            info.column_dimensions["A"].width = 34
            info.column_dimensions["B"].width = 66
            os.makedirs(self.dir, exist_ok=True)
            out = self.path("ocr.xlsx")
            wb.save(out)
            return out
        except PermissionError:
            _warn("ocr.xlsx đang mở trong Excel nên chưa cập nhật được (ocr.json vẫn đã lưu). Đóng file rồi chạy lại để làm mới.")
        except Exception as e:
            _warn(f"Không xuất được ocr.xlsx ({type(e).__name__}: {e}).")
        return None

    # ------------------------------------------------------------ cache OCR đầu trang (scan_splitter)
    @staticmethod
    def file_hash(path: str) -> str:
        return sha256_file(path)

    @staticmethod
    def header_key(pdf_sha: str, page_index: int, dpi: int, top_fraction: float) -> str:
        return f"{pdf_sha[:20]}:p{page_index}:dpi{dpi}:top{top_fraction}"

    def get_header_ocr(self, key: str) -> Optional[str]:
        if self.is_forced("ocr"):
            return None
        return self.read_json("header_ocr.json", {}).get(key)

    def put_header_ocr(self, items: Dict[str, str]) -> None:
        good = {k: v for k, v in items.items() if v and not v.startswith("ERROR:") and v != "FILE_NOT_FOUND"}
        if not good:
            return
        cur = self.read_json("header_ocr.json", {}) or {}
        cur.update(good)
        self.write_json("header_ocr.json", cur)


def parse_force_stages(value: Optional[str], force_ocr: bool = False) -> List[str]:
    """'ocr,pdf' -> ['ocr','pdf']; kiểm tra tên hợp lệ."""
    names = [s.strip() for s in (value or "").split(",") if s.strip()]
    if force_ocr and "ocr" not in names:
        names.append("ocr")
    bad = [s for s in names if s not in STAGES and s != "all"]
    if bad:
        raise ValueError(f"Giai đoạn không hợp lệ: {bad}. Hợp lệ: {', '.join(STAGES)}, all")
    return names
