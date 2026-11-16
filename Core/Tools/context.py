"""Context nhiều phiên không chồng log (AGENTS mục 16) — mỗi file chỉ một phiên ghi.

Trước 12/11/2026 mọi phiên cùng sửa một file PROJECT_CONTEXT.md (117 KB) → hai phiên chạy song song
ghi đè lên nhau, đầu phiên phải đọc cả file. Nay trạng thái động tách thành file nhỏ:

    Context/HoSo/<6 số>.md       một hồ sơ đang mở: trạng thái, bước tiếp, điểm cần người dùng quyết, nhật ký
    Context/Viec/<ID>.md         một lỗi (E<n>) hoặc một việc hệ thống (V<n>) đang mở
    Context/Phien/<ngày-giờ>-<chủ đề>.md   nhật ký hệ thống của MỘT phiên (sửa Core, AGENTS, công cụ)
    Context/_khoa/<tên>.lock     khóa: hồ sơ đang do phiên nào làm
    ContextLuuTru/               lưu trữ nguyên văn (HoSo/, LoiDaSua.md, NhatKy/) — không đọc đầu phiên

Ai ghi file nào: file hồ sơ — chỉ phiên đang giữ khóa hồ sơ đó; file phiên — chỉ phiên tạo ra nó;
file việc / lỗi — phiên đang xử lý việc đó. Không file nào là "của chung" nên không có gì để chồng.

Lệnh:

    python Core/Tools/context.py dau-phien              # dọn (ghi) + in tóm tắt + nhãn phiên gợi ý
    python Core/Tools/context.py xem                    # chỉ in tóm tắt, không ghi gì
    python Core/Tools/context.py khoa 137058 --ai <nhãn> [--ghi-chu "..."]   # giữ / gia hạn khóa
    python Core/Tools/context.py tra  137058 --ai <nhãn>                     # trả khóa
    python Core/Tools/context.py moi 139999 --khach "Công ty …" --doi-tuong "BSC, 1 chứng thư" --ai <nhãn>
    python Core/Tools/context.py mo  134792 --ai <nhãn>   # làm lại hồ sơ đã lưu trữ
    python Core/Tools/context.py phien "chủ đề" --ai <nhãn>     # tạo file phiên, in đường dẫn
    python Core/Tools/context.py viec "tiêu đề" [--loi] [--muc Thấp] --ai <nhãn>   # cấp ID V<n> / E<n>
    python Core/Tools/context.py so                     # sổ mọi hồ sơ (đang mở + lưu trữ)
    python Core/Tools/context.py don [--ghi]            # lưu trữ phần đã xong (không --ghi = chạy thử)
    python Core/Tools/context.py dong 137058 [--ghi]    # đóng hồ sơ ngay

Tạo file bằng O_EXCL (hai phiên cùng tạo một tên thì một phiên thất bại, không ghi đè); ghi lại file
bằng file tạm + os.replace (không có file viết dở). Khóa quá 12 giờ không gia hạn coi là khóa cũ,
phiên khác được nhận lại (in cảnh báo).
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import re
import secrets
import socket
import sys
import time
import unicodedata

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
# Đổi chỗ để context bằng biến môi trường (vd đặt trên ổ dùng chung cho nhiều máy, hoặc thư mục thử):
CTX = os.environ.get("LAMKQ_CONTEXT_DIR") or os.path.join(ROOT, "Context")
HOSO = os.path.join(CTX, "HoSo")
VIEC = os.path.join(CTX, "Viec")
PHIEN = os.path.join(CTX, "Phien")
KHOA = os.path.join(CTX, "_khoa")
ARCHIVE = os.environ.get("LAMKQ_CONTEXT_ARCHIVE") or os.path.join(ROOT, "ContextLuuTru")

DATE_RE = re.compile(r"(?<!\d)(\d{1,2})/(\d{1,2})/(\d{4})(?!\d)")
CODE_RE = re.compile(r"(?<!\d)(?:Q3\d{2}[A-D])?(\d{6})(?!\d)")
FIELD_RE = re.compile(r"^- \*\*(.+?):\*\*\s?(.*)$")
LOCK_STALE_HOURS = 12      # khóa không gia hạn quá ngần này giờ → phiên khác được nhận lại
DOSSIER_IDLE_DAYS = 7      # hồ sơ không đụng tới ≥ 7 ngày → tự xem là xong (người dùng chốt 05/11/2026)
SESSION_KEEP_DAYS = 7      # file phiên cũ hơn ngần này ngày → lưu trữ
RECENT_DAYS = 3            # `xem` liệt kê phiên trong ngần này ngày
Z_ROOT = os.environ.get("LAMKQ_Z_ROOT", r"Z:\9.NV2 - MOI TRUONG\1.Ho So Moi Truong\2026")
Z_SUBDIRS = ("", "0.VIMCERTS078", "0.1.Tân Hiệp")
SKIP_DIRS = {"_checkpoint", "_optimized", "__pycache__"}
LOG_HEADER = ["| Ngày | AI | Nội dung |", "|---|---|---|"]
DONE_RE = re.compile(r"^(Fixed|Done)\s*\(")
NOT_DONE_RE = re.compile(r"một phần|Partly|trong code|chưa|chờ|In Progress", re.I)


# ---------------------------------------------------------------- đọc / ghi an toàn
def read_text(path: str) -> str:
    with open(path, encoding="utf-8") as f:
        return f.read()


def write_atomic(path: str, text: str) -> None:
    """Ghi qua file tạm cùng thư mục rồi os.replace — không bao giờ để lại file viết dở."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = f"{path}.{os.getpid()}.{secrets.token_hex(3)}.tmp"
    with open(tmp, "w", encoding="utf-8", newline="\n") as f:
        f.write(text)
    os.replace(tmp, path)


def create_new(path: str, text: str) -> None:
    """Tạo file mới; đã có thì FileExistsError (không ghi đè)."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    fd = os.open(path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as f:
        f.write(text)


def rel(path: str) -> str:
    try:
        return os.path.relpath(path, ROOT).replace("\\", "/")
    except ValueError:                    # khác ổ đĩa (context đặt ngoài workspace)
        return path


def md_files(folder: str) -> list[str]:
    if not os.path.isdir(folder):
        return []
    return sorted(os.path.join(folder, n) for n in os.listdir(folder) if n.endswith(".md"))


# ---------------------------------------------------------------- phân tích file
def title_of(text: str) -> str:
    for line in text.split("\n"):
        if line.startswith("# "):
            return line[2:].strip()
    return ""


def fields_of(text: str) -> dict[str, str]:
    """Các dòng `- **Khóa:** giá trị` trước tiêu đề `## ` đầu tiên."""
    out = {}
    for line in text.split("\n"):
        if line.startswith("## "):
            break
        m = FIELD_RE.match(line)
        if m:
            out[m.group(1).strip()] = m.group(2).strip()
    return out


def short(text: str, n: int) -> str:
    text = re.sub(r"\s+", " ", text.replace("<br>", " ")).strip()
    return text if len(text) <= n else text[:n - 1].rstrip() + "…"


def plain(text: str) -> str:
    return text.replace("**", "").strip()


def blank(value: str) -> bool:
    return plain(value).strip(" .").lower() in ("", "không", "khong", "—", "-", "chưa có")


def to_date(m: re.Match) -> dt.date | None:
    d, mo, y = map(int, m.groups())
    try:
        return dt.date(y, mo, d)
    except ValueError:
        return None


def context_dates(text: str) -> list[dt.date]:
    """Ngày trong dòng Trạng thái và cột Ngày của bảng nhật ký."""
    out = []
    for line in text.split("\n"):
        src = None
        if line.startswith("- **Trạng thái:**"):
            src = line
        elif re.match(r"^\|\s*\d{1,2}/\d{1,2}/\d{4}", line):
            src = line.split("|")[1]
        if src:
            out += [d for d in (to_date(m) for m in DATE_RE.finditer(src)) if d]
    return out


def norm_code(name: str) -> str:
    m = CODE_RE.fullmatch(name.strip())
    return m.group(1) if m else name.strip()


def slugify(text: str, limit: int = 40) -> str:
    text = text.replace("đ", "d").replace("Đ", "D")
    text = unicodedata.normalize("NFD", text)
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    text = re.sub(r"[^A-Za-z0-9]+", "-", text).strip("-").lower()
    return text[:limit].strip("-") or "phien"


def demote(body: str, by: int) -> str:
    """Lùi cấp tiêu đề markdown khi chép một file vào trong file lưu trữ."""
    return "\n".join(("#" * by + l) if re.match(r"^#{1,5} ", l) else l for l in body.split("\n"))


def suggest_label() -> str:
    return f"{socket.gethostname()}-{dt.datetime.now():%H%M}-{secrets.token_hex(2)}"


# ---------------------------------------------------------------- khóa
def lock_path(name: str) -> str:
    return os.path.join(KHOA, f"{name}.lock")


def read_lock(name: str) -> dict | None:
    try:
        return json.loads(read_text(lock_path(name)))
    except (OSError, ValueError):
        return None


def lock_age_hours(lock: dict) -> float:
    try:
        t = dt.datetime.fromisoformat(lock.get("gia_han") or lock["bat_dau"])
    except (KeyError, ValueError):
        return 1e9
    return (dt.datetime.now() - t).total_seconds() / 3600


def lock_text(lock: dict) -> str:
    age = lock_age_hours(lock)
    old = "  [KHÓA CŨ]" if age >= LOCK_STALE_HOURS else ""
    note = f" — {lock['ghi_chu']}" if lock.get("ghi_chu") else ""
    return f"{lock.get('ai', '?')} (máy {lock.get('may', '?')}, {age:.1f} giờ trước){note}{old}"


def all_locks() -> dict[str, dict]:
    out = {}
    if os.path.isdir(KHOA):
        for n in sorted(os.listdir(KHOA)):
            if n.endswith(".lock") and not n.startswith("_"):
                lock = read_lock(n[:-5])
                if lock:
                    out[n[:-5]] = lock
    return out


def cmd_khoa(name: str, ai: str, note: str = "", force: bool = False) -> int:
    name = norm_code(name)
    if not re.fullmatch(r"[\w.\-]+", name):
        print(f"Tên khóa không hợp lệ: {name!r}")
        return 1
    now = dt.datetime.now().isoformat(timespec="seconds")
    data = {"ten": name, "ai": ai, "may": socket.gethostname(), "bat_dau": now, "gia_han": now, "ghi_chu": note}
    path = lock_path(name)
    try:
        create_new(path, json.dumps(data, ensure_ascii=False, indent=1))
        print(f"[KHÓA] {name}: đã giữ cho {ai}")
        return 0
    except FileExistsError:
        pass
    cur = read_lock(name) or {}
    if cur.get("ai") == ai:
        cur.update(gia_han=now, ghi_chu=note or cur.get("ghi_chu", ""))
        write_atomic(path, json.dumps(cur, ensure_ascii=False, indent=1))
        print(f"[KHÓA] {name}: gia hạn cho {ai}")
        return 0
    if not force and lock_age_hours(cur) < LOCK_STALE_HOURS:
        print(f"[KHÓA BẬN] {name}: đang do {lock_text(cur)} giữ. KHÔNG làm hồ sơ này — báo người dùng. "
              f"(Người dùng xác nhận phiên kia đã dừng thì chạy lại với --ep.)")
        return 2
    write_atomic(path, json.dumps(data, ensure_ascii=False, indent=1))
    time.sleep(0.3)                       # hai phiên cùng nhận lại khóa cũ: phiên ghi sau thắng
    if (read_lock(name) or {}).get("ai") != ai:
        print(f"[KHÓA BẬN] {name}: phiên khác vừa nhận khóa.")
        return 2
    print(f"[KHÓA] {name}: nhận lại từ {lock_text(cur)} → {ai}")
    return 0


def cmd_tra(name: str, ai: str, force: bool = False) -> int:
    name = norm_code(name)
    cur = read_lock(name)
    if cur is None:
        print(f"[KHÓA] {name}: không có khóa")
        return 0
    if cur.get("ai") != ai and not force:
        print(f"[KHÓA] {name}: do {lock_text(cur)} giữ, không phải {ai} — không trả. (--ep để gỡ)")
        return 2
    os.remove(lock_path(name))
    print(f"[KHÓA] {name}: đã trả")
    return 0


# ---------------------------------------------------------------- tạo file
def dossier_path(code: str) -> str:
    return os.path.join(HOSO, f"{code}.md")


def archive_dossier_path(code: str) -> str:
    return os.path.join(ARCHIVE, "HoSo", f"{code}.md")


def dossier_template(code: str, client: str, scope: str, status: str, extra: str = "") -> str:
    return "\n".join([
        f"# Q326D{code} — {client}", "",
        f"- **Trạng thái:** {status}",
        "- **Bước tiếp:** ",
        "- **Cần người dùng quyết:** không",
        f"- **Đối tượng / số chứng thư:** {scope}",
        "- **Ngày GĐ → ban hành:** ",
        f"- **Cách dựng:** `build_{code}.py`",
        *([extra] if extra else []), "",
        "## Nhật ký (mới nhất ở trên)", "", *LOG_HEADER, ""])


def cmd_moi(code: str, client: str, scope: str, ai: str | None) -> int:
    code = norm_code(code)
    if not re.fullmatch(r"\d{6}", code):
        print(f"Mã hồ sơ phải có 6 số: {code!r}")
        return 1
    if os.path.exists(archive_dossier_path(code)):
        print(f"{code} đã có lưu trữ ({rel(archive_dossier_path(code))}) — dùng `mo {code}`.")
        return 1
    if ai and cmd_khoa(code, ai) != 0:
        return 2
    today = dt.date.today()
    try:
        create_new(dossier_path(code), dossier_template(code, client, scope, f"Đang làm ({today:%d/%m/%Y})"))
    except FileExistsError:
        print(f"Đã có {rel(dossier_path(code))} — sửa file đó.")
        return 0
    print(f"Đã tạo {rel(dossier_path(code))}")
    return 0


def cmd_mo(code: str, ai: str | None) -> int:
    code = norm_code(code)
    if os.path.exists(dossier_path(code)):
        if ai and cmd_khoa(code, ai) != 0:
            return 2
        print(f"Hồ sơ đang mở: {rel(dossier_path(code))}")
        return 0
    arch = archive_dossier_path(code)
    if not os.path.exists(arch):
        print(f"Chưa có context cho {code} — hồ sơ mới thì dùng `moi {code} --khach … --doi-tuong …`.")
        return 1
    if ai and cmd_khoa(code, ai) != 0:
        return 2
    today = dt.date.today()
    title = re.sub(r"^Lưu trữ hồ sơ\s*", "", title_of(read_text(arch)))
    client = re.sub(r"^(?:Q3\d{2}[A-D])?\d{6}\S*\s*—\s*", "", title)
    extra = f"- **Lưu trữ trước đó:** [{rel(arch)}]({rel(arch)}) — đọc trước khi làm (quyết định cũ của người dùng nằm ở đó)"
    try:
        create_new(dossier_path(code), dossier_template(code, client, "(xem lưu trữ)",
                                                         f"Mở lại {today:%d/%m/%Y}", extra))
    except FileExistsError:
        pass
    print(f"Đã mở lại {rel(dossier_path(code))}. Đọc lưu trữ: {rel(arch)}")
    return 0


def cmd_phien(topic: str, ai: str) -> int:
    now = dt.datetime.now()
    base = f"{now:%Y%m%d-%H%M}-{slugify(topic)}"
    text = "\n".join([
        f"# Phiên {now:%d/%m/%Y %H:%M} — {ai} — {topic}", "",
        "- **Tóm tắt:** ",
        "- **Hồ sơ liên quan:** ",
        "- **File đổi:** ", "",
        "## Nhật ký hệ thống (việc không thuộc riêng hồ sơ nào)", "", *LOG_HEADER, ""])
    for n in range(1, 50):
        path = os.path.join(PHIEN, base + (f"-{n}" if n > 1 else "") + ".md")
        try:
            create_new(path, text)
            print(rel(path))
            return 0
        except FileExistsError:
            continue
    return 1


def known_ids(prefix: str) -> set[int]:
    ids = set()
    for p in md_files(VIEC):
        m = re.fullmatch(rf"{prefix}(\d+)\.md", os.path.basename(p))
        if m:
            ids.add(int(m.group(1)))
    for root, _, names in os.walk(ARCHIVE):
        for n in names:
            if not n.endswith(".md") or os.sep + "HoSo" in root:
                continue
            for m in re.finditer(rf"^(?:\|\s*|#+\s*){prefix}(\d+)\b", read_text(os.path.join(root, n)), re.M):
                ids.add(int(m.group(1)))
    return ids


def cmd_viec(title: str, is_error: bool, level: str, ai: str) -> int:
    prefix = "E" if is_error else "V"
    n = max(known_ids(prefix), default=0) + 1
    today = dt.date.today()
    while True:
        text = "\n".join([
            f"# {prefix}{n} — {title}", "",
            *([f"- **Mức độ:** {level}"] if is_error else []),
            f"- **Trạng thái:** Open ({today:%d/%m/%Y}, {ai})", "",
            "## Mô tả", "", ""])
        path = os.path.join(VIEC, f"{prefix}{n}.md")
        try:
            create_new(path, text)
            print(rel(path))
            return 0
        except FileExistsError:           # phiên khác vừa lấy số này
            n += 1


# ---------------------------------------------------------------- tóm tắt đầu phiên
def session_date(path: str) -> dt.date | None:
    m = re.match(r"(\d{4})(\d{2})(\d{2})-", os.path.basename(path))
    try:
        return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else None
    except ValueError:
        return None


def id_key(path: str):
    m = re.match(r"([A-Z]+)(\d+)", os.path.basename(path))
    return (m.group(1), int(m.group(2))) if m else ("~", 0)


def cmd_xem(full: bool = False) -> int:
    now = dt.datetime.now()
    locks = all_locks()
    w = 100000 if full else 170
    print(f"CONTEXT {now:%d/%m/%Y %H:%M} — máy {socket.gethostname()} — nhãn phiên gợi ý: {suggest_label()}")
    print("Trước khi làm hồ sơ: `khoa <6 số> --ai <nhãn>` (bận → dừng, báo người dùng), đọc Context/HoSo/<6 số>.md. "
          "Xong: cập nhật file hồ sơ đó, `tra <6 số> --ai <nhãn>`. Sửa Core/AGENTS: `phien \"<chủ đề>\" --ai <nhãn>`.")

    files = md_files(HOSO)
    print(f"\nHỒ SƠ ĐANG MỞ ({len(files)}) — Context/HoSo/<6 số>.md")
    asks = []
    for p in files:
        code = os.path.basename(p)[:-3]
        text = read_text(p)
        f = fields_of(text)
        lk = f"  [KHÓA: {lock_text(locks[code])}]" if code in locks else ""
        client = re.sub(r"^\S+(?: \([^)]*\))?\s*—\s*", "", title_of(text))
        print(f"  {code}  {short(client, 60)}{lk}")
        print(f"          {short(plain(f.get('Trạng thái', '?')), w)}")
        if not blank(f.get("Bước tiếp", "")):
            print(f"          Bước tiếp: {short(plain(f['Bước tiếp']), w)}")
        if not blank(f.get("Cần người dùng quyết", "")):
            asks.append(f"  {code}: {short(plain(f['Cần người dùng quyết']), w + 60)}")

    other = {n: l for n, l in locks.items() if not os.path.exists(dossier_path(n))}
    if other:
        print("\nKHÓA KHÁC")
        for n, l in other.items():
            print(f"  {n}: {lock_text(l)}")

    items = sorted(md_files(VIEC), key=id_key)
    print(f"\nLỖI / VIỆC HỆ THỐNG ĐANG MỞ ({len(items)}) — Context/Viec/<ID>.md")
    for p in items:
        text = read_text(p)
        f = fields_of(text)
        t = title_of(text)
        level = f" [{f['Mức độ']}]" if f.get("Mức độ") else ""
        print(f"  {short(t, 110 if not full else w)}{level} — {short(plain(f.get('Trạng thái', '?')), 60 if not full else w)}")
        if not blank(f.get("Cần người dùng quyết", "")):
            asks.append(f"  {os.path.basename(p)[:-3]}: {short(plain(f['Cần người dùng quyết']), w + 60)}")

    print("\nCẦN NGƯỜI DÙNG QUYẾT")
    print("\n".join(asks) if asks else "  (không có)")

    cutoff = now.date() - dt.timedelta(days=RECENT_DAYS)
    recent = [p for p in md_files(PHIEN) if (session_date(p) or dt.date.min) >= cutoff]
    print(f"\nPHIÊN {RECENT_DAYS} NGÀY GẦN ĐÂY ({len(recent)}) — Context/Phien/")
    for p in reversed(recent):
        text = read_text(p)
        summary = fields_of(text).get("Tóm tắt", "")
        print(f"  {os.path.basename(p)}: {short(plain(summary) or title_of(text), w + 60)}")
    print("\nLưu trữ: ContextLuuTru/ (không đọc đầu phiên). Sổ mọi hồ sơ: `context.py so`.")
    return 0


def cmd_so() -> int:
    rows = {}
    for p in md_files(os.path.join(ARCHIVE, "HoSo")):
        code = os.path.basename(p)[:-3]
        rows[code] = ("lưu trữ", re.sub(r"^Lưu trữ hồ sơ\s*", "", title_of(read_text(p))))
    for p in md_files(HOSO):
        code = os.path.basename(p)[:-3]
        again = " (mở lại)" if code in rows else ""
        rows[code] = ("ĐANG MỞ" + again, title_of(read_text(p)))
    for code in sorted(rows):
        print(f"{code}  {rows[code][0]:<18} {short(rows[code][1], 150)}")
    return 0


# ---------------------------------------------------------------- lưu trữ
def _latest_mtime(folder: str, subdirs=None) -> dt.date | None:
    roots = [os.path.join(folder, s) for s in subdirs] if subdirs else [folder]
    best = None
    for r in roots:
        if not os.path.isdir(r):
            continue
        for dp, dns, fns in os.walk(r):
            dns[:] = [d for d in dns if d not in SKIP_DIRS]
            for fn in fns:
                if fn.startswith("~$") or fn.lower() == "thumbs.db":
                    continue
                try:
                    t = os.path.getmtime(os.path.join(dp, fn))
                except OSError:
                    continue
                if best is None or t > best:
                    best = t
    return dt.date.fromtimestamp(best) if best else None


def _z_folders(code: str) -> list[str]:
    out = []
    for sub in Z_SUBDIRS:
        base = os.path.join(Z_ROOT, sub) if sub else Z_ROOT
        try:
            out += [os.path.join(base, n) for n in os.listdir(base)
                    if code in n and os.path.isdir(os.path.join(base, n))]
        except OSError:
            continue
    return out


def last_touched(code: str, text: str, use_files: bool, idle_days: int, today: dt.date):
    """(ngày đụng tới gần nhất, nguồn). Chỉ quét file Input / Z khi context đã cũ ≥ idle_days (quét Z chậm)."""
    ds = context_dates(text)
    best = (max(ds), "context") if ds else (None, "không có ngày")
    if best[0] is None or (today - best[0]).days < idle_days or not use_files:
        return best
    cands = [best]
    d = _latest_mtime(os.path.join(ROOT, "Input", code))
    if d:
        cands.append((d, "file Input"))
    for zf in _z_folders(code):
        d = _latest_mtime(zf, ("ThucHien", "TBKQ", "pdf"))
        if d:
            cands.append((d, "file Z"))
    return max(cands)


def append_block(path: str, header: str, block: str) -> bool:
    """Thêm khối vào cuối file lưu trữ (tạo file nếu chưa có); khối đã có nguyên văn thì bỏ qua."""
    doc = read_text(path) if os.path.exists(path) else header
    if block.strip() in doc:
        return False
    write_atomic(path, doc.rstrip("\n") + "\n\n" + block.strip("\n") + "\n")
    return True


def split_title(text: str) -> tuple[str, str]:
    lines = text.split("\n")
    i = next((k for k, l in enumerate(lines) if l.startswith("# ")), None)
    if i is None:
        return "", text
    return lines[i][2:].strip(), "\n".join(lines[i + 1:]).strip("\n")


def archive_dossier(code: str, today: dt.date) -> str:
    src = dossier_path(code)
    title, body = split_title(read_text(src))
    dst = archive_dossier_path(code)
    header = "\n".join([
        f"# Lưu trữ hồ sơ {title}", "",
        f"> Hồ sơ đã xong, chuyển khỏi `Context/HoSo/` (AGENTS mục 16). Chỉ đọc khi người dùng mở lại hồ sơ này "
        f"(`python Core/Tools/context.py mo {code}`)."])
    append_block(dst, header, f"## Lưu trữ {today:%d/%m/%Y}\n\n{demote(body, 1)}")
    os.remove(src)
    if read_lock(code):
        os.remove(lock_path(code))
    return rel(dst)


def archive_item(path: str, today: dt.date) -> str:
    name = os.path.basename(path)[:-3]
    title, body = split_title(read_text(path))
    if name.startswith("E"):
        dst = os.path.join(ARCHIVE, "LoiDaSua.md")
        header = "# Lưu trữ lỗi đã sửa\n\n> Lỗi đã Fixed, chép nguyên văn (AGENTS mục 16). Tái phát thì mở ID mới, dẫn chiếu ID cũ."
    else:
        dst = os.path.join(ARCHIVE, "NhatKy", f"{today:%Y-%m}.md")
        header = f"# Lưu trữ nhật ký hệ thống {today:%Y-%m}\n\n> Chép nguyên văn khỏi `Context/` (AGENTS mục 16)."
    append_block(dst, header, f"### {title} (lưu trữ {today:%d/%m/%Y})\n\n{demote(body, 2)}")
    os.remove(path)
    return rel(dst)


def archive_session(path: str) -> str:
    d = session_date(path)
    title, body = split_title(read_text(path))
    dst = os.path.join(ARCHIVE, "NhatKy", f"{d:%Y-%m}.md")
    header = f"# Lưu trữ nhật ký hệ thống {d:%Y-%m}\n\n> Chép nguyên văn khỏi `Context/` (AGENTS mục 16)."
    append_block(dst, header, f"## {title}\n\n{demote(body, 1)}")
    os.remove(path)
    return rel(dst)


def item_done(text: str) -> bool:
    status = plain(fields_of(text).get("Trạng thái", ""))
    return bool(DONE_RE.match(status)) and not NOT_DONE_RE.search(status[:80]) and "[giữ]" not in status


def cmd_don(write: bool, close: list[str] | None = None, idle_days: int = DOSSIER_IDLE_DAYS,
            use_files: bool = True, quiet: bool = False) -> int:
    today = dt.date.today()
    locks = all_locks()
    plan = []       # (mô tả, hàm thực hiện)

    close = [norm_code(c) for c in (close or [])]
    for code in close:
        if not os.path.exists(dossier_path(code)):
            print(f"  [Bỏ qua] {code}: không có trong Context/HoSo/")
    for p in md_files(HOSO):
        code = os.path.basename(p)[:-3]
        lk = locks.get(code)
        busy = lk is not None and lock_age_hours(lk) < LOCK_STALE_HOURS
        if code in close:
            if busy:
                print(f"  [Bỏ qua] {code}: đang bị khóa bởi {lock_text(lk)}")
                continue
            plan.append((f"đóng hồ sơ {code}", lambda c=code: archive_dossier(c, today)))
        elif close is not None and not close and not busy:
            d, src = last_touched(code, read_text(p), use_files, idle_days, today)
            if d and (today - d).days >= idle_days:
                plan.append((f"tự đóng hồ sơ {code}: không đụng tới từ {d:%d/%m/%Y} ({src})",
                             lambda c=code: archive_dossier(c, today)))
    if not close:
        for p in sorted(md_files(VIEC), key=id_key):
            if item_done(read_text(p)):
                plan.append((f"lưu trữ {os.path.basename(p)[:-3]} (đã xong)", lambda q=p: archive_item(q, today)))
        cutoff = today - dt.timedelta(days=SESSION_KEEP_DAYS)
        for p in md_files(PHIEN):
            d = session_date(p)
            if d and d < cutoff:
                plan.append((f"lưu trữ phiên {os.path.basename(p)}", lambda q=p: archive_session(q)))

    if not plan:
        if not quiet:
            print("Không có gì để lưu trữ.")
        return 0
    if not write:
        print("Sẽ làm (chạy thử — thêm --ghi để thực hiện):")
        for desc, _ in plan:
            print(f"  {desc}")
        return 0
    guard = os.path.join(KHOA, "_don.lock")           # hai phiên không dọn cùng lúc
    try:
        if os.path.exists(guard) and time.time() - os.path.getmtime(guard) > 600:
            os.remove(guard)
        create_new(guard, socket.gethostname())
    except FileExistsError:
        print("Phiên khác đang dọn — bỏ qua lần này.")
        return 0
    try:
        for desc, fn in plan:
            try:
                print(f"  {desc} → {fn()}")
            except FileNotFoundError:
                print(f"  {desc} → phiên khác đã làm")
    finally:
        try:
            os.remove(guard)
        except OSError:
            pass
    return 0


# ---------------------------------------------------------------- CLI
def main(argv=None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="cmd", required=True)
    sub.add_parser("dau-phien", help="Dọn (ghi) rồi in tóm tắt")
    x = sub.add_parser("xem", help="In tóm tắt, không ghi gì")
    x.add_argument("--day-du", action="store_true", help="Không cắt ngắn dòng")
    for name in ("khoa", "tra"):
        k = sub.add_parser(name)
        k.add_argument("ten")
        k.add_argument("--ai", required=True, help="Nhãn phiên (lấy từ dòng đầu của `dau-phien`)")
        k.add_argument("--ep", action="store_true", help="Người dùng đã xác nhận phiên kia dừng")
        if name == "khoa":
            k.add_argument("--ghi-chu", default="")
    m = sub.add_parser("moi", help="Tạo file hồ sơ mới")
    m.add_argument("ma")
    m.add_argument("--khach", required=True)
    m.add_argument("--doi-tuong", required=True)
    m.add_argument("--ai")
    o = sub.add_parser("mo", help="Mở lại hồ sơ đã lưu trữ")
    o.add_argument("ma")
    o.add_argument("--ai")
    s = sub.add_parser("phien", help="Tạo file phiên")
    s.add_argument("chu_de")
    s.add_argument("--ai", default="Claude Code")
    v = sub.add_parser("viec", help="Cấp ID và tạo file lỗi / việc hệ thống")
    v.add_argument("tieu_de")
    v.add_argument("--loi", action="store_true", help="Lỗi / rủi ro (E<n>); không có cờ = việc (V<n>)")
    v.add_argument("--muc", default="Trung bình")
    v.add_argument("--ai", default="Claude Code")
    sub.add_parser("so", help="Sổ mọi hồ sơ")
    for name in ("don", "kiem-tra", "dong"):
        c = sub.add_parser(name)
        if name == "dong":
            c.add_argument("ma", nargs="+")
        if name != "kiem-tra":
            c.add_argument("--ghi", action="store_true")
        c.add_argument("--ngay-ho-so", type=int, default=DOSSIER_IDLE_DAYS)
        c.add_argument("--khong-xet-file", action="store_true", help="Không xét giờ sửa file Input / Z")
    a = p.parse_args(argv)

    if a.cmd == "dau-phien":
        cmd_don(True, quiet=True)
        return cmd_xem()
    if a.cmd == "xem":
        return cmd_xem(a.day_du)
    if a.cmd == "khoa":
        return cmd_khoa(a.ten, a.ai, a.ghi_chu, a.ep)
    if a.cmd == "tra":
        return cmd_tra(a.ten, a.ai, a.ep)
    if a.cmd == "moi":
        return cmd_moi(a.ma, a.khach, a.doi_tuong, a.ai)
    if a.cmd == "mo":
        return cmd_mo(a.ma, a.ai)
    if a.cmd == "phien":
        return cmd_phien(a.chu_de, a.ai)
    if a.cmd == "viec":
        return cmd_viec(a.tieu_de, a.loi, a.muc, a.ai)
    if a.cmd == "so":
        return cmd_so()
    if a.cmd == "dong":
        return cmd_don(a.ghi, close=a.ma, idle_days=a.ngay_ho_so, use_files=not a.khong_xet_file)
    return cmd_don(getattr(a, "ghi", False), close=[], idle_days=a.ngay_ho_so, use_files=not a.khong_xet_file)


if __name__ == "__main__":
    sys.exit(main())
