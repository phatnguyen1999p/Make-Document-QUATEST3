"""Sổ file do AI tạo của từng hồ sơ + chép lên Z an toàn (AGENTS mục 0).

Người dùng chốt 09/11/2026: "các file được ghi là do Claude tạo thì được phép ghi đè",
kể cả khi quay lại bằng một phiên khác. Sổ ghi băm SHA-256 mọi bản AI đã tạo, nên phiên
sau biết bản trên Z còn là bản AI (ghi đè được) hay đã có người sửa sau đó (không ghi đè).

Sổ: Input\\<6 số>\\ThucHien\\_checkpoint\\ai_files.json  (thư mục ngoại lệ, không lên Z)
    {"code": "137035", "files": {"ThucHien/BCKQ - ….docx": [{"sha256", "size", "mtime", "ghi"}, …]}}

File AI (theo AGENTS mục 1): ThucHien/BCKQ - *.docx, ThucHien/*TinhToan*.xlsx,
TBKQ/TBKQ - *.docx, TBKQ/*tem tu*.docx, TBKQ/layout*.png, TBKQ/_diagrams/*,
TBKQ/Edoc/<6 số> - BBLM|Ghi Chep|Khac.pdf, pdf/*.pdf. File khác trong 3 thư mục
(vd BBLM đợt sau người dùng tự đặt, PDF người dùng tự xuất) không bao giờ bị chép/ghi đè.

Quy tắc khi chép lên Z (lệnh `chep`):
  - Z chưa có file                     → chép.
  - Z giống hệt Input                  → bỏ qua.
  - Băm bản Z có trong sổ (bản AI)     → ghi đè (phiên nào tạo cũng được).
  - Băm bản Z không có trong sổ        → có người sửa sau AI → KHÔNG ghi đè, in [Z EDITED]
                                         (chép bản Z về Input / hỏi người dùng — mục 0).
  Sau khi chép, băm bản mới được ghi vào sổ.

Lệnh:
  python Core/Tools/ai_files.py liet-ke <6 số>…          # file AI trong Input + tình trạng trên Z
  python Core/Tools/ai_files.py ghi <6 số>… [--kem-z]    # ghi bản Input (và bản Z nếu --kem-z) vào sổ
  python Core/Tools/ai_files.py chep <6 số>… [--ghi] [--z <thư mục hồ sơ trên Z>]
Không `--ghi` = chạy thử. `--kem-z` chỉ dùng khi người dùng xác nhận bản đang trên Z là bản AI.
"""
from __future__ import annotations

import argparse
import datetime as dt
import fnmatch
import hashlib
import json
import os
import re
import shutil
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
Z_ROOT = os.environ.get("LAMKQ_Z_ROOT", r"Z:\9.NV2 - MOI TRUONG\1.Ho So Moi Truong\2026")
Z_SUBDIRS = ("", "0.VIMCERTS078", "0.1.Tân Hiệp")
SUBS = ("ThucHien", "TBKQ", "pdf")
SKIP_DIRS = {"_checkpoint", "_optimized"}
AI_PATTERNS = (
    "thuchien/bckq - *.docx",
    "thuchien/*tinhtoan*.xlsx",
    "tbkq/tbkq - *.docx",
    "tbkq/*tem tu*.docx",
    "tbkq/layout*.png",
    "tbkq/_diagrams/*",
    "pdf/*.pdf",
)


def is_ai_file(rel: str, code: str) -> bool:
    r = rel.replace("\\", "/").lower()
    if os.path.basename(r).startswith("~$"):
        return False
    if re.fullmatch(rf"tbkq/edoc/{code} - (bblm|ghi chep|khac)\.pdf", r):
        return True
    return any(fnmatch.fnmatchcase(r, p) for p in AI_PATTERNS)


def sha256(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def input_dir(code: str) -> str:
    return os.path.join(ROOT, "Input", code)


def ledger_path(code: str) -> str:
    return os.path.join(input_dir(code), "ThucHien", "_checkpoint", "ai_files.json")


def load(code: str) -> dict:
    try:
        with open(ledger_path(code), encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {"code": code, "files": {}}


def save(code: str, led: dict) -> None:
    p = ledger_path(code)
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        json.dump(led, f, ensure_ascii=False, indent=1)


def known(led: dict, rel: str, digest: str) -> bool:
    return any(e["sha256"] == digest for e in led["files"].get(rel, []))


def record(led: dict, rel: str, path: str, digest: str | None = None) -> bool:
    digest = digest or sha256(path)
    if known(led, rel, digest):
        return False
    st = os.stat(path)
    led["files"].setdefault(rel, []).append({
        "sha256": digest, "size": st.st_size,
        "mtime": dt.datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
        "ghi": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    })
    return True


def ai_files(code: str) -> list[str]:
    """Đường dẫn tương đối (dấu /) của file AI trong Input\\<6 số>\\{ThucHien,TBKQ,pdf}."""
    base = input_dir(code)
    out = []
    for sub in SUBS:
        for dp, dns, fns in os.walk(os.path.join(base, sub)):
            dns[:] = [d for d in dns if d not in SKIP_DIRS]
            for fn in fns:
                rel = os.path.relpath(os.path.join(dp, fn), base).replace("\\", "/")
                if is_ai_file(rel, code):
                    out.append(rel)
    return sorted(out)


def z_folder(code: str, override: str | None = None) -> str | None:
    if override:
        return override
    found = []
    for sub in Z_SUBDIRS:
        b = os.path.join(Z_ROOT, sub) if sub else Z_ROOT
        try:
            found += [os.path.join(b, n) for n in os.listdir(b)
                      if code in n and os.path.isdir(os.path.join(b, n))]
        except OSError:
            continue
    # Bản sao kiểu "… (1)" là thư mục trùng, không phải thư mục hồ sơ.
    main = [p for p in found if not re.search(r" \(\d+\)$", p)]
    dup = [p for p in found if p not in main]
    if dup:
        print(f"[WARNING] [Z Duplicate] {code}: có thư mục trùng trên Z (không dùng): "
              + "; ".join(os.path.basename(p) for p in dup))
    if len(main) != 1:
        print(f"[CRITICAL WARNING] {code}: tìm thấy {len(main)} thư mục hồ sơ trên Z — "
              "dùng --z hoặc hỏi người dùng (AGENTS mục 0).")
        return None
    return main[0]


def status(code: str, led: dict, rel: str, z: str | None) -> str:
    src = os.path.join(input_dir(code), rel)
    if not z:
        return "?"
    dst = os.path.join(z, rel)
    if not os.path.exists(dst):
        return "Z chưa có"
    dz = sha256(dst)
    if dz == sha256(src):
        return "giống Z"
    if known(led, rel, dz):
        return "Z là bản AI cũ → ghi đè được"
    return "Z ĐÃ SỬA SAU AI → không ghi đè"


def cmd_list(code: str, z_override: str | None) -> None:
    led = load(code)
    z = z_folder(code, z_override)
    files = ai_files(code)
    print(f"== {code}: {len(files)} file AI — Z: {z or '-'}")
    for rel in files:
        n = len(led["files"].get(rel, []))
        print(f"  {rel:<70} sổ {n} bản | {status(code, led, rel, z)}")


def cmd_record(code: str, with_z: bool, z_override: str | None, write: bool) -> None:
    led = load(code)
    z = z_folder(code, z_override) if with_z else None
    added = 0
    for rel in ai_files(code):
        added += record(led, rel, os.path.join(input_dir(code), rel))
        if z and os.path.exists(os.path.join(z, rel)):
            added += record(led, rel, os.path.join(z, rel))
    print(f"{code}: thêm {added} bản vào sổ" + ("" if write else " (chạy thử)"))
    if write:
        save(code, led)


def cmd_copy(code: str, z_override: str | None, write: bool) -> int:
    led = load(code)
    z = z_folder(code, z_override)
    if not z:
        return 2
    blocked = copied = 0
    for rel in ai_files(code):
        src = os.path.join(input_dir(code), rel)
        dst = os.path.join(z, rel)
        ds = sha256(src)
        record(led, rel, src, ds)
        if os.path.exists(dst):
            dz = sha256(dst)
            if dz == ds:
                continue
            if not known(led, rel, dz):
                blocked += 1
                print(f"  [Z EDITED] {rel} — bản Z không phải bản AI nào trong sổ, không ghi đè")
                continue
        print(f"  {'chép' if write else 'sẽ chép'}: {rel}")
        if write:
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            try:
                shutil.copy2(src, dst)
                copied += 1
            except OSError as e:  # file đang mở trên Z (ERROR 32)…
                blocked += 1
                print(f"  [ERROR] {rel}: {e}")
    if write:
        save(code, led)
    print(f"{code}: {'đã chép' if write else 'chạy thử'} {copied}, chặn {blocked} — Z: {z}")
    return 1 if blocked else 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("lenh", choices=["liet-ke", "ghi", "chep"])
    ap.add_argument("codes", nargs="+", help="6 số hồ sơ")
    ap.add_argument("--ghi", action="store_true", help="ghi thật (mặc định chạy thử) — cho lệnh chep")
    ap.add_argument("--kem-z", action="store_true", help="lệnh ghi: ghi cả bản đang trên Z vào sổ")
    ap.add_argument("--z", help="thư mục hồ sơ trên Z (khi tìm thấy nhiều)")
    a = ap.parse_args(argv)
    rc = 0
    for code in a.codes:
        if a.lenh == "liet-ke":
            cmd_list(code, a.z)
        elif a.lenh == "ghi":
            cmd_record(code, a.kem_z, a.z, True)
        else:
            rc = max(rc, cmd_copy(code, a.z, a.ghi))
    return rc


if __name__ == "__main__":
    sys.exit(main())
