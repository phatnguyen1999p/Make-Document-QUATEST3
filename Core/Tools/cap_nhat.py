"""Kiểm tra phiên bản và tải bản mới nhất đầu phiên (người dùng yêu cầu 20/11/2026, AGENTS đầu file + mục 12).

Nhiều máy cùng dùng một repo git → máy nào cũng phải chạy code / quy tắc mới nhất trước khi làm hồ sơ.
`context.py dau-phien` tự gọi `kiem_tra()` trước khi in tóm tắt; chạy riêng: `python Core/Tools/context.py cap-nhat`
hoặc `python Core/Tools/cap_nhat.py`.

Làm gì:
1. `git fetch origin` (hết 40 giây hoặc mất mạng → báo, không chặn phiên).
2. Nhánh hiện tại có upstream và đang **sau** upstream:
   - không có commit riêng chưa đẩy, không có file đã theo dõi đang sửa → `git merge --ff-only` (tải về, an toàn);
   - ngược lại → không tải, in `[CẢNH BÁO] [Phiên bản]` kèm lý do → báo người dùng.
3. Nhánh khác `main` mà thiếu commit của `origin/main` → cảnh báo (không tự chuyển nhánh / merge).
4. So dòng "Phiên bản dd/mm/yyyy HH:MM" của AGENTS.md trên máy với `origin/main`.
Không bao giờ `reset`, `stash`, `checkout` hay ghi đè file người dùng đang sửa.
"""
from __future__ import annotations

import os
import re
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
MAIN = "origin/main"
_VER = re.compile(r"Phiên bản (\d{2})/(\d{2})/(\d{4}) (\d{2}):(\d{2})")


def _git(*args, timeout: int = 30):
    try:
        r = subprocess.run(["git", *args], cwd=ROOT, capture_output=True, text=True, encoding="utf-8",
                           errors="replace", timeout=timeout)
        return r.returncode, (r.stdout or "").rstrip(), (r.stderr or "").strip()
    except (OSError, subprocess.TimeoutExpired) as e:
        return 1, "", str(e)


def _agents_version(text: str):
    m = _VER.search(text or "")
    if not m:
        return None, ""
    d, mo, y, h, mi = m.groups()
    return (y, mo, d, h, mi), f"{d}/{mo}/{y} {h}:{mi}"


def kiem_tra() -> bool:
    """In tình trạng phiên bản; trả True khi vừa tải code mới về (người gọi nên chạy lại bằng code mới)."""
    if _git("rev-parse", "--is-inside-work-tree")[0] != 0:
        print("[Phiên bản] Không phải repo git — bỏ qua kiểm tra bản mới.")
        return False
    code, _, err = _git("fetch", "--quiet", "origin", timeout=40)
    if code != 0:
        print(f"[CẢNH BÁO] [Phiên bản] Không tải được thông tin từ origin ({err[:120] or 'lỗi mạng'}) — "
              "chưa biết có bản mới không; báo người dùng nếu sắp làm hồ sơ.")
        return False

    branch = _git("rev-parse", "--abbrev-ref", "HEAD")[1]
    dirty = [l for l in _git("status", "--porcelain", "--untracked-files=no")[1].splitlines() if l.strip()]
    upstream = _git("rev-parse", "--abbrev-ref", "--symbolic-full-name", "@{u}")
    up = upstream[1] if upstream[0] == 0 else (MAIN if branch == "main" else "")
    pulled = False
    if up:
        cnt = _git("rev-list", "--left-right", "--count", f"HEAD...{up}")[1].split()
        ahead, behind = (int(cnt[0]), int(cnt[1])) if len(cnt) == 2 else (0, 0)
        if behind and not ahead and not dirty:
            code, _, err = _git("merge", "--ff-only", "--quiet", up, timeout=60)
            if code == 0:
                pulled = True
                print(f"[Phiên bản] Đã tải bản mới nhất: {behind} commit từ {up} vào nhánh {branch}.")
            else:
                print(f"[CẢNH BÁO] [Phiên bản] Tải {behind} commit từ {up} thất bại ({err[:160]}) — báo người dùng.")
        elif behind:
            why = []
            if ahead:
                why.append(f"nhánh có {ahead} commit chưa đẩy")
            if dirty:
                why.append(f"{len(dirty)} file đã theo dõi đang sửa ({', '.join(l[3:] for l in dirty[:3])}…)")
            print(f"[CẢNH BÁO] [Phiên bản] Nhánh {branch} sau {up} {behind} commit nhưng CHƯA tải: "
                  f"{'; '.join(why)}. Báo người dùng trước khi làm hồ sơ (commit / đẩy rồi `git pull`).")
        else:
            print(f"[Phiên bản] Nhánh {branch} đã là bản mới nhất của {up}"
                  + (f" ({ahead} commit chưa đẩy)." if ahead else "."))
    else:
        print(f"[Phiên bản] Nhánh {branch} chưa có upstream — không tự tải.")

    if branch != "main" and _git("rev-parse", "--verify", "--quiet", MAIN)[0] == 0:
        miss = _git("rev-list", "--count", "--no-merges", f"HEAD..{MAIN}")[1]
        if miss.isdigit() and int(miss):
            print(f"[CẢNH BÁO] [Phiên bản] Nhánh {branch} thiếu {miss} commit của main — hỏi người dùng có "
                  "chuyển về main / merge main trước khi làm hồ sơ không (không tự chuyển nhánh).")

    try:
        with open(os.path.join(ROOT, "AGENTS.md"), encoding="utf-8") as f:
            local = _agents_version(f.read(4000))
    except OSError:
        local = (None, "")
    remote = _agents_version(_git("show", f"{MAIN}:AGENTS.md")[1][:4000])
    if local[0] and remote[0]:
        if remote[0] > local[0]:
            print(f"[CẢNH BÁO] [Phiên bản] AGENTS.md trên máy {local[1]} cũ hơn main {remote[1]} — đọc quy tắc "
                  "mới ở main trước khi làm.")
        else:
            print(f"[Phiên bản] AGENTS.md {local[1]} (main {remote[1]}).")
    return pulled


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    kiem_tra()
