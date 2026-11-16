"""Hook PreToolUse (Bash | PowerShell) của plugin lamkq — chặn lệnh trái AGENTS.md mục 0 trên ổ Z.

Chặn (exit 2, lý do in ra stderr để Claude đọc):
- robocopy có /MIR hoặc /PURGE (xóa file trên đích);
- lệnh xóa / di chuyển / đổi tên nhắm vào ổ Z (Remove-Item, del, rd, rmdir, rm, Move-Item, mv, ren, Rename-Item).
Mọi lệnh khác cho qua (exit 0). Lỗi đọc dữ liệu hook → cho qua, không làm hỏng phiên.
"""
import json
import re
import sys

Z_DRIVE = re.compile(r"(?i)(?<![\w$])(?:z:[\\/]|/z/)")
# Đường dẫn UNC (\\server\share\...) của cùng thư mục hồ sơ trên Z.
Z_SHARE_HINT = re.compile(r"(?i)9\.NV2 - MOI TRUONG|Ho So Moi Truong")
DELETE_VERBS = re.compile(
    r"""(?ix) (?:^|[\s;|&(]) (?: remove-item | rm | rmdir | rd | del | erase | ri
                               | move-item | mv | move | ren | rename | rename-item | clear-content )
        (?=\s) """
)
ROBO_MIR = re.compile(r"(?i)\brobocopy\b[^\n;|&]*\s/(?:MIR|PURGE)\b")


def _segments(cmd: str):
    return [s for s in re.split(r"[;\n]|&&|\|\|", cmd) if s.strip()]


def check(cmd: str):
    if ROBO_MIR.search(cmd):
        return "robocopy /MIR hoặc /PURGE bị cấm (AGENTS.md mục 0: không xóa gì trên Z). Dùng `python Core/Tools/ai_files.py chep <6 số> --ghi` hoặc robocopy /E."
    for seg in _segments(cmd):
        if not DELETE_VERBS.search(seg):
            continue
        if Z_DRIVE.search(seg) or Z_SHARE_HINT.search(seg):
            return ("Lệnh xóa / di chuyển / đổi tên trên ổ Z bị cấm (AGENTS.md mục 0: trên Z không xóa, "
                    "không đổi tên, không di chuyển). Nếu thật sự cần, báo người dùng tự làm.")
    return None


def main() -> int:
    try:
        data = json.load(sys.stdin)
        cmd = (data.get("tool_input") or {}).get("command") or ""
    except Exception:
        return 0
    reason = check(cmd)
    if reason:
        try:
            sys.stderr.reconfigure(encoding="utf-8")
        except Exception:
            pass
        print(f"[lamkq] Đã chặn: {reason}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
