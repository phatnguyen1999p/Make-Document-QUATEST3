r"""Hook UserPromptSubmit của plugin lamkq — kiểm tra cập nhật định kỳ trong phiên.

Người dùng yêu cầu 20/11/2026: "đầu mỗi ngày hoặc mỗi giờ kiểm tra lại cập nhật". Mỗi lần người dùng nhắn,
gọi `python Core/Tools/cap_nhat.py --neu-qua 60`: chỉ fetch / tải khi lần kiểm tra trước đã quá 60 phút
(`LAMKQ_CAP_NHAT_PHUT` để đổi), chưa tới hạn thì không in gì. Kết quả in ra stdout → vào ngữ cảnh phiên.
Chỉ chạy trong dự án "Làm KQ"; lỗi không chặn tin nhắn: luôn thoát 0.
"""
import json
import os
import subprocess
import sys


def main() -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    root = os.environ.get("CLAUDE_PROJECT_DIR")
    if not root:
        try:
            root = json.load(sys.stdin).get("cwd")
        except Exception:
            root = None
    root = os.path.normpath(root or os.getcwd())
    tool = os.path.join(root, "Core", "Tools", "cap_nhat.py")
    if not os.path.isfile(tool):
        return 0
    py = os.environ.get("LAMKQ_PYTHON") or sys.executable
    try:
        r = subprocess.run([py, tool, "--neu-qua", "60"], cwd=root, capture_output=True, text=True,
                           encoding="utf-8", errors="replace", timeout=110,
                           env=dict(os.environ, PYTHONIOENCODING="utf-8"))
        if (r.stdout or "").strip():
            print(r.stdout.strip())
    except Exception:
        pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
