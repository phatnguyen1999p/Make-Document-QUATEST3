"""Lệnh cũ — chuyển tiếp sang `Core/Tools/context.py` (AGENTS mục 16).

Từ 12/11/2026 context tách thành file riêng theo hồ sơ / phiên / lỗi trong `Context/`; PROJECT_CONTEXT.md
không còn mục 1–4, 5E để dọn. Các lệnh cũ vẫn gõ được:

    python Core/Tools/context_archive.py kiem-tra        -> context.py kiem-tra
    python Core/Tools/context_archive.py don --ghi       -> context.py don --ghi
    python Core/Tools/context_archive.py dong 132809 --ghi -> context.py dong 132809 --ghi
    python Core/Tools/context_archive.py mo 132809       -> context.py mo 132809

Bản trước khi tách (cắt dòng khỏi PROJECT_CONTEXT.md) không còn dùng; PROJECT_CONTEXT.md gốc lưu ở
`ContextLuuTru/PROJECT_CONTEXT_truoc_khi_tach_20261112.md`.
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import context  # noqa: E402


def main(argv=None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    out, i = [], 0
    while i < len(args):
        if args[i] == "--ngay":                      # tùy chọn cũ: số ngày giữ nhật ký hệ thống
            i += 2
            continue
        if args[i] == "--giu":                       # tùy chọn cũ: ID lỗi giữ lại → nay ghi [giữ] trong file lỗi
            i += 1
            while i < len(args) and not args[i].startswith("--"):
                i += 1
            continue
        out.append(args[i])
        i += 1
    return context.main(out)


if __name__ == "__main__":
    sys.exit(main())
