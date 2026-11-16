r"""Hook SessionStart của plugin lamkq.

- Chỉ làm việc khi phiên mở trong dự án "Làm KQ" (có Core/Tools/context.py); dự án khác → im lặng.
- Đặt UPCHUNGTHU_PATH, DRAWIO_TOOL_DIR trỏ vào bản trong CongCu\ (người dùng chốt 16/11/2026)
  cho các lệnh Claude chạy trong phiên (qua CLAUDE_ENV_FILE).
- Chạy `python Core/Tools/context.py dau-phien` (AGENTS.md: bắt buộc đầu phiên) và in kết quả ra stdout
  → Claude Code đưa vào ngữ cảnh phiên.
Lỗi ở đây không được chặn phiên: luôn thoát 0.
"""
import json
import os
import subprocess
import sys


def _project_dir() -> str:
    d = os.environ.get("CLAUDE_PROJECT_DIR")
    if not d:
        try:
            d = json.load(sys.stdin).get("cwd")
        except Exception:
            d = None
    return os.path.normpath(d or os.getcwd())


def _write_env(root: str) -> list:
    tools = {
        "UPCHUNGTHU_PATH": os.path.join(root, "CongCu", "UpChungThu"),
        "DRAWIO_TOOL_DIR": os.path.join(root, "CongCu", "AutoAddImage"),
    }
    found = {k: v for k, v in tools.items() if os.path.isdir(v)}
    env_file = os.environ.get("CLAUDE_ENV_FILE")
    if env_file and found:
        try:
            with open(env_file, "a", encoding="utf-8") as f:
                for k, v in found.items():
                    f.write(f"export {k}='{v}'\n")
        except OSError:
            pass
    return [f"{k}={v}" for k, v in found.items()] + [
        f"[THIẾU] {k}: không có {v}" for k, v in tools.items() if k not in found
    ]


def main() -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    root = _project_dir()
    ctx = os.path.join(root, "Core", "Tools", "context.py")
    if not os.path.isfile(ctx):
        return 0  # không phải dự án Làm KQ

    lines = ["[lamkq] Công cụ dùng bản trong CongCu:"] + ["  " + s for s in _write_env(root)]

    py = os.environ.get("LAMKQ_PYTHON") or sys.executable
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    try:
        r = subprocess.run([py, ctx, "dau-phien"], cwd=root, env=env, capture_output=True,
                           text=True, encoding="utf-8", errors="replace", timeout=170)
        out = (r.stdout or "") + (("\n" + r.stderr) if r.returncode and r.stderr else "")
        lines += ["", "[lamkq] Kết quả `python Core/Tools/context.py dau-phien` (đã chạy tự động — "
                      "không cần chạy lại; dùng nhãn phiên trong đó cho --ai):", out.strip()]
    except Exception as e:  # noqa: BLE001
        lines += ["", f"[lamkq] Không chạy được context.py dau-phien ({e}) — tự chạy lệnh này trước khi làm việc."]

    print("\n".join(lines))
    return 0


if __name__ == "__main__":
    sys.exit(main())
