# Phiên 16/11/2026 08:12 — N2-PHAT-0812-c2d5 — plugin lamkq cho nhieu PC

- **Tóm tắt:** Dựng khung plugin Claude Code `lamkq` (marketplace cục bộ `CongCu\lamkq-marketplace`) để triển khai quy trình Làm KQ lên nhiều PC: hook SessionStart (đặt UPCHUNGTHU_PATH / DRAWIO_TOOL_DIR trỏ `CongCu\`, chạy `context.py dau-phien`), hook PreToolUse chặn robocopy /MIR /PURGE và xóa / di chuyển trên Z, 3 skill `lam-ho-so`, `chep-z`, `setup-may-moi`. Người dùng chốt 16/11/2026: "Dùng bản trong /CongCu" (UpChungThu, AutoAddImage). Chưa cài plugin lên máy nào; chưa sửa Core, AGENTS.
- **Hồ sơ liên quan:** không
- **File đổi:** mới `CongCu/lamkq-marketplace/` (`.claude-plugin/marketplace.json`, `README.md`, `lamkq/.claude-plugin/plugin.json`, `lamkq/hooks/hooks.json`, `lamkq/scripts/dau_phien.py`, `lamkq/scripts/chan_lenh_z.py`, `lamkq/skills/{lam-ho-so,chep-z,setup-may-moi}/SKILL.md`, `lamkq/skills/lam-ho-so/references/khach-quen.md`)

## Nhật ký hệ thống (việc không thuộc riêng hồ sơ nào)

| Ngày | AI | Nội dung |
|---|---|---|
| 16/11/2026 | Claude Code | Người dùng đã đẩy source lên GitHub `phatnguyen1999p/Make-Document-QUATEST3` (Private — API không đăng nhập trả 404; commit `2ddff31`). Thêm `.claude-plugin/marketplace.json` ở gốc repo (plugin `./CongCu/lamkq-marketplace/lamkq`) để cài bằng `claude plugin marketplace add phatnguyen1999p/Make-Document-QUATEST3`; thêm mục cài plugin vào `README.md` gốc và `CongCu/lamkq-marketplace/README.md`; `plugin validate .` passed. Chưa commit / push thay đổi này. |
| 16/11/2026 | Claude Code | Người dùng "Cài thử máy này": `claude plugin marketplace add` + `install lamkq@lamkq-marketplace` (scope user, enabled; 3 skill, 2 hook, ~184 tok luôn bật). Chưa thử trong phiên mới. Người dùng hỏi đưa source lên GitHub: tạo `.gitignore` (loại `Input/`, `HieuChuan/` 969 MB, `MTLD/` 1 GB, dist, cache, log, `Context/_khoa/`) — phần còn lại ~443 file / 182 MB, không file > 20 MB; chưa `git init` (người dùng tự chạy). Liên quan V9. |
| 16/11/2026 | Claude Code | Dựng khung plugin lamkq; `claude plugin validate` marketplace + plugin: passed; thử `chan_lenh_z.py` 8 ca (MIR/PURGE, Remove-Item Z, rm /z bị chặn; robocopy /E, ai_files, đọc Z, xóa file cục bộ cho qua); thử `dau_phien.py`: ghi đúng CLAUDE_ENV_FILE, in tóm tắt dau-phien (~13 KB), dự án khác im lặng; Core nhận UpChungThu / AutoAddImage từ `CongCu\` khi có biến. Còn mở (việc V9): khóa `Context/_khoa` và sổ `ai_files.json` chưa dùng chung giữa các máy. |
