# Phiên 11/11/2026 — nhật ký hệ thống chuyển từ PROJECT_CONTEXT.md mục 4 (nguyên văn)

- **Tóm tắt:** 4 dòng nhật ký không thuộc riêng một hồ sơ đang mở (tách 12/11/2026).

## Nhật ký hệ thống (việc không thuộc riêng hồ sơ nào)

| Ngày | AI | Nội dung |
|---|---|---|
| 11/11/2026 | Claude Code | Hệ thống (chuyển máy): chép kèm công cụ ngoài vào workspace `CongCu\UpChungThu` (chỉ mã nguồn + giấy tiêu đề + `app_settings.json`, bỏ build/dist/Test Data, 1 MB) và `CongCu\AutoAddImage`; `upchungthu_bridge`, `drawio_layout` dùng bản gốc `Desktop\Source\Source Code\…` khi có, không có thì dùng bản `CongCu` (đã thử cả hai). Bản `CongCu` là ảnh chụp 11/11 — sửa công cụ gốc thì chép lại. |
| 11/11/2026 | Claude Code | Hệ thống (Edoc): phân loại trang theo tiêu đề đứng trước + từ khóa nguyên từ (169/173 trang Edoc đúng, trước 163); `finish_edoc` (bỏ trắng + xoay) cho 9 script; bỏ trang trắng Edoc 137035 (3 tr.), 132115 (1 tr.), đã chép đè Z. |
| 11/11/2026 | Claude Code | Hệ thống (Edoc): tách PDF từ scan tự xoay trang về đúng chiều đọc — `Core/Tools/pdf_orient.py`, nối vào `scan_splitter` (trước OCR header, khi ghi, `write_edoc`) và `split_edoc` 9 script hồ sơ; AGENTS mục 8A. Edoc hiện có: 9 trang ngang (120423, 132115, 132809, 136839) — người dùng cho xoay, đã chép đè 4 file Edoc lên Z (từng file, bản Z trùng băm bản AI) + ghi sổ `ai_files`. |
| 11/11/2026 | Claude Code | Hệ thống (khí nén 133266, 137035; phòng sạch): rà chế độ số liệu máy đếm hạt theo yêu cầu người dùng (khí nén Δ / Diff, phòng sạch Σ / Cumul) — mọi hồ sơ đúng. Thêm `Core/Tools/particle_mode.py`, kiểm `Count Mode` trong `read_aerotrak`, `dust["mode"]` trong `khinen_calc`; AGENTS mục 8C. |
