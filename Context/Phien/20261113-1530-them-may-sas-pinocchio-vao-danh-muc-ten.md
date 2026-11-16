# Phiên 13/11/2026 15:30 — N2-PHAT-1505-c080 — them may sas_pinocchio vao danh muc ten may

- **Tóm tắt:** Thêm máy `sas_pinocchio` (Thiết bị lấy mẫu vi sinh khí nén SAS Pinocchio Super II, Italy) vào danh mục tên máy — hồ sơ 134179 là hồ sơ đầu tiên dùng (AGENTS mục 14: "dùng tới thì thêm vào EQUIPMENT trước").
- **Hồ sơ liên quan:** 134179
- **File đổi:** `Core/Tools/equipment_names.py` (thêm 1 mục cuối `EQUIPMENT`), `AGENTS.md` (mục 14: thêm dòng bảng, bỏ SAS khỏi câu "Chưa có trong danh mục"; phiên bản 13/11/2026 15:30), `ContextLuuTru/AGENTS-LichSuPhienBan.md` (dòng phiên bản 12/11 21:45)

## Nhật ký hệ thống (việc không thuộc riêng hồ sơ nào)

| Ngày | AI | Nội dung |
|---|---|---|
| 13/11/2026 | Claude Code | Thêm `sas_pinocchio` vào `EQUIPMENT` (regex `pinocchio`) + bảng AGENTS mục 14. Chưa chép AGENTS.md sang bản Project (phiên này không truy cập được Project) — người dùng chép giúp. Ghi nhận (chưa sửa Core): template TBKQ khí nén và `CompressedAirHandler` chưa có mục vi sinh khí nén — 134179 thêm bằng hậu xử lý trong `build_134179.py::_kn_tbkq_patch`; `TinhToanKhiNen_template.xlsx` còn số phiếu MTTN mẫu (J43 `26018422.01`) và số Lasair mẫu (C54:J54) — script hồ sơ phải xóa khi dùng AeroTrak. |
