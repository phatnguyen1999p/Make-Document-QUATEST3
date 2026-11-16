# Phiên 15/11/2026 10:17 — N2-PHAT-0952-0d5e — Khí nén EN số ít location/particle

- **Tóm tắt:** TBKQ khí nén EN in "01 locations", "1 particles/m3" -> sửa `build_context` của handler khí nén: số ít khi n = 1.
- **Hồ sơ liên quan:** 096547
- **File đổi:** `Core/Builders/handlers/compressed_air.py` (2 dòng); `Core/Tools/equipment_names.py` (`oil_kit`); `AGENTS.md` mục 14 + phiên bản 15/11/2026 10:30

## Nhật ký hệ thống (việc không thuộc riêng hồ sơ nào)

| Ngày | AI | Nội dung |
|---|---|---|
| 15/11/2026 | Claude Code | `ctx["so_vi_tri"]` EN: "location" khi 1 vị trí; mô tả hạt d > 5 µm EN: "particle" khi = 1. Thử trên 096547 /1, /2 (final-check CLEAN). |
| 15/11/2026 | Claude Code | Người dùng chốt: lấy mẫu tổng dầu khí nén ghi "Giấy lọc thủy tinh Pall life Sciences, ống than hoạt tính ORBO®" (EN "Glass fiber filter Pall life Sciences, carbon activated tube ORBO®") — sửa `EQUIPMENT["oil_kit"]` (regex nhận cả tên cũ để `normalize_equipment` tự quy đổi), AGENTS mục 14 (bảng + câu "Dụng cụ …"), lịch sử phiên bản. Chưa chép AGENTS lên bản Project. Hồ sơ cũ dùng `oil_kit` (133266, 134179, 135843) giữ nguyên tới khi dựng lại. |
