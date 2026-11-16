# Phiên 14/11/2026 15:17 — N2-PHAT-1457-b4db — khuon TBKQ vi sinh khong khi be mat theo mau K Pack

- **Tóm tắt:** Người dùng đưa TBKQ - 1 Q326D124796 (K Pack, vi sinh không khí + swab) làm mẫu, yêu cầu dựng lại Hon Chuan (137058), Duy Tân (134179/2, 135843/2), rồi "Duyệt hết Hon chuan, Duy Tân để lại, nhưng cập nhật nguyên tắc chung cho các hồ sơ sau này" → thêm khuôn phụ lục TBKQ vi sinh vào AGENTS mục 8C.
- **Hồ sơ liên quan:** 137058, 134179, 135843
- **File đổi:** `AGENTS.md` (mục 8C: bullet mới "Vi sinh không khí / vi sinh bề mặt (swab) — khuôn phụ lục TBKQ…", thêm cuối dòng bảng vi sinh bề mặt; mục 3: 2 câu viện dẫn sang 8C ở "Ảnh hiện trường không chú thích" và "Câu công nhận bắt buộc"; phiên bản 14/11/2026 15:25), `ContextLuuTru/AGENTS-LichSuPhienBan.md` (dòng phiên bản 13/11 15:30)

## Nhật ký hệ thống (việc không thuộc riêng hồ sơ nào)

| Ngày | AI | Nội dung |
|---|---|---|
| 14/11/2026 | Claude Code | AGENTS mục 8C: khuôn phụ lục TBKQ vi sinh không khí / bề mặt theo mẫu K Pack — bìa mục 3 gạch đầu dòng, 9.1 ba dòng dẫn + bảng Vị trí / Ký hiệu vị trí / Nội dung thực hiện (thêm cột khu vực khi KH chia dây chuyền), ảnh theo mẫu (nhãn ký hiệu) hoặc theo nhóm (chú thích nhóm), 9.2 bảng chỉ ghi ký hiệu, Ghi chú "< 10" → VILAS → VIAS (VIAS đứng cuối — ngoại lệ mục 3). Đã duyệt: phần áp trên 137058. Chưa duyệt (ghi rõ trong quy tắc): chú thích nhóm ảnh; gộp hay tách mẫu bề mặt thiết bị / tay nhân viên. Core không đổi (`qa_check` không kiểm vị trí câu VIAS, 3 hồ sơ đều CLEAN). Khuôn hiện nằm trong 3 script hồ sơ (`location_table`, `photo_groups`, `unit_header`, `unshade` lặp ở `build_134179.py`, `build_135843.py`) — hồ sơ vi sinh kế tiếp nên đưa các hàm này vào Core thay vì chép lần nữa. Chưa chép AGENTS.md sang bản Project (phiên này không truy cập được Project) — người dùng chép giúp. |
