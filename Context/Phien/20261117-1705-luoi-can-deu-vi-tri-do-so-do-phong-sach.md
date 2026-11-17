# Phiên 17/11/2026 17:05 — N2-PHAT-1647-059c — luoi can deu vi tri do so do phong sach

- **Tóm tắt:** Thêm `Core/Tools/even_grid.py` — căn đều vị trí đo trên sơ đồ phòng sạch theo lưới thẳng vách (10 → 3-4-3); AGENTS mục 8C + phiên bản 17/11/2026 17:20 (chưa chép bản Project).
- **Hồ sơ liên quan:** 120423 (áp đầu tiên)
- **File đổi:** Core/Tools/even_grid.py (mới), AGENTS.md, ContextLuuTru/AGENTS-LichSuPhienBan.md

## Nhật ký hệ thống (việc không thuộc riêng hồ sơ nào)

| Ngày | AI | Nội dung |
|---|---|---|
| 17/11/2026 | Claude Code | Người dùng: "sơ đồ đo phòng sạch tự động căn đều các điểm đo, và căn đều theo vách tường của khu vực phòng sạch, ví dụ như ô lưới 3 4 3 cho phòng sạch 10 điểm". Mới `Core/Tools/even_grid.py`: `row_counts` (ô gần vuông, ưu tiên đối xứng: 10→3-4-3, 5→2-1-2, 16→4×4), `grid_slots` (nội suy song tuyến trong tứ giác 4 góc — mặt bằng nghiêng / phòng hình thang), `snap_to_grid` (gán nhãn theo chỗ gần gốc nhất, tối ưu đổi cặp; ≤ 7 điểm duyệt hết), `rect_quad`. AGENTS mục 8C thêm quy tắc; bản AGENTS trong Project chưa chép (máy không truy cập Project). |
| 17/11/2026 | Claude Code | `even_grid.add_callouts` (ký hiệu đè chữ -> nhãn ra chỗ trống, ưu tiên trong phòng, mũi tên dẫn vào; người dùng: "nếu đè chữ thì thêm label ngoài hình và dẫn mũi tên vào khu vực"); `drawio_layout`: marker có `callout` -> nhãn tại đó + `_callout_arrow` gắn source; AGENTS 8C thêm vi sinh lưới riêng + hình thoi xanh + nhãn dẫn mũi tên. |
