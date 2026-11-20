# Phiên 16/11/2026 18:36 — N2-PHAT-1832-6b8a — Ký hiệu (n) bảng Ghi chú in chỉ số trên

- **Tóm tắt:** Ký hiệu `(1)`…`(9)` đầu dòng bảng Ghi chú tự in chỉ số trên cho mọi TBKQ / BCKQ (template phòng sạch để `(1) …` chữ thường — người dùng báo 16/11/2026 trên 067460).
- **Hồ sơ liên quan:** 067460
- **File đổi:** `Core/Builders/tag_context.py` (mới `superscript_note_markers`, gọi trong `style_note_tables`), `Core/Builders/bckq_builder.py` (gọi sau `_tnr13`), `AGENTS.md` 16/11/2026 18:45 mục 3 (bản Project cần chép cùng nội dung).

## Nhật ký hệ thống (việc không thuộc riêng hồ sơ nào)

| Ngày | AI | Nội dung |
|---|---|---|
| 16/11/2026 | Claude Code | Hàm chung 	ag_context.superscript_note_markers(doc): bảng Ghi chú / Note (is_note_table), run chữ đầu tiên của mỗi đoạn bắt đầu (n) mà chưa superscript → tách (n) (vertAlign superscript, giữ cỡ / nghiêng) + phần còn lại. TBKQ: chạy trong style_note_tables (mọi handler + script hồ sơ gọi hàm này); BCKQ: BCKQBuilder.build sau bước Times 13. Kiểm trên 067460 /1, /2: TBKQ (1) sz 11 superscript +  …, BCKQ (1) sz 13 superscript; final-check CLEAN; soát PDF đúng. Hồ sơ phòng sạch khác chỉ đổi khi dựng lại. |
