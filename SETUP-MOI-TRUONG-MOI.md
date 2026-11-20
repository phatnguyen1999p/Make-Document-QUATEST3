# Prompt setup môi trường Claude mới — Làm KQ

> Soạn ngày 11/11/2026. Giải nén thư mục `Làm KQ` vào `C:\Users\<user>\Desktop\Source\Làm KQ` rồi dán khối dưới vào phiên Claude mới.
> Thư mục `CongCu\` đã chứa sẵn UpChungThu (chỉ mã nguồn + giấy tiêu đề) và DrawIO AutoAddImage.
> Memory của Claude không nằm trong zip; nội dung đã chép vào mục 7 của prompt. Nếu máy mới vẫn dùng user `phat.nt` và đã chép nguyên thư mục memory thì bỏ mục 7.

```text
Mình vừa giải nén dự án "Làm KQ" (bộ công cụ dựng hồ sơ giám định QUATEST 3: BCKQ, TBKQ VN/EN, phiếu tính toán Excel, PDF) sang máy/môi trường mới. Hãy setup và kiểm tra môi trường, CHƯA làm hồ sơ nào.

Quy tắc trong lúc setup: không sửa code trong Core/, không xóa file, không ghi gì lên ổ Z:, không chạy build_<6 số>.py thật (chỉ chạy thử ở %TEMP%).

1. Đọc trước: AGENTS.md (quy tắc bất biến, ưu tiên cao nhất), README.md; chạy python Core/Tools/context.py xem (tóm tắt hồ sơ / lỗi đang mở, không ghi gì). KHÔNG đọc file lưu trữ. Context nằm trên Z (thư mục ẩn Claude - Work, AGENTS mục 16) — phải map ổ Z trước.
2. Vị trí thư mục: kiểm tra dự án có nằm ở C:\Users\<user>\Desktop\Source\Làm KQ không. Nếu tên user khác "phat.nt", tìm các chỗ còn cài cứng "phat.nt" trong code đang dùng (bỏ qua Core/Source Code/ — script cũ) và báo lại; không sửa khi chưa hỏi.
3. Python ≥ 3.9: chạy `pip install -e .` theo pyproject.toml (python-docx, openpyxl, PyMuPDF, lxml, matplotlib, pillow, pywin32, xlrd, PyPDF2, markitdown). Thử import từng gói (docx, openpyxl, fitz, lxml, win32com.client, PIL, matplotlib, xlrd, PyPDF2). pywin32 lỗi thì chạy postinstall. Cài thêm requirements.txt của CongCu\AutoAddImage nếu có gói thiếu.
4. Phần mềm ngoài, mỗi thứ báo có/thiếu:
   - Microsoft Word + Excel (Word/Excel COM dùng để xuất PDF): mở thử qua win32com rồi đóng.
   - Máy in "Microsoft Print to PDF" (code tránh máy in tem HPL/HPRT/ZPL).
   - draw.io desktop (CLI dùng để xuất sơ đồ).
   - Windows OCR (Windows.Media.Ocr, dùng cho Core/Tools/pdf_orient.py và scan_splitter): kiểm tra có gói ngôn ngữ OCR tiếng Việt/Anh.
   - Công cụ chép kèm trong workspace: CongCu\UpChungThu (phải có "GIAY TIEU DE.pdf", "GIAY TIEU DE - E.pdf", core\, app_settings.json) và CongCu\AutoAddImage (drawio_builder.py, drawio_exporter.py). Code tự tìm: biến môi trường UPCHUNGTHU_PATH / DRAWIO_TOOL_DIR → bản gốc Desktop\Source\Source Code\… → bản CongCu. Thử nạp: `from Core.Services import upchungthu_bridge as u; u._init_upchungthu()` và `from Core.Tools import drawio_layout as d; d._tool()`, in ra thư mục đang dùng.
5. Ổ mạng: Test-Path "Z:\9.NV2 - MOI TRUONG\1.Ho So Moi Truong\2026". Chỉ đọc, không ghi. (app_settings.json của UpChungThu cũng trỏ vào Z:.)
6. Chạy thử (không đụng Input\ thật, không đụng Z):
   - python Core/run.py --status "Input\<một hồ sơ có sẵn>"
   - chay_thu.bat (chạy trong %TEMP%\lamkq_chay_thu) rồi đọc log.
   - python Core\Tools\qa_check.py final-check trên một TBKQ có sẵn trong Input\<hồ sơ>\TBKQ\.
   - python Core/Tools/context.py kiem-tra (chạy thử, không ghi gì).
   - python Core/Tools/equipment_names.py
7. Tạo lại memory của project mới (mỗi ý 1 file + dòng trong MEMORY.md, đúng định dạng memory):
   - z-drive-sync (feedback): từ 30/10/2026, xong hồ sơ (dựng + QA xong) thì chép ThucHien/TBKQ/pdf lên thư mục hồ sơ có sẵn trên Z:\9.NV2 - MOI TRUONG\1.Ho So Moi Truong\2026\ (dạng "Q326D<6 số> - <khách> - <loại>"). Từ 09/11/2026 chép bằng `python Core/Tools/ai_files.py chep <6 số> --ghi` (chạy thử khi bỏ --ghi). Dự phòng: robocopy /E /XD _checkpoint /XF ~$*, không /MIR, không /PURGE, không xóa gì trên Z. Không tìm thấy thư mục hoặc thấy nhiều thư mục khớp → hỏi. Không gõ cứng đường dẫn Z từ phiên cũ.
   - ai-files-ledger-overwrite (feedback, 09/11/2026): file Claude tạo, đã ghi trong sổ Input\<6 số>\ThucHien\_checkpoint\ai_files.json, được ghi đè trên Z kể cả từ phiên khác. Bản trên Z không khớp sổ → có người sửa tay → chép về hoặc hỏi. Không bao giờ tạo thư mục trùng kiểu "… (1)".
   - cert-date-default-today (feedback, 04 và 06/11/2026): hồ sơ không có KQTN → ngày ban hành = ngày hiện tại (CERT_DATE = date.today()), tự cập nhật cả khi dựng lại, không hỏi lại. Có KQTN → ngày KQTN muộn nhất.
   - batch-questions-once (feedback, 04/11/2026): gom mọi điểm cần người dùng quyết của một hồ sơ, hỏi một lần; bỏ những điểm đã có quy tắc mặc định.
   - nextern-recurring-dossier (project): Công ty TNHH Nextern Việt Nam (KCN Mỹ Phước 3) làm hồ sơ hằng tháng (108033, 122266, 133266, 137035…). /1 phòng sạch: Phòng sản xuất 139,6 m² 21 vị trí, Air lock 4,08 m² 2 vị trí, Phòng thay đồ 27,6 m² 8 vị trí, ISO 8, Lasair III 110. /2 khí nén KN1/KN2. Ngưỡng nội bộ ghi tay trên phiếu: dầu < 1 mg/m³; nước < 1 200 mg/m³; hạt d > 0,5 µm < 27 615 hạt/m³ (lấy mẫu 5 phút × 28,3 L/phút). Hồ sơ mới thì copy build_133266.py rồi đổi số liệu.
8. Báo cáo cuối: bảng hạng mục Đạt / Thiếu / Cần mình làm (vd cài draw.io, map ổ Z, cài gói OCR), kèm lệnh đã chạy và lỗi nếu có.
```
