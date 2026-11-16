---
name: setup-may-moi
description: Cài đặt và kiểm tra môi trường "Làm KQ" trên một PC mới (Python, Word/Excel COM, draw.io, Windows OCR, ổ Z, công cụ trong CongCu). Dùng khi người dùng nói "setup máy mới", "cài Làm KQ cho máy này", "kiểm tra môi trường".
---

# Setup PC mới cho "Làm KQ"

Trong lúc setup: **không** sửa code trong `Core/`, không xóa file, không ghi gì lên Z, không chạy `build_<6 số>.py` thật (chỉ chạy thử trong `%TEMP%`). Hướng dẫn đầy đủ hơn: `SETUP-MOI-TRUONG-MOI.md` ở gốc dự án.

## 1. Vị trí và đọc trước

- Dự án nên ở `C:\Users\<user>\Desktop\Source\Làm KQ`. Đọc `AGENTS.md`, `README.md`; chạy `python Core/Tools/context.py xem` (chỉ in).

## 2. Python

- Python ≥ 3.9: `pip install -e .` ở gốc dự án (theo `pyproject.toml`).
- Thử import: `docx, openpyxl, fitz, lxml, win32com.client, PIL, matplotlib, xlrd, PyPDF2`. pywin32 lỗi → chạy postinstall của pywin32.
- `pip install -r CongCu\AutoAddImage\requirements.txt` nếu thiếu gói.
- Python không nằm trên PATH → đặt biến `LAMKQ_PYTHON` trỏ vào `python.exe` (hook của plugin và các `chay_*.ps1` cần `python`).

## 3. Công cụ — dùng bản trong `CongCu\` (người dùng chốt 16/11/2026)

Mọi máy dùng bản UpChungThu và AutoAddImage chép kèm trong `CongCu\`, không dùng bản `Desktop\Source\Source Code\…` riêng của từng máy.

- Hook đầu phiên của plugin tự đặt `UPCHUNGTHU_PATH` và `DRAWIO_TOOL_DIR` trỏ vào `CongCu\` cho các lệnh Claude chạy trong phiên.
- Để `chay_*.bat` / chạy tay ngoài Claude cũng dùng `CongCu\`, **hỏi người dùng rồi mới** đặt biến môi trường cố định cho user (đổi cấu hình máy):
  ```
  setx UPCHUNGTHU_PATH "<gốc dự án>\CongCu\UpChungThu"
  setx DRAWIO_TOOL_DIR "<gốc dự án>\CongCu\AutoAddImage"
  ```
- Kiểm tra thư mục đang dùng:
  ```
  python -c "from Core.Services import upchungthu_bridge as u; print(u.get_upchungthu_dir()); u._init_upchungthu()"
  python -c "from Core.Tools import drawio_layout as d; print(d.TOOL_DIR); d._tool()"
  ```
- `CongCu\UpChungThu` phải có `GIAY TIEU DE.pdf`, `GIAY TIEU DE - E.pdf`, `core\`, `app_settings.json`.

## 4. Phần mềm ngoài — báo có / thiếu từng thứ

- Microsoft Word + Excel (mở thử qua `win32com` rồi đóng).
- Máy in "Microsoft Print to PDF" (code tránh máy in tem HPL/HPRT/ZPL).
- draw.io desktop (CLI xuất sơ đồ).
- Windows OCR có gói ngôn ngữ tiếng Việt / tiếng Anh (`pdf_orient.py`, `scan_splitter`).

## 5. Ổ mạng Z (chỉ đọc)

`Test-Path "Z:\9.NV2 - MOI TRUONG\1.Ho So Moi Truong\2026"`.

## 6. Chạy thử (không đụng `Input\` thật, không đụng Z)

- `python Core/run.py --status "Input\<một hồ sơ có sẵn>"`
- `chay_thu.bat` (chạy trong `%TEMP%\lamkq_chay_thu`) rồi đọc log
- `python Core\Tools\qa_check.py final-check` trên một TBKQ có sẵn
- `python Core/Tools/context.py kiem-tra`
- `python Core/Tools/equipment_names.py`

## 7. Lưu ý khi nhiều máy cùng làm

Khóa hồ sơ (`Context/_khoa/`) và sổ `ai_files.json` hiện nằm trên từng máy → hai máy không thấy khóa / sổ của nhau. Cho đến khi xong việc V9 (`Context/Viec/V9.md`): **mỗi hồ sơ chỉ làm trên một máy**, người dùng tự phân công.

## 8. Báo cáo cuối

Bảng hạng mục Đạt / Thiếu / Cần người dùng làm (cài draw.io, map ổ Z, cài gói OCR, `setx`…), kèm lệnh đã chạy và lỗi nếu có.
