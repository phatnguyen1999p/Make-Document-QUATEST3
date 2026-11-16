# QUATEST3 - HỆ THỐNG XỬ LÝ VĂN BẢN (Word, Excel, PDF)

Ứng dụng desktop chuyên dụng hỗ trợ tự động hóa các thao tác xử lý tài liệu kỹ thuật, chứng thư và hồ sơ thử nghiệm môi trường tại **QUATEST3**.

---

## 🌟 Tính Năng Nổi Bật

### 1. Phân hệ Xử lý File Word
- **Tìm kiếm & Cập nhật ngày Header chuẩn xác**: Tự động nhận diện và cập nhật ngày tháng riêng trong Header của tài liệu Word (không làm ảnh hưởng đến ngày thử nghiệm, ngày lấy mẫu hay nội dung trong thân bài).
- **Quy tắc cộng ngày theo mốc giờ**: Tùy chỉnh cộng ngày tự động dựa trên thời điểm xử lý (trước 09:00, trước 16:00, sau 16:00) và tự động bỏ qua Thứ 7, Chủ Nhật.
- **Bộ chọn ngày trực quan (Mini Calendar)**: Cho phép nhập tay hoặc chọn ngày nhanh qua pop-up lịch.
- **Xuất PDF & Chèn Giấy Tiêu Đề**: Tự động chuyển đổi file Word sang PDF và ghép hình nền tiêu đề (`GIAY TIEU DE.pdf` hoặc `GIAY TIEU DE - E.pdf`).
- **Kiểm tra tính hợp lệ đa luồng**: Kiểm tra số lượng ngày ban hành trên PDF (chuẩn 1 trang = 1 ngày, báo chi tiết trang thừa/thiếu), đồng bộ chữ ký (GDV & Boss), số hiệu hồ sơ (`Q3\d{2}[A-D]\d{6}`) và số trang (`i/total`).

### 2. Phân hệ Xử lý File Excel
- **Quét cấu trúc đa file nhanh chóng**: Tự động mở và phân tích danh sách sheet trong toàn bộ thư mục Excel.
- **Gộp Sheet linh hoạt**: Hỗ trợ đánh dấu `[ + ]` để gộp nhiều sheet vào cùng một file PDF duy nhất.
- **Chọn nhanh (Bulk Select)**: Cửa sổ mini cho phép bật/tắt hàng loạt theo tên sheet trên toàn bộ danh sách file.
- **Hai chế độ xuất**:
  - *Xuất tính toán*: Xuất PDF thuần túy không kiểm tra.
  - *Xuất chứng thư*: Tự động chèn giấy tiêu đề và kiểm tra toàn bộ chữ ký, ngày, số hiệu, số trang.

### 3. Tiện ích Hệ thống
- **Tìm kiếm & Mapping hồ sơ (Autocomplete Search)**: Quét nhiều thư mục gốc, nhận diện mã hồ sơ 6 chữ số và tự động tìm thư mục `tbkq`.
- **Thư viện Người ký (GDV & Boss)**: Quản lý danh sách người ký, tự động nhận diện cả tiếng Việt có dấu và không dấu (cho bản tiếng Anh).
- **Cơ chế Đồng bộ Dữ liệu LAN/Local**: Cho phép một máy chủ lưu file cấu hình dùng chung hoặc chạy độc lập trên từng máy.
- **Giao diện Hiện đại**: Hỗ trợ chuyển đổi nhanh Dark Mode / Light Mode.

---

## 📁 Cấu Trúc Dự Án

Dự án được phân tách hoàn toàn giữa **Front-End (UI Tkinter)** và **Back-End (Logic nghiệp vụ & COM Automation)**:

```
UpChungThu/
├── core/                               # [BACK-END] Xử lý nghiệp vụ & COM automation
│   ├── __init__.py                     # Package entry & exports
│   ├── config_manager.py               # Quản lý cấu hình JSON (Local/LAN, reset)
│   ├── utils.py                        # Xử lý dấu tiếng Việt, tính toán ngày
│   ├── pdf_service.py                  # Ghép nền, kiểm tra chữ ký, ngày, ID, trang
│   ├── word_service.py                 # Tự động hóa MS Word COM, tìm & thay thế ngày
│   ├── excel_service.py                # Tự động hóa MS Excel COM, quét & gộp sheet
│   └── mapping_service.py              # Quét thư mục gốc, lập chỉ mục hồ sơ 6 số
│
├── gui/                                # [FRONT-END] Giao diện người dùng Tkinter
│   ├── __init__.py
│   ├── theme.py                        # ThemeManager (Dark/Light), palette màu, fonts
│   ├── app.py                          # UnifiedDocumentSystem (Controller & routing)
│   ├── components/                     # Components dùng lại
│   │   ├── __init__.py
│   │   └── base_processor.py           # Base Frame chứa search bar autocomplete
│   ├── dialogs/                        # Các cửa sổ modal / popup
│   │   ├── __init__.py
│   │   ├── signer_dialog.py            # Quản lý thư viện người ký (GDV & Boss)
│   │   ├── time_rules_dialog.py        # Cài đặt mốc giờ cộng ngày
│   │   ├── calendar_dialog.py          # Pop-up chọn ngày trực quan
│   │   ├── bulk_config_dialog.py       # Cấu hình chọn nhanh & gộp sheet (+)
│   │   └── status_dialog.py            # Tiến trình quét file Excel
│   └── views/                          # Các màn hình chính
│       ├── __init__.py
│       ├── main_menu.py                # Menu chính & Cấu hình dữ liệu LAN
│       ├── word_view.py                # Màn hình phân hệ xử lý Word
│       └── excel_view.py               # Màn hình phân hệ xử lý Excel
│
├── docs/                               # [TÀI LIỆU DỰ ÁN]
│   ├── ARCHITECTURE.md                 # Tài liệu kiến trúc phân lớp & sơ đồ luồng
│   └── USER_GUIDE.md                   # Hướng dẫn sử dụng chi tiết
│
├── main.py                             # Entry point chính
├── UpChungThuv6.py                     # Entry point tương thích ngược
├── UpChungThu_v6.spec                  # Cấu hình đóng gói PyInstaller
├── app_settings.json                   # File cấu hình ứng dụng
├── settings_source.json                # Cấu hình nguồn dữ liệu LAN/Local
├── GIAY TIEU DE.pdf                    # Mẫu nền tiếng Việt
└── GIAY TIEU DE - E.pdf                # Mẫu nền tiếng Anh
```

---

## 🚀 Yêu Cầu & Cài Đặt

### 1. Yêu cầu hệ thống
- Hệ điều hành: **Windows 10 / 11**
- Đã cài đặt **Microsoft Office (Word & Excel)** bản quyền hoặc có thể sử dụng COM automation.
- Python: **3.8+** (Khuyến nghị 3.10+)

### 2. Cài đặt thư viện phụ thuộc
```bash
pip install pywin32 PyPDF2
```

### 3. Chạy ứng dụng từ mã nguồn
```bash
# Sử dụng entry point chính:
python main.py

# Hoặc sử dụng entry point tương thích:
python UpChungThuv6.py
```

---

## 📦 Hướng Dẫn Đóng Gói (PyInstaller)

Để build ứng dụng thành tệp thực thi `.exe` độc lập:

### Cách 1: Build dạng thư mục (`--onedir` - Khuyến nghị cho tốc độ mở nhanh nhất)
```bash
pyinstaller UpChungThu_v6.spec
```

### Cách 2: Build bằng dòng lệnh trực tiếp
```bash
# Build onedir:
pyinstaller --noconsole --onedir --clean --add-data "GIAY TIEU DE.pdf;." --add-data "GIAY TIEU DE - E.pdf;." --add-data "settings_source.json;." --noconfirm --name "UpChungThu_v6" main.py

# Build onefile:
pyinstaller --noconsole --onefile --clean --add-data "GIAY TIEU DE.pdf;." --add-data "GIAY TIEU DE - E.pdf;." --add-data "settings_source.json;." --noconfirm --name "UpChungThu_v6" main.py
```

---

## 📖 Tài Liệu Tham Khảo

- [Tài liệu Kiến trúc Hệ thống (ARCHITECTURE.md)](file:///docs/ARCHITECTURE.md)
- [Sổ tay Hướng dẫn Sử dụng (USER_GUIDE.md)](file:///docs/USER_GUIDE.md)

---
*Bản quyền thuộc Bộ phận Kỹ thuật / Môi trường - QUATEST3.*
