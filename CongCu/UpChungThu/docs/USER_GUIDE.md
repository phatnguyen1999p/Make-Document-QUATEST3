# SỔ TAY HƯỚNG DẪN SỬ DỤNG - HỆ THỐNG XỬ LÝ VĂN BẢN (QUATEST3)

Tài liệu hướng dẫn thao tác chi tiết từng bước cho người sử dụng phần mềm **UpChungThu**.

---

## MỤC LỤC
1. [Giới thiệu & Giao diện chung](#1-giới-thiệu--giao-diện-chung)
2. [Cấu hình Dữ liệu Dùng chung qua Mạng LAN](#2-cấu-hình-dữ-liệu-dùng-chung-qua-mạng-lan)
3. [Quản lý Thư viện Người ký (GDV & Boss)](#3-quản-lý-thư-viện-người-ký-gdv--boss)
4. [Hướng dẫn Phân hệ Xử lý File Word](#4-hướng-dẫn-phân-hệ-xử-lý-file-word)
5. [Hướng dẫn Phân hệ Xử lý File Excel](#5-hướng-dẫn-phân-hệ-xử-lý-file-excel)
6. [Ý nghĩa các Ký hiệu & Mã Lỗi Kiểm tra](#6-ý-nghĩa-các-ký-hiệu--mã-lỗi-kiểm-tra)
7. [Khắc phục sự cố thường gặp (Troubleshooting)](#7-khắc-phục-sự-cố-thường-gặp)

---

## 1. Giới thiệu & Giao diện chung

Khi khởi động ứng dụng, màn hình chính gồm:
- **📊 THAO TÁC VỚI FILE EXCEL**: Mở phân hệ xử lý, quét sheet và xuất PDF từ file Excel.
- **📝 THAO TÁC VỚI FILE WORD**: Mở phân hệ cập nhật ngày, xuất PDF và kiểm tra tài liệu Word.
- **🗑️ KHÔI PHỤC CÀI ĐẶT MẶC ĐỊNH**: Xóa toàn bộ cấu hình đã lưu và đưa ứng dụng về trạng thái ban đầu.
- **🌙 / ☀️ Nút Dark Mode (Góc trên bên phải)**: Chuyển đổi giao diện sáng/tối bảo vệ mắt.
- **🔌 CẤU HÌNH DỮ LIỆU**: Thiết lập chế độ lưu dữ liệu trên máy tính cá nhân hoặc qua mạng nội bộ LAN.

---

## 2. Cấu hình Dữ liệu Dùng chung qua Mạng LAN

Nếu cơ quan có nhiều máy tính cùng xử lý hồ sơ:
1. Đặt tệp `app_settings.json` tại một thư mục chia sẻ trên ổ mạng (ví dụ: `Z:\16.KHAC\UpChungThu\app_settings.json`).
2. Tại màn hình chính, ở mục **CẤU HÌNH DỮ LIỆU**, chọn radio button **Dùng chung (LAN)**.
3. Nhấn nút 📁 và chọn tới tệp `app_settings.json` trên ổ mạng.
4. Hệ thống sẽ tự động đồng bộ danh sách thư viện người ký, thư mục mapping và các quy tắc giờ.

---

## 3. Quản lý Thư viện Người ký (GDV & Boss)

Để phần mềm có thể tự động kiểm tra chữ ký trên các chứng thư:
1. Nhấn nút **👤 Cấu hình Chữ ký** (có ở cả màn hình Word và Excel).
2. Nhập họ tên người ký vào ô **Nhập tên mới**:
   - Nhấn **+ GDV**: Thêm vào danh sách Giám định viên / Thử nghiệm viên.
   - Nhấn **+ Boss**: Thêm vào danh sách Lãnh đạo phê duyệt / Trưởng phòng.
3. Để xóa một người khỏi danh sách, nhấn biểu tượng 🗑️ tương ứng.
4. Nhấn **LƯU VÀ ĐÓNG** để hoàn tất.

> [!NOTE]
> Hệ thống hỗ trợ nhận diện tự động cả tên tiếng Việt có dấu (ví dụ: *Phạm Quốc Nghĩa*) và tên không dấu (ví dụ: *Pham Quoc Nghia*) đối với các chứng thư tiếng Anh (`- E`).

---

## 4. Hướng dẫn Phân hệ Xử lý File Word

### Bước 1: Chọn thư mục hồ sơ
- **Cách 1 (Tìm kiếm nhanh qua Mapping)**:
  - Nhập mã hồ sơ (ví dụ: `26A01234`) vào ô **🔍 Tìm hồ sơ**.
  - Danh sách gợi ý sẽ xổ xuống. Nhấp chọn hoặc nhấn phím `↓` và `Enter`.
  - Phần mềm sẽ tự động tìm và chọn đúng thư mục `tbkq`.
  - Nếu chưa có thư mục gốc, nhấn nút **➕ Thêm gốc** để thêm thư mục chứa hồ sơ năm.
- **Cách 2 (Chọn thủ công)**:
  - Nhấn **📁 Chọn thư mục** và duyệt tới thư mục chứa các file Word cần xử lý.

### Bước 2: Thiết lập cấu hình xử lý
- **☑ Cập nhật Ngày**: Tự động tìm và đổi ngày trên toàn bộ file.
  - **☑ Theo giờ**: Tự động cộng ngày dựa theo thời điểm thực tế (Nhấn nút ⚙️ để cài đặt mốc giờ).
    - Ví dụ: Trước 09:00 cộng 0 ngày, từ 09:00 - 16:00 cộng 1 ngày, sau 16:00 cộng 2 ngày (tự động bỏ qua Thứ 7, Chủ Nhật).
  - **☑ Nhập tay**: Nhập ngày cố định `DD/MM/YYYY` hoặc nhấn 📅 để chọn từ lịch pop-up.
- **☑ Xuất PDF**: Tự động xuất sang tệp PDF có cùng tên.
- **☑ Chèn Giấy Tiêu Đề & Kiểm tra PDF**: Ghép nền giấy tiêu đề lên trang 1 và kích hoạt bộ kiểm tra lỗi tổng thể.

### Bước 3: Thực hiện & Xem kết quả
- Nhấn **▶ BẮT ĐẦU XỬ LÝ WORD**.
- Bảng kết quả sẽ hiển thị tiến độ và trạng thái chi tiết của từng file.
- Dòng màu **Xanh lá (Good)** biểu thị file hoàn toàn hợp lệ; dòng màu **Đỏ (Bad)** báo lỗi chi tiết.

---

## 5. Hướng dẫn Phân hệ Xử lý File Excel

### Bước 1: Chọn thư mục chứa file Excel
- Nhấn **📁 1. Chọn thư mục Excel**.
- Hệ thống sẽ hiển thị thanh tiến trình quét nhanh toàn bộ các sheet có trong các file Excel.

### Bước 2: Chọn nơi lưu PDF (Tùy chọn)
- Mặc định tệp PDF xuất ra sẽ nằm cùng thư mục file gốc.
- Nếu muốn lưu ra nơi khác, nhấn **🎯 2. Chọn nơi lưu PDF**.

### Bước 3: Lựa chọn Sheet và Gộp Sheet (+)
- **Chọn thủ công từng ô**: Nhấp chuột trực tiếp vào ô tương ứng trên bảng để chuyển đổi giữa `☐` (bỏ chọn) và `☑` (chọn).
- **Chọn nhanh hàng loạt**: Nhấn **🛠️ Cấu hình Chọn nhanh & Gộp sheet**:
  - Tích vào ô **Tất cả [Tên Sheet]** để chọn sheet đó trên toàn bộ danh sách file.
  - Tích vào ô **Gộp (+)** giữa các sheet để ghép nối chúng thành một tệp PDF duy nhất.

### Bước 4: Xuất PDF
- **▶ Xuất tính toán (Không kiểm tra)**: Dùng cho các bảng tính nội bộ, không chèn giấy tiêu đề và không kiểm tra lỗi.
- **▶ Xuất chứng thư (Có kiểm tra)**: Dùng cho chứng thư phát hành, tự động chèn giấy tiêu đề và kiểm tra đầy đủ chữ ký, ngày, số hiệu, số trang.

---

## 6. Ý nghĩa các Ký hiệu & Mã Lỗi Kiểm tra

| Ký hiệu / Thông báo | Ý nghĩa & Hướng dẫn xử lý |
| :--- | :--- |
| **✅ OK** | Tài liệu hợp lệ 100% (Ngày đúng, đủ chữ ký, đúng số hiệu Q3, đủ số trang). |
| **❌ Lỗi BG (VN/EN)** | Không tìm thấy tệp mẫu `GIAY TIEU DE.pdf` hoặc `GIAY TIEU DE - E.pdf`. |
| **❌ Ngày (...)** | Ngày trong tài liệu không khớp hoặc không đủ số lần xuất hiện trên các trang. |
| **❌ ID (MISSING)** | Không tìm thấy mã số hồ sơ định dạng `Q3\d{2}[A-D]\d{6}`. |
| **❌ ID (CHECK P1(thiếu /E header) / P1(thiếu số gốc body) / P2(thiếu /E) / P1(thừa /E))** | Bản Tiếng Anh thiếu `/E` ở header các trang hoặc thiếu số gốc ở body trang 1; hoặc bản Tiếng Việt bị thừa `/E`. |
| **❌ ID (CHECK Không khớp mã...)** | Mã số hồ sơ không đồng nhất giữa các trang trong cùng tài liệu. |
| **❌ Trang (Miss: P3)** | Bị thiếu hoặc sai định dạng đánh số trang tại Trang 3. |
| **❌ Chữ ký (THIẾU GDV/Boss)** | Chưa tìm thấy tên của GDV hoặc Boss trong thư viện người ký. |
| **❌ Chữ ký (KHÁC MẪU)** | Người ký ở file này không đồng nhất với người ký ở file mẫu đầu tiên trong cùng folder. |
| **➖ N/A / Bỏ qua** | Tùy chọn kiểm tra bị tắt hoặc tài liệu không yêu cầu kiểm tra. |

---

## 7. Khắc phục sự cố thường gặp (Troubleshooting)

### 1. Báo lỗi "Không tìm thấy file mẫu nền" khi khởi động
- **Nguyên nhân**: Thiếu 2 tệp `GIAY TIEU DE.pdf` và `GIAY TIEU DE - E.pdf` trong thư mục chạy phần mềm.
- **Khắc phục**: Sao chép 2 tệp PDF mẫu vào cùng thư mục với tệp chạy của ứng dụng.

### 2. Xử lý file Word/Excel bị treo hoặc báo lỗi COM
- **Nguyên nhân**: Có một tiến trình Word/Excel ẩn bị treo từ trước trên hệ điều hành Windows.
- **Khắc phục**: Mở Task Manager (Ctrl + Shift + Esc), tìm và End Task các tiến trình `Microsoft Word` hoặc `Microsoft Excel` đang chạy ngầm, sau đó khởi chạy lại phần mềm.

### 3. Không lưu được cấu hình qua mạng LAN
- **Nguyên nhân**: Máy trạm không có quyền ghi (Write Permission) vào thư mục chia sẻ trên mạng.
- **Khắc phục**: Kiểm tra quyền truy cập mạng hoặc chuyển về chế độ **Máy này (Local)**.

---
*Mọi thắc mắc kỹ thuật vui lòng liên hệ phòng Công nghệ Thông tin / Kỹ thuật.*
