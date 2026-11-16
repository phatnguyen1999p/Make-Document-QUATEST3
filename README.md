# QUATEST 3 - Environmental & Cleanroom Monitoring Framework (`quatest3-core`)

Hệ thống tự động hóa xử lý hồ sơ quan trắc môi trường, phòng sạch, khí nén, tủ an toàn sinh học (BSC), tủ hút khí độc (Fume Hood), dòng chảy tầng (LAF) của **Trung tâm Kỹ thuật 3 (QUATEST 3) - Phòng NV2**.

---

## 1. Cấu trúc đóng gói

Dự án đã được cấu hình và đóng gói hoàn chỉnh thành một Python Package chuẩn với đầy đủ:
- **`Core/`**: Mã nguồn toàn bộ logic (Models, Builders, Services, Styling, Tools, Templates).
- **`Core/Templates/`**: Bộ template chuẩn DOCX, XLSX, hình ảnh của QUATEST 3 (tự động đính kèm khi cài đặt).
- **`AGENTS.md`**: Bộ quy tắc chuẩn dành cho AI Agent khi thao tác trên hồ sơ.
- **`pyproject.toml` & `setup.py`**: Cấu hình metadata, dependencies, và lệnh CLI (`quatest3`, `lamkq`).
- **`Context/`**: trạng thái, lỗi, nhật ký — mỗi hồ sơ / phiên / lỗi một file (đầu phiên chạy `python Core/Tools/context.py dau-phien`, AGENTS mục 16); **`PROJECT_CONTEXT.md`**: cách dùng context, bản đồ mã nguồn; **`QUY-TRINH-NGUONG-DANH-GIA.md`**: ngưỡng theo quy trình N2.TC; **`Core/Templates/TAGS.md`**: danh mục TAG template.
- **Thư mục `dist/`** (bản đóng gói ngày 24/10/2026 — **cũ hơn code hiện tại**, dựng lại bằng `python -m build` trước khi cài sang máy khác):
  - `quatest3_core-1.0.0-py3-none-any.whl`: File Wheel cài đặt nhanh không cần build.
  - `quatest3_core-1.0.0.tar.gz`: Bản phát hành mã nguồn chuẩn.
  - `quatest3_core_bundle.zip`: File nén độc lập (chứa `Core/`, `AGENTS.md`, cấu hình).

---

## 2. Cách thêm (add) vào project khác

Bạn có thể thêm và sử dụng theo một trong các cách sau tùy theo nhu cầu:

### Cách 1: Sử dụng trực tiếp trong môi trường Python (Khuyên dùng nhất)
Gói đã được cài đặt ở chế độ Editable (`pip install -e "C:\Users\phat.nt\Desktop\Source\Làm KQ"`).
- Khi bạn sửa code trong `Làm KQ`, các project khác sử dụng package này sẽ được cập nhật ngay lập tức.
- Ở **bất kỳ thư mục nào** hoặc project nào trên máy tính, bạn có thể:

#### Dùng lệnh dòng lệnh (CLI):
```bash
# Xem hướng dẫn trợ giúp
quatest3 --help
# hoặc
lamkq --help

# Tự động quét và xử lý hồ sơ trong thư mục hiện tại
quatest3 --auto .

# Khởi tạo nhanh cấu trúc thư mục chuẩn và AGENTS.md cho hồ sơ mới
quatest3 --init .

# Xử lý intake PDF scan tự động
quatest3 --intake .
```

#### Import trực tiếp trong code Python của project khác:
```python
import Core
from Core import Dossier, DossierRunner, BCKQBuilder, TBKQBuilder
from Core.Builders.handlers import BSCHandler
from Core.Tools.bsc_calc import parse_bsc_excel
from Core.Tools.khinen_calc import parse_compressed_air_workbook

# Sử dụng các công cụ tính toán và tạo văn bản
```

---

### Cách 2: Cài đặt vào môi trường ảo (venv) hoặc máy tính khác bằng file Wheel
Nếu bạn tạo một project mới có `venv` riêng hoặc muốn cài trên máy khác:
```bash
pip install "C:\Users\phat.nt\Desktop\Source\Làm KQ\dist\quatest3_core-1.0.0-py3-none-any.whl"
```
Toàn bộ code, thư viện phụ thuộc và template DOCX/XLSX sẽ được tự động tích hợp đầy đủ.

---

### Cách 3: Thêm trực tiếp thư mục `Core/` vào project mới
Nếu bạn muốn nhúng trực tiếp mã nguồn vào source code của project khác:
1. Giải nén file `dist/quatest3_core_bundle.zip` (hoặc copy thư mục `Core/` và file `AGENTS.md`).
2. Dán thư mục `Core/` và file `AGENTS.md` vào thư mục gốc của project mới.
3. Cài đặt các thư viện cần thiết:
   ```bash
   pip install python-docx openpyxl PyMuPDF lxml matplotlib pillow pywin32 xlrd PyPDF2 markitdown
   ```

---

### Cách 4: Sử dụng cho AI Agent (Antigravity / Gemini / Claude) trong hồ sơ mới
Khi mở một thư mục hồ sơ mới trong workspace (`Input\<6 số>\` — **không chạy trên ổ `Z:\`**: theo `AGENTS.md` mục 0, hồ sơ làm trong `Input`, xong mới robocopy `ThucHien\`, `TBKQ\`, `pdf\` lên Z):
1. Chạy lệnh:
   ```powershell
   quatest3 --init .
   ```
2. Lệnh trên sẽ tự động:
   - Tạo cấu trúc thư mục chuẩn: `ThucHien/`, `TBKQ/`, `TBKQ/Edoc/`, `pdf/`, `Photo/`.
   - Copy file `AGENTS.md` vào thư mục đó để AI Agent tự động nắm toàn bộ quy tắc format, font chữ, khoảng cách ký 3cm, QA checklist.

---

## 3. Danh sách các lệnh CLI hỗ trợ

| Lệnh | Ý nghĩa |
| :--- | :--- |
| `quatest3 --init [thư mục]` | Khởi tạo cấu trúc thư mục hồ sơ chuẩn & copy rule AGENTS.md |
| `quatest3 --auto [thư mục]` | Tự động nhận diện hồ sơ (từ profile.json hoặc Excel tính toán) |
| `quatest3 --intake [thư mục]` | Chạy OCR bóc tách thông tin từ file PDF scan |
| `quatest3 --bsc-excel <file>` | Tạo hồ sơ tủ an toàn sinh học (BSC) từ bảng tính Excel |
| `quatest3 --fumehood-excel <file>` | Tạo hồ sơ tủ hút (Fume Hood) từ bảng tính Excel |
| `quatest3 --khinen-excel <file>` | Tạo hồ sơ khí nén từ bảng tính Excel |
| `quatest3 --split-scan <thư mục>` | Tách PDF scan thành `BBLM.pdf` và `Ghi Chep.pdf` vào `TBKQ/Edoc` |
| `quatest3 --split-all-scans` | Tách scan cho mọi hồ sơ trong `Input\` |
| `quatest3 --status <thư mục>` | Xem tiến độ hồ sơ (`ThucHien\_checkpoint\manifest.json`) và việc tiếp theo |
| `quatest3 --resume` / `--force-stage <ds>` / `--force-ocr` | Chạy tiếp phần chưa xong / làm lại từ giai đoạn / OCR lại (bỏ cache) |
| `quatest3 --upchungthu [all \| <mã>]` | Chèn giấy tiêu đề trang 1 TBKQ + kiểm tra PDF bằng UpChungThu |
| `quatest3 --sample-bsc` / `--sample-laf` / `--sample-fumehood` / `--sample-khinen` | Dựng hồ sơ mẫu (mã thử `Q326D9999xx`, không chép lên Z) |
| `quatest3 --khong-danh-gia` | Tủ hút: form Không đánh giá |
| `quatest3 --both-forms` | Xuất đồng thời cả form Có đánh giá và Không đánh giá |
| `quatest3 --no-pdf` | Bỏ qua bước chuyển đổi sang file PDF |
| `quatest3 --no-qa` | Bỏ qua bước kiểm tra chất lượng (QA) |

Hồ sơ đặc thù (nhiều báo cáo, phòng sạch) chạy bằng script riêng ở gốc workspace: `build_132488.py` / `chay_132488.bat`, `build_cleanroom_132115.py` + `generate_all_reports.py` / `chay_132115.bat` (xem `AGENTS.md` mục 10).

Kiểm tra QA: `python Core\Tools\qa_check.py final-check <thư mục|file> [--code <mã>]` (phải `[CLEAN]`), `check-pdf <thư mục pdf>`.
