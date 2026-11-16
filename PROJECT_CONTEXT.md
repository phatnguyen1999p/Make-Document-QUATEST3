# PROJECT_CONTEXT — Bản đồ dự án và cách dùng context

> **Dự án:** `quatest3-core` — tự động hóa hồ sơ quan trắc môi trường (phòng sạch, khí nén, BSC, LAF, tủ hút, quan trắc môi trường lao động — QTMTLĐ) của QUATEST 3, Phòng NV2.
> **File này chỉ còn phần tĩnh** (cách dùng context, bản đồ mã nguồn, ghi nhớ nhanh). Từ 12/11/2026 trạng thái, lỗi, nhật ký, bàn giao **không ghi vào đây nữa** — nằm trong `Context/` (mỗi hồ sơ, mỗi phiên, mỗi lỗi một file) để nhiều phiên chạy song song không ghi đè nhau. Bản gốc trước khi tách: `ContextLuuTru/PROJECT_CONTEXT_truoc_khi_tach_20261112.md`.

---

## 0. Cách dùng context (đọc trước khi làm bất cứ việc gì)

**Đầu phiên chạy một lệnh, không cần đọc file này:**

```
python Core/Tools/context.py dau-phien
```

Lệnh in: hồ sơ đang mở (trạng thái, bước tiếp, ai đang giữ khóa), lỗi / việc hệ thống đang mở, điểm cần người dùng quyết, phiên 3 ngày gần đây, và **nhãn phiên gợi ý** (dùng cho `--ai` ở các lệnh dưới). Nó cũng tự lưu trữ phần đã xong.

| Thứ | Ở đâu | Ai được ghi |
|---|---|---|
| Một hồ sơ đang mở: trạng thái, bước tiếp, điểm cần người dùng quyết, quyết định đã chốt, nhật ký | `Context/HoSo/<6 số>.md` | Chỉ phiên đang giữ khóa hồ sơ đó |
| Việc hệ thống của một phiên (sửa Core, AGENTS, công cụ) | `Context/Phien/<ngày-giờ>-<chủ đề>.md` | Chỉ phiên tạo ra file |
| Một lỗi / rủi ro (`E<n>`) hoặc một việc hệ thống (`V<n>`) đang mở | `Context/Viec/<ID>.md` | Phiên đang xử lý việc đó |
| Khóa | `Context/_khoa/<tên>.lock` | Chỉ qua lệnh `khoa` / `tra` |
| Lưu trữ nguyên văn phần đã xong (**không đọc đầu phiên**) | `ContextLuuTru/` | Chỉ qua lệnh `don` / `dong` |
| Quy tắc bất biến | `AGENTS.md` | Chỉ khi người dùng yêu cầu / chốt |

**Quy trình:**

1. **Làm hồ sơ:** `python Core/Tools/context.py khoa <6 số> --ai <nhãn>`. In `[KHÓA BẬN]` → phiên khác đang làm hồ sơ đó: **dừng, báo người dùng**, không dựng, không chép Z. Hồ sơ mới: `moi <6 số> --khach "…" --doi-tuong "…" --ai <nhãn>`. Hồ sơ đã lưu trữ: `mo <6 số> --ai <nhãn>` rồi đọc file lưu trữ nó chỉ ra. Sau đó đọc `Context/HoSo/<6 số>.md` (chỉ file của hồ sơ đang làm).
2. **Ghi lại (sửa thẳng file hồ sơ):** cập nhật 3 dòng đầu — `Trạng thái` (có ngày `dd/mm/yyyy`), `Bước tiếp`, `Cần người dùng quyết` (không còn thì ghi `không`) — và thêm một dòng **trên cùng** bảng Nhật ký: `| dd/mm/yyyy | Claude Code | <việc đã làm>; file: …; QA: … |`. Quyết định của người dùng ghi nguyên lời kèm ngày.
3. **Sửa code chung / quy tắc (Core, template, AGENTS):** `python Core/Tools/context.py phien "<chủ đề>" --ai <nhãn>` → ghi vào file phiên vừa tạo (dòng `Tóm tắt` + bảng nhật ký). Không sửa file phiên của phiên khác.
4. **Lỗi / việc hệ thống mới:** `python Core/Tools/context.py viec "<tiêu đề>" [--loi --muc Thấp|Trung bình|Cao] --ai <nhãn>` (tự cấp ID, không trùng giữa các phiên). Sửa xong: đổi dòng `Trạng thái` thành `Fixed (dd/mm/yyyy, AI)` / `Done (dd/mm/yyyy, AI)` — lần `don` sau tự lưu trữ; còn phần dở cần nhớ thì thêm `[giữ]`.
5. **Xong việc:** `python Core/Tools/context.py tra <6 số> --ai <nhãn>`. Khóa quá 12 giờ không gia hạn coi là khóa cũ (phiên đã tắt), phiên khác được nhận lại; phiên làm lâu thì chạy lại `khoa` để gia hạn.
6. Không biết thì ghi "suy luận" hoặc "chưa xác nhận", không đoán thành sự thật.

**Lệnh khác:** `xem` (tóm tắt, không ghi gì; `--day-du` không cắt dòng) · `so` (sổ mọi hồ sơ đang mở + lưu trữ) · `don [--ghi]` (lưu trữ: hồ sơ không đụng tới ≥ 7 ngày, lỗi / việc đã Fixed / Done, file phiên cũ hơn 7 ngày) · `dong <6 số> --ghi` (đóng hồ sơ ngay). Lệnh cũ `context_archive.py don | dong | mo | kiem-tra` vẫn chạy, chuyển tiếp sang `context.py`.

### Câu lệnh mở đầu phiên (copy và dán)

```
Chạy python Core/Tools/context.py dau-phien, rồi làm việc tôi giao. Làm hồ sơ thì khóa hồ sơ trước, xong thì cập nhật Context/HoSo/<6 số>.md và trả khóa.
```

---

## 5. Bản đồ mã nguồn (tóm tắt)

**Luồng chính:** `Dossier` (Models) → `DossierRunner.run_all()` (Services) → `BCKQBuilder.build()` + `TBKQBuilder.build_all()` → Excel tính toán + tem tủ → `PDFExporter` (Word COM) → `qa_check` → `upchungthu_bridge` (chèn giấy tiêu đề, kiểm tra PDF).

```
Core/
├── run.py            CLI (quatest3 / lamkq), ~1.200 dòng
├── Models/           Dossier, ClientInfo, Equipment, PersonnelInfo, MonitoringLocation
├── Styling/          elements, tables, paragraphs, styles (AutoFit, row height, tblHeader, ending line, Table 0)
├── Builders/         bckq_builder, tbkq_builder
│   └── handlers/     bsc, laf_bsc, fumehood, cleanroom, compressed_air, mtld
├── Exporters/        pdf_exporter (Word/Excel COM)
├── Services/         dossier_runner, upchungthu_bridge, checkpoint (tiến độ + cache OCR)
├── Tools/            bsc_calc, fumehood_calc, khinen_calc, mtld_calc, iso8573_cal, calc_engine, noise_corr (ồn nền),
│                     excel_borders (viền mảnh, gạch chéo), tag_engine, tag_excel, build_tag_templates,
│                     scan_splitter, intake_helper, tax_helper, qa_check, docx_patcher, md_ocr...
└── Templates/        DOCX/XLSX mẫu dạng TAG (BCKQ, TBKQ VN/EN, tem tủ, TinhToan*) — TAGS.md; MTLD/ (bộ mẫu QTMTLĐ 2026); bản gốc _goc_truoc_TAG/
```

`Core/Builders/tag_context.py`: context TAG, tem tủ (`seal_content_lines`, `seal_standard`, `build_tem_tu_sheet` gom tem).

**Script ngoài `Core/`:** `build_132488.py` + `chay_132488.ps1/.bat` (tủ cấy + 3 tủ hút 132488); `build_cleanroom_132115.py`, `fill_excel_132115.py`, `format_excel_132115.py`, `generate_all_reports.py` + `chay_132115.ps1/.bat` (phòng sạch 132115); `build_999901.py` (hồ sơ ảo thử nghiệm QTMTLĐ `Q326D999901`, không ghi lên Z); `chay_thu.ps1/.bat` (chạy thử); cũ: `build_all_tbkq_khinen.py`, `build_hulo_care_laf.py`.

**Phụ thuộc:** python-docx, openpyxl, PyMuPDF, lxml, matplotlib, Pillow, pywin32 (chỉ chạy trên Windows), xlrd, PyPDF2, markitdown.

**Lệnh hay dùng:**

| Việc | Lệnh |
|---|---|
| Tách scan 1 hồ sơ | `python Core/run.py --split-scan "Input\<tên hồ sơ>"` |
| Kiểm tra QA DOCX | `python Core\Tools\qa_check.py final-check <path> [--code <Mã_hồ_sơ>]` |
| Kiểm tra mở Word, đếm trang | `python Core\Tools\qa_check.py test-open <file.docx>` |
| Kiểm tra PDF / chèn nền | `python Core\Tools\qa_check.py check-pdf <path> [--date dd/mm/yyyy] [--add-bg]` |
| BSC mẫu | `python Core/run.py --sample-bsc` |
| BSC từ Excel | `python Core/run.py --bsc-excel "path\TinhToanBSC.xlsx"` |
| Tủ hút từ Excel | `python Core/run.py --fumehood-excel "path\TinhToanFumehood.xlsx"` |
| Khí nén từ Excel | `python Core/run.py --khinen-excel "path"` |
| UpChungThu hàng loạt | `python -m Core.run --upchungthu [all \| <mã_hồ_sơ>]` |
| Xem tiến độ hồ sơ | `python Core/run.py --status "Input\<hồ sơ>"` |
| OCR lại file scan (bỏ cache) | thêm `--force-ocr` vào `--intake` / `--auto` / `--split-scan` |
| Làm lại từ giai đoạn X | thêm `--force-stage ocr,split_scan,...` (giai đoạn sau X tự mất hiệu lực) |
| Chạy tiếp, bỏ qua phần đã xong | thêm `--resume` |
| OCR 1 file có cache | `python Core/Tools/md_ocr.py <file> --dossier "Input\<hồ sơ>" -o out.md` |
| Hồ sơ 132488 (DOCX + XLSX) | `python build_132488.py` (thêm `--pdf` để xuất PDF + UpChungThu; `--only 1,3` chạy một số báo cáo) |
| Hồ sơ 132488 đầy đủ (Windows) | bấm đúp `chay_132488.bat` → log `chay_132488_log.txt` |
| Sinh lại bộ Template chuẩn QTMTLĐ 2026 | `python Core/Tools/mtld_calc.py` |
| Chạy hồ sơ ảo thử nghiệm QTMTLĐ (`Q326D999901`) | `python build_999901.py` (xuất vào `Input\999901\`, không ghi lên Z) |
| Lưu trữ context: xem trước / đóng hồ sơ / dọn / đọc lại | `python Core/Tools/context_archive.py kiem-tra` · `dong <6 số> --ghi` · `don --ghi [--ngay 7] [--giu E46]` · `mo <6 số>` |

---

## 5B. Phân đoạn tiến độ và checkpoint

**Ý tưởng:** mỗi giai đoạn tốn thời gian/token chỉ chạy một lần, kết quả ghi ra file (JSON là bản chính, Excel để xem/sửa tay). Chạy lại thì đọc file thay vì làm lại. Thư mục: `Input\<hồ sơ>\ThucHien\_checkpoint\` (mã: `Core/Services/checkpoint.py`).

| Giai đoạn | Tên | Đầu ra / checkpoint | Khi nào làm lại |
|---|---|---|---|
| 1 | `ocr` | `ocr.json` + `ocr.xlsx` (văn bản từng trang, kèm SHA-256 file scan), `scan_ocr.md` | File scan đổi, đổi dpi, `--force-ocr` |
| 2 | `intake` | `metadata.json`, `profile.json` | Văn bản OCR đổi, hoặc `--force-stage intake` |
| 3 | `split_scan` | `TBKQ/Edoc/<mã> - BBLM.pdf`, `- Ghi Chep.pdf`, `header_ocr.json` | File scan đổi, hoặc `--force-stage split_scan` (không OCR header lại trừ khi `--force-ocr`) |
| 4 | `bckq` | `ThucHien/BCKQ ... .docx` | Luôn dựng lại (chỉ ghi nhận) |
| 5 | `tbkq` | TBKQ VN/EN, tem tủ, Excel tính toán | Luôn dựng lại (chỉ ghi nhận) |
| 6 | `pdf` | `pdf/*.pdf` | Luôn xuất lại (chỉ ghi nhận) |
| 7 | `qa` | kết quả `qa_check` trong `manifest.json` | Luôn chạy lại (chỉ ghi nhận) |

**Quy tắc:**
- `manifest.json` ghi trạng thái từng giai đoạn: `done` / `failed` / `running` / `skipped`, thời gian, file đầu ra, lỗi. `--status` đọc file này và chỉ ra "Việc tiếp theo".
- Làm lại một giai đoạn thì mọi giai đoạn đứng sau nó bị xóa trạng thái (vì dựa trên dữ liệu cũ).
- Chỉ lưu OCR khi có chữ; OCR trả lỗi hoặc rỗng toàn bộ thì không lưu (lần sau thử lại).
- Sửa chữ OCR sai ở `ocr.xlsx` (cột "Văn bản OCR"); lần chạy sau tự nạp lại. Nếu `--force-ocr` sắp ghi đè bản sửa tay, bản cũ được lưu ở `ocr.before_reocr.json`. Đóng Excel trước khi chạy, nếu không `ocr.xlsx` không cập nhật được (`ocr.json` vẫn lưu).
- Lỗi đọc/ghi checkpoint không làm hỏng pipeline, chỉ in `[Checkpoint] ...` rồi chạy như cũ.
- **Cho AI:** cần nội dung file scan thì đọc `ocr.json`/`ocr.xlsx` trước, không đọc ảnh hay OCR lại nếu `--status` báo `ocr` đã xong và file scan không đổi.

---

## 5C. Số phụ chứng thư `/1 /2 /3` (tự động, từ 27/10/2026)

**Quy tắc:** hồ sơ có từ 2 báo cáo kết quả trở lên → tự thêm `/1`, `/2`, `/3`... theo thứ tự báo cáo; hồ sơ chỉ 1 báo cáo → mã trần (AGENTS mục 2).

| Thành phần | Cách hoạt động |
|---|---|
| `Dossier.cert_no_vn` / `cert_no_en` | **Thuộc tính** tính theo `sub_no` (không còn gán cố định lúc tạo): `Q326D…/n` và `Q326D…/n/E`; 1 báo cáo → `Q326D…` và `Q326D…/E`. |
| Chuẩn hóa `sub_no` (`Dossier._normalize_sub_no`) | `is_multi_cert=False` → `sub_no=None`; nhiều báo cáo mà `sub_no` trống/0 → `1`; chuỗi số `"2"` → `2`. |
| `assign_cert_numbers([d1, d2, ...])` (`Core.Models`) | Nhóm theo mã hồ sơ; 1 báo cáo → mã trần; ≥ 2 → giữ số phụ hợp lệ không trùng, còn lại nhận số trống nhỏ nhất theo thứ tự. |
| `DossierRunner.run_batch([...], **kwargs)` | Gọi `assign_cert_numbers` rồi `run_all()` từng báo cáo. |
| Nơi số phụ xuất hiện | Header TBKQ VN/EN; **số BCKQ** (`so_bc`); viện dẫn phiếu tính toán trong BCKQ (BSC/LAF, tủ hút, phòng sạch); header + ô G1 "Số phụ" Excel tính toán; số tem tủ; tên file (`BCKQ - n - …`, `TBKQ - viet - n - …`, `<mã> - n - TinhToan….xlsx`, `tem tu - n - …`). |
| Không đổi | Thân TBKQ (Table 0 dòng 4) vẫn mã trần; viện dẫn phiếu khí nén giữ dạng `<mã> - KN1`; `populate_bckq_table0` (script cũ 132115) giữ mã trần. |
| QA (`qa_check.py`) | Thân BCKQ được phép `<mã>/<số>`; vẫn báo lỗi nếu có `/E` hoặc đuôi khác; TBKQ giữ quy tắc mã trần trong thân. |

Kiểm chứng (Linux, 27/10/2026): 132488 → BCKQ `Q326D132488/1..4`, Excel `Số: Q326D132488/1..4`, tem tủ `/1..4`, 12 DOCX CLEAN; 132918 (1 báo cáo) vẫn mã trần, CLEAN; 132115 (`is_multi_cert=False`) không đổi. **Chưa chạy trên Windows.**

## 5D. Tùy chọn bỏ chỉ tiêu trong handler (thêm 27/10/2026, tương thích ngược)

| Handler | Khóa `results_data` | Tác dụng |
|---|---|---|
| `BSCHandler` | `inflow: {"skip": True}` | Bỏ dòng inflow (tủ cấy LAF) |
| | `hepa: {"skip_exhaust": True}` | Bỏ dòng lọc khí thải |
| | `downflow` / `lux`: `{"no_eval": True}` | Mức quy định "Không quy định", đánh giá "-" |
| | `downflow: {"skip": True}`, `hepa: {"skip_downflow": True}` | Bỏ downflow / lọc khí cấp (tủ ATSH cấp I, 134804/1) — thêm 02/11/2026 |
| | `inflow: {"limit_range": (0.70, 1.00)}` | Mức inflow theo khoảng (EN 12469 tủ cấp I) thay cho ≥ 0,40 m/s |
| | `airflow_pattern: {"rows": [...]}` / `{"skip": True}` | Nội dung hình thái dòng khí riêng (`param_vn/en`, `result_vn/en`, `limit_vn/en`, `assess_vn/en`, `height`) / bỏ mục |
| | `has_assessment=False` (Dossier) | Bảng kết quả 3 cột STT / Chỉ tiêu / Kết quả (bỏ mức + đánh giá), Ghi chú chỉ câu công nhận (134792) — thêm 03/11/2026 |
| | `microbial: {"title_vn/en", "points": [{"name_vn/en", "result_vn/en"}], "diagram", "diagram_en", "caption_vn/en"}` | Chỉ tiêu vi sinh không khí (không mức) + sơ đồ vị trí sau bảng Ghi chú (134792) |
| | `cabinet.cabinet_name_en`, `file_label` | Tên tủ EN trong câu xác nhận; nhãn tên file TBKQ (mặc định `BSC`) |
| `FumeHoodHandler` | `lux: {"skip": True}`, `airflow_pattern: {"skip": True}` | Bỏ chỉ tiêu; STT, nội dung thực hiện, chiều cao dòng tự tính lại |
| `BCKQBuilder` | `tasks_vn` (BSC/LAF), `cabinet.chi_tieu_vn` (tủ hút), `bckq_label` | Danh sách chỉ tiêu thực hiện riêng; nhãn tên file BCKQ |
| Cả 2 handler | — | Ảnh chụp dọc tự xoay 90° để chèn ngang (AGENTS mục 3) |

## 5E. Quan trắc Môi trường Lao động — QTMTLĐ (thêm 15/11/2026)

**Nguồn mẫu khảo sát:** 8 bộ hồ sơ thực tế trong `MTLD/` (`Q326D033771`, `1019N2.25DG`, `1126N2.25DG`, `334N2.25DG`, `435N2.25DG`, `876N2.25Dg`, `891N2.25DG`, `930N2.25DG` — thư mục mẫu chỉ đọc).

| Thành phần | Đường dẫn / Hàm chính | Ghi chú kỹ thuật |
|---|---|---|
| **Bộ Template chuẩn 2026** | `Core/Templates/MTLD/` (và bản alias tại `Core/Templates/`) | Gồm 3 file Excel tính toán (`TinhToan_HoiKhiDoc_template.xlsx`, `TinhToan_ViKhiHau_VatLy_Bui_template.xlsx`, `TinhToan_Ecgonomy_template.xlsx`) và 3 file Word mẫu 2026 (`FORM - TBKQ - viet - mtld - 2026.docx`, `FORM - TBKQ - viet - mtld (Full Master) - 2026.docx`, `FORM - BCKQ - mtld - 2026.docx` chuẩn `BM3030.03 HL:01/8/2025`). |
| **CSDL Hóa chất & Ngưỡng QCVN (`Sheet1`)** | `Core/Tools/mtld_calc.py` (`build_complete_chemical_db`, `lookup_chemical_limit`, `populate_hoi_khi_doc_xlsx`) | Nâng cấp `Sheet1` từ 356 dòng cũ lên **678 dòng tra cứu**: đủ 50/50 chất `QCVN 03:2019/BYT` (chuẩn hóa số thực), đủ 70/70 chất `QCVN 01:2025/BYT` (bổ sung `Butyl acrylat` = 55, `Aceton cyanohydrin` = 4, `Caprolactam (dạng hơi)` = 20, `Đá talc (bụi toàn phần)` = 2), đủ 80/80 chất `QCVN 31:2026/BYT` (HL 09/04/2027), 3 cột `STEL` (15 phút) và 180+ tên đồng nghĩa / công thức / tên trả KQTN để `VLOOKUP` khớp 100%. Giữ công thức `LET` làm tròn 2 chữ số có nghĩa và nhân an toàn khi `KQTN, mg/L` có dấu `<`. |
| **Phiếu tính toán VKH - Vật lý - Bụi - Ecgonomy & Header Excel** | `Core/Tools/mtld_calc.py` (`create_vikhihau_vatly_bui_template`, `create_ecgonomy_template`, `populate_vikhihau_vatly_bui_xlsx`, `populate_ecgonomy_xlsx`, `update_mtld_excel_headers`, `export_mtld_pdfs`) | Sửa triệt để lỗi `#REF!` trên sheet `CO2`, sửa công thức cột Không đạt của `AnhSang` (`=IF(F3<G3,ROUND(F3,0),"-")`), cắt dòng trống thừa `> 25`, bổ sung sheet `Rung` (`QCVN 27:2016/BYT`), `DienTuTruong` (`QCVN 25:2016/BYT`), `GanhNangTriOc`. Chuẩn hóa PageSetup Header góc trái có đuôi phân loại phiếu: `Số: <Mã>/Tính Toán Nhiệt Ẩm`, `Số: <Mã>/Tính Toán Hơi Khí Độc`, `Số: <Mã>/Tính Toán Ecgonomy`; góc phải `Ngày <dd/mm/yyyy>`. |
| **Bố cục bảng TBKQ/BCKQ & Sinh Khuyến nghị giải pháp bằng AI** | `Core/Builders/handlers/mtld.py` (`MTLDHandler`, `build_mtld_reports_from_template`, `ensure_mtld_table_headers_and_cantsplit`, `fix_mtld_header_and_page_fields`, `generate_mtld_recommendations`, `apply_mtld_recommendations_to_doc`, `sync_mtld_summary_table`) | Tự động đặt lặp lại dòng tiêu đề bảng (`<w:tblHeader/>`) khi sang trang; tách khối dòng không viền `"Tổng hợp kết quả quan trắc:"` / `"Ghi chú:"` ra bảng không viền liền kề để không lặp lại header bảng trống phía trên dòng tổng hợp; lược bỏ toàn bộ dòng trống thừa giữa các mục (`I..IV`) và cuối tài liệu; giữ nguyên khối bảng chữ ký trên cùng trang cuối (`Trang {PAGE}/{NUMPAGES}` động); gắn `noWrap` cho ô `< 0,94`. Tự động gọi LLM API hoặc `_compose_contextual_recommendations` để viết 5 nhóm giải pháp với giọng văn nhẹ nhàng, mang tính ghi nhận và sẽ khắc phục. Đăng ký `MTLDHandler` vào `TBKQBuilder`. |
| **Hồ sơ ảo thử nghiệm** | `build_999901.py` $\rightarrow$ `Input/999901/` | Kiểm thử đầu-cuối mã ảo `Q326D999901` (3 file Excel + TBKQ + BCKQ + 5 file PDF trong `Input/999901/pdf/`, khóa cứng không ghi lên `Z:`). |


## 6. Ghi nhớ nhanh các quy tắc dễ sai (chi tiết xem `AGENTS.md`)

- **Nơi lưu file (mục 1):** kết quả vào `Input\<hồ sơ>\ThucHien\` (BCKQ, TinhToan xlsx), `TBKQ\` (TBKQ, tem tủ), `TBKQ\Edoc\` (scan tách), `pdf\` (mọi PDF). **Không** lưu vào thư mục `Output\` và không lưu ở gốc hồ sơ.
- **Mã hồ sơ (mục 2):** thân TBKQ dùng mã trần; đuôi `/1`, `/2`, `/E` chỉ ở Số chứng thư trên header. Hồ sơ nhiều báo cáo: **tự đánh số phụ** bằng `assign_cert_numbers` / `DossierRunner.run_batch` — không gõ tay `sub_no` cho từng báo cáo; số BCKQ và số phiếu tính toán đi theo số chứng thư `/n` (mục 5C).
- **Số liệu (mục 6):** dấu phẩy ở bản Việt, dấu chấm ở bản Anh; kết quả đo bằng `0` phải ghi `0`, không ghi `-`.
- **Excel (mục 11):** giữ nguyên công thức động, không ghi đè bằng giá trị tĩnh.
- **BCKQ (mục 7):** không viện dẫn TBKQ/chứng thư; không có đường kết thúc; font 13 toàn bộ.
- **Ảnh (mục 3):** chỉ đặt `width`, không đặt `height`, không crop.
- **Mục 0 AGENTS:** xong hồ sơ thì chép `ThucHien`/`TBKQ`/`pdf` lên Z: (robocopy, không xóa trên Z:); **trước khi chép, so thời gian sửa file trên Z — người dùng sửa trên Z thì chép về Input trước, không ghi đè** (bài học 132115 phiên 18); không dùng máy in tem HPL; sai lệch thông tin khách hàng → báo người dùng quyết.
- **PDF (AGENTS mục 4):** dựng lại DOCX/XLSX thì phải xuất lại PDF; kiểm tra `Trang X/Y`, số `/n` trên PDF BCKQ.
- **Commit file về máy (Claude):** `device_commit_files` có thể ghi bản cũ nếu commit ngay sau khi ghi — chờ vài giây rồi kiểm tra kích thước file trên máy.

---

## 7. Quy tắc kỹ thuật khi nhiều phiên cùng làm

- **Mỗi hồ sơ một phiên** (khóa, mục 0). Hai phiên được chạy song song trên hai hồ sơ khác nhau.
- **File dùng chung** (`Core/`, template, `AGENTS.md`): trước khi sửa một file, kiểm tra thời gian sửa đổi — mới hơn lần cuối bạn đọc thì đọc lại, vì phiên khác có thể vừa sửa. Sửa từng chỗ nhỏ bằng thay chuỗi, **không ghi đè cả file** từ bản đã đọc trước đó. Sửa lớn một file Core (đổi cấu trúc, đổi chữ ký hàm) thì khóa trước: `khoa <tên file không đuôi, vd tag_context> --ai <nhãn>`, xong `tra`.
- Sửa Core xong, hồ sơ của phiên khác có thể đang dựng dở bằng bản cũ → ghi rõ trong file phiên những hồ sơ / handler bị ảnh hưởng.
- File văn bản trong dự án dùng **UTF-8 (không BOM)**. `AGENTS.md` và nhiều file `.py` dùng xuống dòng CRLF, giữ nguyên kiểu xuống dòng của từng file khi sửa.
- Sửa file đã có thì sửa tại chỗ, không tạo bản sao đổi tên. Muốn xóa hoặc dọn file phải hỏi người dùng.
- Không hard-code đường dẫn máy (`C:\Users\...`, `Z:\...`) vào code mới; lấy từ tham số, biến môi trường hoặc `os.path.expanduser("~")`.
- Không nuốt lỗi bằng `except: pass` trong code mới; ít nhất phải in cảnh báo.
- Không `taskkill` Word/Excel của người dùng. Nếu thật sự cần, dùng cờ `QUATEST3_FORCE_KILL_OFFICE=1` do người dùng bật.
- Sau khi sửa code: chạy `qa_check.py final-check` trên tài liệu bị ảnh hưởng, ghi kết quả (`[CLEAN]` hay lỗi) vào nhật ký.

---

## 8. Mẫu dòng cập nhật

**Nhật ký (file hồ sơ hoặc file phiên), dòng mới ở trên cùng:** `| dd/mm/yyyy | Claude Code/Antigravity | <việc đã làm>; file: <danh sách>; QA: <kết quả> |`

**Đầu file hồ sơ:** `- **Trạng thái:** Chờ KQTN (12/11/2026) — bản nháp đã lên Z` · `- **Bước tiếp:** …` · `- **Cần người dùng quyết:** (a) …; (b) …` hoặc `không`.
