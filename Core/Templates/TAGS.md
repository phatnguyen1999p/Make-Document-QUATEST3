# Danh mục TAG của template (Core/Templates)

> Cập nhật 26/10/2026 (Claude); E20, E25 đã sửa 30/10/2026 (rà lại 30/10/2026 phiên 21). Template trong `Core/Templates/` là **form dạng TAG**: không còn dữ liệu của hồ sơ cũ,
> mọi chỗ cần điền là TAG. Bản gốc (còn dữ liệu cũ) nằm ở `Core/Templates/_goc_truoc_TAG/` — không sửa, dùng để dựng lại.

## 1. Cú pháp

| Cú pháp | Ý nghĩa | Ví dụ |
|---|---|---|
| `{{ ten }}` | Thay bằng giá trị. `\n` trong giá trị → xuống dòng, `\t` → tab | `{{ kh_ten }}` |
| `{{ a.b }}`, `{{ a[0] }}` | Thuộc tính / phần tử | `{{ r.v[2] }}` |
| `{{ x if dk else y }}` | Biểu thức Python đơn giản | `{{ ';' if not loop.last else '.' }}` |
| `{%p if dk %}` … `{%p elif dk %}` … `{%p else %}` … `{%p endif %}` | Giữ / xóa các đoạn, bảng nằm giữa (TAG phải nằm riêng 1 đoạn) | xóa cả mục khi chỉ tiêu không đo |
| `{%p for x in ds %}` … `{%p endfor %}` | Lặp các đoạn/bảng nằm giữa | 1 bảng cho mỗi nhóm phòng |
| `{%tr if dk %}` … `{%tr endif %}` | Giữ / xóa các dòng bảng nằm giữa (TAG nằm riêng 1 dòng) | dòng "Mức yêu cầu" |
| `{%tr for r in ds %}` … `{%tr endfor %}` | Lặp các dòng bảng nằm giữa | mỗi điểm đo 1 dòng |

- Trong vòng lặp có `loop.index` (1, 2, …), `loop.index0`, `loop.first`, `loop.last`, `loop.length`.
- Biến không có dữ liệu → để trống và in cảnh báo `[TAG] … thiếu dữ liệu`. `qa_check.py` báo lỗi `[Leftover TAG]` nếu còn TAG chưa điền.
- Ảnh/lưới ảnh: đoạn chỉ chứa `{{ bang_anh }}` được code thay bằng bảng ảnh 1/2/3 cột theo AGENTS.md mục 3 (chỉ đặt chiều rộng, không crop).
- Bộ máy điền: `Core/Tools/tag_engine.py` (không cần cài thêm thư viện; cú pháp tương thích docxtpl).
- Sửa template bằng Word: chỉ sửa chữ/định dạng xung quanh; giữ nguyên TAG (mỗi TAG nên gõ liền trong một lần, không đổi định dạng giữa chừng TAG).

## 2. Dựng lại template

```
python -m Core.Tools.build_tag_templates                 # dựng tất cả từ _goc_truoc_TAG/
python -m Core.Tools.build_tag_templates --only bckq     # khi_nen, phong_sach, tem_tu, bckq, excel
```

## 3. Trang bìa TBKQ (header + Table 0) — chung cho mọi TBKQ

Dùng trong: `TBKQ - vn/viet/en - phong sach.docx`, `TBKQ - vn/viet/en - khi nen.docx` (BSC, LAF, tủ hút dùng trang bìa của `TBKQ - vn/en - phong sach.docx`).

| TAG | Nội dung | Nguồn (Dossier) |
|---|---|---|
| `so_ct`, `ngay_ct` | Số chứng thư (có /1, /E) và ngày ban hành ở header | `cert_no_vn/en`, `cert_date` |
| `ma_hs`, `ngay_dk` | Mã hồ sơ trần + ngày giấy yêu cầu (Table 0 dòng 4) | `code`, `reg_date` |
| `kh_ten`, `kh_dia_chi` | Đơn vị yêu cầu, địa chỉ | `client.*` |
| `noi_dung_yc` | Nội dung yêu cầu giám định | `scope_vn/en` |
| `ngay_gd` | Ngày giám định (template ghi sẵn "Từ ngày …" / "From …") | `sample_time_*` / `sample_date` |
| `dia_diem` | Địa điểm giám định | `location_desc_*` (trống → "khách hàng - địa chỉ") |
| `quy_trinh` | Quy trình thực hiện | `procedure_*`, trống → mặc định theo loại (PS.01 / KN.01 / BSC.02 / LAF.01 / T.01) |
| `thiet_bi` | Danh sách thiết bị, mỗi dòng "- …;", dòng cuối "." | `equipments` |
| `can_cu` | Căn cứ áp dụng/đánh giá | `standards_*` |
| `ket_qua_gd` | "Xem Phụ lục đính kèm" | `results_pointer_*` |
| `gdv_chuc_danh`, `gdv_ten`, `ld_chuc_danh_1`, `ld_chuc_danh_2`, `ld_ten` | Chữ ký (GĐV ở dòng 2, ngang PHÓ GIÁM ĐỐC; khoảng ký 3 cm có sẵn trong template) | `personnel.sign_*` |

Footer "Trang X/Y" dùng trường Word `PAGE`/`SECTIONPAGES` → tổng số trang Phụ lục luôn đúng (không ghi cứng).

## 4. TBKQ khí nén (`TBKQ - vn/en - khi nen.docx`) — `CompressedAirHandler`

| TAG | Nội dung |
|---|---|
| `so_vi_tri` | "02 vị trí" / "02 locations" |
| `trang_thai_he_thong`, `trang_thai_nha_xuong` | mặc định "Đang hoạt động" / "In operation"; ghi đè: `results_data["system_status_vn/en"]`, `["factory_status_vn/en"]` |
| `so_do_he_thong` | Sơ đồ hệ thống (các thiết bị nối mũi tên). Dữ liệu: `results_data["system_chain_vn/en"]` = list hoặc chuỗi "A → B → C". Không có → "………" |
| `vi_tri[]` → `vt.ten`, `vt.ky_hieu`, `vt.chi_tieu[]` | Bảng vị trí |
| `co_anh`, `bang_anh` | Ảnh vị trí (Dossier.photos) |
| `dau.dong[]` (`vi_tri`, `gia_tri`, `phan_loai`), `dau.yc`, `dau.yc_phan_loai`, `dau.danh_gia` | Hàm lượng dầu. `dau.yc` lấy từ `results_data["oil"]["requirement"]` hoặc `parsed_points[*].total_oil.requirement` |
| `dp.dong[]` (`vi_tri`, `v[0..5]`), `dp.yc[0..5]`, `dp.yc_co`, `dp.danh_gia` | Điểm sương; mức yêu cầu hàm lượng nước |
| `hat.dong[]` (`vi_tri`, `v[0..2]`, `phan_loai`, `mo_ta_5um`), `hat.co_khong_phan_loai`, `hat.yc`, `hat.yc_dong[]` (`vi_tri`, `d05`), `hat.danh_gia` | Nồng độ hạt |

Không có mức yêu cầu → xóa dòng "Mức yêu cầu", ghi chú (2) và câu "Đánh giá" tương ứng. Kết luận phù hợp/không phù hợp được **tự tính** so với mức yêu cầu; ghi đè bằng `results_data["oil"|"dewpoint"|"particles"]["assessment_vn/en"]`. Không tính được → "………".

## 5. TBKQ phòng sạch (`TBKQ - viet/vn/en - phong sach.docx`) — `CleanroomHandler`

Mục nào không có dữ liệu thì bị xóa cả mục (AGENTS.md mục 3). Tên khu vực bản EN tự đổi theo `locations[].name_en` (hoặc truyền `khu_vuc_en`).

| TAG | Nội dung |
|---|---|
| `ds_khu_vuc`, `trang_thai` | “A”; “B” — trạng thái phòng sạch (`locations[0].state_*`) |
| `khu_vuc[]` → `kv.ten`, `kv.dien_tich`, `kv.cap`, `kv.chi_tieu[]` | Bảng khu vực (`area_m2`, `clean_class`, `points_summary_*`) |
| `co_anh`/`bang_anh`, `co_so_do`/`bang_so_do` | Ảnh (`photos`), sơ đồ (`layouts`) |
| `bui3[]`, `bui6[]` → `g.cap`, `g.gioi_han[]`, `g.trang_so_do`, `g.diem[]` (`khu_vuc`, `vi_tri`, `v[]`), `g.ds_khu_vuc`, `g.ket_luan` | Nồng độ bụi 3 / 6 kênh, mỗi nhóm 1 bảng |
| `vi_sinh[]` → `g.cap`, `g.gioi_han`, `g.trang_so_do`, `g.diem[]` (`khu_vuc`, `vi_tri`, `gia_tri`) | Vi sinh |
| `nhiet_am` → `dong[]` (`khu_vuc`, `nhiet_do`, `do_am`, `do_on`), `muc_nhiet_do`, `muc_do_am`, `muc_do_on`, `ds_khu_vuc`, `ket_luan` | Nhiệt độ - độ ẩm - độ ồn |
| `chenh_ap` → `dong[]` (`khu_vuc_1`, `khu_vuc_2`, `gia_tri`, `muc`), `ket_luan` | Chênh áp |
| `hepa[]` → `g.diem[]` (`khu_vuc`, `vi_tri`, `gia_tri`), `g.muc`, `g.trang_so_do`, `g.ds_khu_vuc`, `g.ket_luan` | Độ rò rỉ HEPA (`g.muc`: mặc định `≤ 0,01`; nhóm có `cap_loc: "H13"` → `≤ 0,1`; hoặc ghi `muc`) |
| `scan` → `khu_vuc`, `so_do` (ảnh), `trang_so_do`, `diem[]` (`khu_vuc`, `vi_tri`, `hepa`, `gasket`) | Quét rò rỉ (Scanloc) |
| `luu_luong[]` → `g.muc_ach`, `g.trang_so_do`, `g.diem[]` (`khu_vuc`, `vi_tri`, `luu_luong`, `tong`, `the_tich`, `ach`), `g.ds_khu_vuc`, `g.ket_luan` | Lưu lượng / ACH |

Ô "Khu vực" (và tổng lưu lượng / thể tích / ACH) của các điểm cùng phòng được gộp dọc tự động.

Ví dụ `results_data` (rút gọn):

```python
results_data = {
    "bui_3kenh": [{"cap": "ISO 8", "gioi_han": [3520000, 832000, 29300], "trang_so_do": "1",
                   "diem": [{"khu_vuc": "Phòng sản xuất", "vi_tri": 1, "gia_tri": [5368, 2331, 71]}]}],
    "nhiet_am": {"dong": [{"khu_vuc": "Phòng sản xuất", "nhiet_do": "22,5", "do_am": "55,1", "do_on": "62,3"}],
                 "muc_nhiet_do": "18 - 25", "muc_do_am": "40 - 65", "muc_do_on": "≤ 65"},
    "chenh_ap": {"dong": [{"khu_vuc_1": "Phòng sản xuất", "khu_vuc_2": "Air lock", "gia_tri": "12,5", "muc": "≥ 10"}]},
    "hepa": [{"trang_so_do": "2", "diem": [{"khu_vuc": "Phòng sản xuất", "vi_tri": 1, "gia_tri": "< 0,0001"}]}],
    "luu_luong": [{"muc_ach": 20, "trang_so_do": "2",
                   "diem": [{"khu_vuc": "Phòng sản xuất", "vi_tri": 1, "luu_luong": 850}],
                   "phong": {"Phòng sản xuất": {"tong": 1670, "the_tich": "418,8", "ach": "4,0"}}}],
}
```

Kết luận tự tính: bụi ≤ giới hạn; HEPA ≤ g.muc (mặc định 0,01 %, H13 0,1 %); ACH ≥ mức; nhiệt ẩm/chênh áp theo mức dạng "≤ x", "≥ x", "a - b". Ghi đè: `ket_luan_vn` / `ket_luan_en` trong từng nhóm.

## 6. BCKQ (`BCKQ - khi nen / phong sach / bsc / fumehood.docx`) — `BCKQBuilder`

Chung: `so_bc`, `kh_ten`, `kh_dia_chi`, `noi_dung_yc`, `dia_diem`, `tg_th` (thời gian chi tiết), `thiet_bi`, `nguoi_th`, `ngay_kt`, `tre_han`, `tg_tre_han`, `nn_tre_han`, `bat_thuong`, `tab_bc` (7 hoặc 6 tab), `ngay_bc`, `nguoi_bc`.

| Form | TAG riêng |
|---|---|
| Khí nén | `vi_tri[]` (`ten`, `ky_hieu`, `chi_tieu_dong`, `kqtn`, `kqtt`), `so_do_he_thong`, `co_anh`/`bang_anh`, `co_dau`, `co_hat`, `co_dp`, `ma_hs`, `ngay_ct`, `kqtn[]` (`stt`, `so`, `ngay`, `ngay_gui`, `ngay_yc`, `ngay_nhan`, `tre_han`) — dữ liệu: `results_data["kqtn"]`, `results_data["kqtn_by_location"]` |
| Phòng sạch | `ds_khu_vuc`, `trang_thai`, `khu_vuc[]`, `bang_anh`, `bang_so_do`, `tham_chieu[]` (chỉ viện dẫn phiếu tính toán/ghi chép, không viện dẫn TBKQ), `danh_gia[]` (`khu_vuc`, `dat`, `khong_dat`, `can_cu`) — dữ liệu: `assessment["rows"]` |
| BSC / LAF | `tu[]` (`thong_tin`, `chi_tieu`), `tham_chieu[]`, `danh_gia` (TagBlock `assessment_block`: 3 phần phù hợp / không phù hợp / không đánh giá, từ 30/10/2026) |
| Tủ hút | như BSC / LAF (từ 30/10/2026 bỏ bảng `dat` / `khong_dat` / `can_cu` của form gốc, thay bằng `{{ danh_gia }}` để gộp được BCKQ nhiều tủ — `BCKQBuilder.build_merged`) |

`tu[]` lấy từ `results_data["cabinets"]` (nhiều tủ) hoặc `results_data["cabinet"]`.

> **E20 (đã sửa 30/10/2026):** header `BCKQ - bsc / fumehood / khi nen.docx` từng có 2 trường `NUMPAGES` liền nhau → PDF in `Trang 1/22`; `build_tag_templates.use_page_fields` nay bỏ qua đoạn đã có trường. Khi sửa template phải còn đúng 1 trường `PAGE` + 1 trường `NUMPAGES`.

## 7. Tem tủ (`tem tu.docx`, `tem tu - fumehood.docx`)

`so_tem`, `ten_thiet_bi`, `serial`, `model`, `ngay_kiem_tra`, `ngay_kiem_tra_tiep` (+1 năm). **Mỗi tủ = 2 tem, xếp dọc ở cột 1** của lưới 2 dòng x 3 cột (người dùng chốt 30/10/2026; `build_tag_templates._keep_two_seals`). Lề trang tem: trên/dưới 0,5 cm, trái/phải 1 cm (`_seal_page_margins`) → 1 trang A4 ngang chứa tối đa **6 tem = 3 tủ** (mỗi cột 1 tủ; cột 2, 3 để trống cho tủ khác). Áp dụng cả tem BSC/LAF và tủ hút. Ô "Nội dung kiểm tra" cao cố định 1250 twips để mọi tem cùng kích thước. Hồ sơ nhiều tủ: `tag_context.build_tem_tu_sheet` gom tem các tủ vào `TBKQ\<6 số> - tem tu.docx`, 3 tủ / trang (tự gọi trong `DossierRunner.run_batch`, `build_132488.py`). BSC/LAF: tên thiết bị mặc định “Tủ cấy an toàn sinh học cấp 2” (ghi đè: `cabinet["seal_item_name"]`).

`noi_dung_kt` (vòng lặp `{%p for dong in noi_dung_kt %}` / `{{ dong }}` / `{%p endfor %}`, sửa E25 30/10/2026): dòng "Nội dung kiểm tra" chỉ liệt kê chỉ tiêu **đã thực hiện**, do `tag_context.seal_content_lines(results_data, kind)` dựng, mỗi dòng ≤ 3 chỉ tiêu cách nhau bằng tab:
- BSC/LAF: `Inflow | Downflow | HEPA filter leak` / `Lighting level | Noise level | Airflow pattern` / `UVC intensity | Particle count` (dòng 3 chỉ khi có `results_data["uvc"]` / `["particle"]`; chỉ tiêu có `{"skip": True}` bị bỏ).
- Tủ hút: `Face velocity` / `Lighting level | Noise level | Airflow pattern`.
- Ghi đè thủ công: `results_data["seal_lines"] = ["-A\t-B", "-C"]`.

`can_cu_tem` (căn cứ sau "Nội dung kiểm tra:", `tag_context.seal_standard`): BSC/LAF `EN 12469:2000`; **có đo nồng độ bụi thì thêm `ISO 14644-1:2015`** ở dòng dưới (`EN 12469:2000;` ↵ `ISO 14644-1:2015`, người dùng chốt 30/10/2026); tủ hút `ASHRAE 110, SEFA 1` (đã sửa lỗi chính tả `ASHREA`). Ghi đè: `results_data["seal_standard"]`. Cỡ chữ trên tem: nhãn 6pt; giá trị tên thiết bị, SN/ID (`serial/ mã thiết bị`), chủng loại và căn cứ 8pt (người dùng chốt 30/10/2026; bản TAG 26/10 để tên 10pt nên tên dài bị rớt dòng — đã sửa).

## 8. Excel tính toán (`TinhToanBSC / Fumehood / KhiNen _template.xlsx`)

Chỉ các ô thông tin hồ sơ là TAG (công thức, ảnh, định dạng giữ nguyên — AGENTS.md mục 11). Điền bằng `Core/Tools/tag_excel.py` ngay khi mở template; số liệu đo vẫn ghi vào đúng ô như trước.

| Template | Ô TAG |
|---|---|
| BSC – Bảng tính | C1 `ma_hs`, G1 `so_phu`, I1 `ngay_dk`, C2 `ngay_gd`, F2 `gio_bd`, I2 `gio_kt`, B3 `ten_tu`, G3 `nhan_hieu`, B4 `cap_tu`, G4 `model`, B5 `serial`, H5 `ma_tb` |
| BSC – TBKQ-VN | T11 `kh_ten`, T13 `kh_dia_chi`, T15 `noi_dung_yc`, U19 `dia_diem`, T33 `ngay_ct`, T40 `gdv_ten` |
| BSC – TBKQ - ENG | T9 `kh_ten_en`, T11 `kh_dia_chi_en`, T13 `noi_dung_yc_en`, U17 `dia_diem_en`, T41 `gdv_ten_en`, T57 `ten_tu_en` |
| Tủ hút – Bảng tính | như BSC (B4 = `loai_tu`) |
| Tủ hút – Thong Tin Chung Thu | C9/Q9 `kh_ten`/`kh_ten_en`, C11/Q11 địa chỉ, C13/Q13 nội dung, D17/Q17 địa điểm, C28 `ngay_ct`, C40 `gdv_ten`, Q56 `ten_tu_en`; AD18 (2 sheet EN) `dia_diem_en` |
| Khí nén | H2 `vi_tri_ten`, H3 `ky_hieu`, I6 `yc_bui`, I8 `yc_nuoc`, I9 `yc_dau` (mức yêu cầu của khách hàng — nhập theo từng hồ sơ: `point_data["requirements"] = {"dust","water","oil"}`) |
| Header trang in mọi sheet | `so_ct`, `ngay_ct` |

## 9. Chưa chuyển sang TAG

- Phần Phụ lục TBKQ **BSC / LAF / tủ hút** vẫn do code dựng (handler `bsc.py`, `laf_bsc.py`, `fumehood.py`); chỉ trang bìa và tem tủ dùng TAG.
- `TBKQ - viet/en - fumehood*.docx` (6 file mẫu tủ hút) chưa được code dùng, giữ nguyên.
- Các form biến thể `Templates/PhongSach/FORM - ...` (Y tế, Công nghiệp, Vi sinh, Scanloc) chưa gắn TAG.
- Script riêng hồ sơ 132115 (`generate_all_reports.py`) điền theo vị trí nên dùng form gốc trong `_goc_truoc_TAG/`.
