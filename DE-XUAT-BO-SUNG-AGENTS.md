# Đề xuất bổ sung AGENTS.md — nguyên tắc rút từ hồ sơ đã làm còn thiếu

> Lập 06/11/2026 (Claude Code) để người dùng duyệt. Nguồn: `PROJECT_CONTEXT.md`, `ContextLuuTru/HoSo/*.md`, `ContextLuuTru/LoiDaSua.md`, docstring đầu các `build_<6 số>.py`, memory. Đối chiếu với AGENTS.md bản 06/11/2026 18:30.
> Mỗi điểm có mã `Nxx` để duyệt nhanh (vd "đưa N03, N07; bỏ N12; N05 sửa thành …").
> Cột **Tình trạng**: **Đã chốt** = người dùng đã quyết ở hồ sơ đó (chỉ cần xác nhận áp chung); **AI suy luận** = AI tự làm, chưa ai duyệt; **Cần quyết** = có lựa chọn.
> File này là bản nháp của AI, duyệt xong có thể xóa.

---

## A. Mâu thuẫn với AGENTS hiện tại — **đã chốt 06/11/2026, đã đưa vào AGENTS bản 18:30 – bản 2**

| Mã | Người dùng chốt | Đã làm |
|---|---|---|
| N01 | Ghi 1 trong 2 mã | 8C: cấp I `BSC.01`, cấp II `BSC.02`; code `inspection_procedure(text, bsc_class)`, `bsc_class_of`, `DEFAULT_PROCEDURE` |
| N02 | Tách BBLM trước; còn lại `Ghi Chep` + `Khac` (3 file) | 8A; `scan_splitter.py` thêm `Khac.pdf`, bỏ trang trắng |
| N03 | Có mã của KH thì tự ghi mã KH | 13 |
| N04 | Tự cập nhật thành ngày hiện tại (kể cả dựng lại) | 2; 6 build script không có KQTN dùng `date.today()` |
| N05 | Xung đột → người dùng quyết | 0: không tự áp tiền lệ |
| N06 | Bỏ chú thích ảnh hiện trường (sơ đồ giữ) | 3; `photo_grid(captions=False)` |

Bảng dưới giữ để đối chiếu.

| Mã | Vấn đề | AGENTS đang ghi | Thực tế hồ sơ | Đề xuất |
|---|---|---|---|---|
| N01 | **Quy trình thực hiện BSC** | 8C (chốt 02/11): BSC cấp I và cấp II đều ghi `N2.TC.BSC.01, N2.TC.BSC.02` | 03/11 (Antigravity, theo yêu cầu người dùng) 134804 tách: tủ cấp I `N2.TC.BSC.01`, tủ cấp II `N2.TC.BSC.02`; 06/11 136247, 136839 lại ghi cả hai; 134792 (tủ cấy) ghi `BSC.02` theo BBLM | Chốt một trong hai: (a) luôn ghi cả hai; (b) cấp I → `.01`, cấp II → `.02`. Sửa 8C + `DEFAULT_PROCEDURE` theo lựa chọn |
| N02 | **Edoc — trang nào vào `Ghi Chep.pdf`** | 8A: "mọi trang không có header rõ" vào Ghi Chep | 131671, 136839, 137058, 138143, 134804: scan gộp cả PKSCV, PYC, báo giá, PO, đề cương KH, hóa đơn lạ, trang trắng → AI tách tay, **không** đưa các trang đó vào Edoc | 8A sửa: BBLM = chỉ trang BM3030.02; Ghi Chep = BM3030.01 + phiếu thông số + phiếu in nhiệt / máy đo; **bỏ** trang trắng, PKSCV, PYC, báo giá, PO, đề cương, mặt bằng KH, trang không thuộc hồ sơ |
| N03 | **Ký hiệu mẫu** | 13: ghi đúng theo BBLM | 137058 (chốt 05/11): TBKQ ghi theo bảng vị trí của KH (`C1…C9`, `P1…P12`), BCKQ thêm cột "Ký hiệu mẫu (BBLM)" 1 → 21 để đối chiếu | Thêm ngoại lệ: KH có bảng vị trí/mã riêng → TBKQ dùng mã KH, BCKQ giữ cả số BBLM (cột đối chiếu) |
| N04 | **Ngày ban hành khi dựng lại hồ sơ đã ban hành** | 2: không có KQTN → ngày hiện tại | 132809 dựng lại 05/11 giữ ngày 04/11; 132488 dựng lại 01/11 người dùng chọn 01/11 | Thêm: dựng lại hồ sơ **đã có ngày ban hành** → giữ ngày cũ, chỉ đổi khi người dùng nói. "Ngày hiện tại" chỉ áp lần dựng đầu |
| N05 | **Nguồn ưu tiên khi lệch** | 0: lệch thông tin KH → báo người dùng quyết | Người dùng đã quyết theo hướng khác nhau: 138143 giữ BBLM (ngày, thiết bị, bỏ "ấp 3"); 137058 giữ PYC (địa chỉ + địa điểm); 134792 model theo BBLM, nội dung yêu cầu theo PYC; 132809 tên theo GCN | Giữ "báo người dùng", nhưng khi hỏi đưa sẵn **đề xuất mặc định**: tên + địa chỉ pháp lý → GCN > PYC; dữ liệu hiện trường (ngày, giờ, thiết bị, model/serial, vị trí, người thực hiện) → BBLM; nội dung yêu cầu, ngôn ngữ chứng thư → PYC. Cần quyết |
| N06 | **Chú thích dưới ảnh** | 3: chú thích ngay dưới ảnh, 13pt nghiêng | 132115 (phiên 14) người dùng chốt **bỏ** chú thích dưới ảnh hiện trường và sơ đồ | Hỏi: riêng 132115 hay áp chung phòng sạch? |

---

## B. Quy tắc mới đề xuất (theo mục AGENTS sẽ đặt vào)

> **Đã duyệt 07/11/2026** ("các nội dung còn lại tạm thời oke"), đã đưa vào AGENTS bản 07/11/2026 12:36. **N22 sửa:** tên EN theo GCN → tên quốc tế tra theo mã số thuế (`tax_helper.py --mst`) → không có mới tự dịch (bỏ bước "theo PO"). **N32** người dùng chốt 07/11/2026: tủ không đạt **vẫn cấp tem** — đã đưa vào AGENTS 8C.

### Mục 0 / 1 / 8A — đầu vào, quy trình làm việc

| Mã | Đề xuất | Nguồn | Tình trạng |
|---|---|---|---|
| N07 | Hồ sơ giao "trên Z" → chép **đầu vào** (scan, ảnh, Data Do, KQTN) từ Z về `Input\<6 số>\` trước khi làm; các file này chỉ đọc | 131671, 133266, 134792, 134804, 136839, 138143 | Đã làm thường xuyên |
| N08 | Hồ sơ 1 báo cáo: ảnh chèn chứng thư lấy từ `Photo\Use\` (hoặc `USE`, `Photos-…\Use`) do người dùng chọn; có `Use` thì chỉ dùng `Use`. Vi sinh bề mặt: ảnh `n.jpg` = mẫu ký hiệu n | 131671, 134804, 136839, 138143, 137058 | Đã chốt (131671 "hình ảnh dùng Photo/Use") |
| N09 | Scan gộp: liệt kê nội dung từng trang (PKSCV, PYC, báo giá, PO, BBLM, đề cương, mặt bằng, ghi chép, phiếu in, trang trắng) ở docstring script. Trang lạ không thuộc hồ sơ (vd hóa đơn YIWU trong 136247) → bỏ, nêu trong báo cáo cuối | 136247, 138143, 132809 | AI làm, chưa duyệt |
| N10 | **Đọc nhãn tủ trong ảnh** (model, serial, năm SX, setpoint, hệ số k) để đối chiếu phiếu/BBLM trước khi dựng | 134804, 136247, 136839 | AI làm, chưa duyệt |
| N11 | Người dùng **sửa tay trên Z** → ngoài việc chép về Input (mục 0), phải **đưa chỉnh sửa đó vào script** (để lần dựng sau không mất) và so chữ bản dựng lại với bản Z trước khi chép lên (138143: `_bckq_client_note` làm lại đúng bản Z) | 138143, 132115 | AI làm, chưa duyệt |
| N12 | Hồ sơ **lặp lại cùng khách** (Nextern, Cường Thịnh…): được lấy TBKQ cũ cùng khách trên Z làm mẫu cho diện tích, cấp sạch, sơ đồ vị trí, sơ đồ hệ thống khí nén; **ghi là suy luận**, báo người dùng; không lấy số liệu đo từ hồ sơ cũ | 133266 (108033, 122266), 138143 (88904) | AI làm, chưa duyệt |
| N13 | Docstring đầu `build_<6 số>.py` theo khuôn 4 khối: **Nguồn** (file + trang) → **Quyết định của người dùng (ngày)** → **Suy luận của AI (cần xác nhận)** → **Chạy** | mọi build script từ 01/11 | Đã làm thống nhất — bổ sung vào mục 10 |
| N14 | Sửa template → sửa ở bộ dựng `Core/Tools/build_tag_templates.py` (hoặc `_goc_truoc_TAG`), **không sửa tay** file TAG trong `Core/Templates/` (dựng lại template sẽ mất — E46c). Mục 15 hiện chỉ nói cho cỡ chữ | E46, E49 | Đề xuất áp chung |

### Mục 2 / 13 — ngày, KQTN, ngôn ngữ

| Mã | Đề xuất | Nguồn | Tình trạng |
|---|---|---|---|
| N15 | **Ngôn ngữ chứng thư theo PYC mục 15**: chỉ Tiếng Việt → chỉ TBKQ VN; chỉ Tiếng Anh → chỉ TBKQ EN (script có cờ `--vn` dựng thêm); Việt - Anh → cả hai. BCKQ, Excel luôn tiếng Việt. Quy tắc VN = EN (mục 9) chỉ áp khi có cả hai bản | 136247, 137058, 136839 | AI làm theo PYC; 137058 còn hỏi "có cần bản VN không" → cần chốt áp chung |
| N16 | "Ngày yêu cầu cấp KQTN" khi PYC để trống → **tạm** ngày lấy mẫu + 8 ngày; "KQTN có trễ hạn" tự tính theo ngày nhận KQTN so với ngày này | 133266 ("tạm thời tính +8 ngày") | Đã chốt tạm — xác nhận có áp chung không |
| N17 | Ghép mẫu KQTN ↔ vị trí khi phiếu KQTN chỉ đánh số (vd dầu hơi `MTTN….01` → KN1, `.02` → KN2, `.03` → blank): ghi rõ là suy luận, báo người dùng | 133266 | AI suy luận |

### Mục 3 / 15 — trình bày, ngắt trang

| Mã | Đề xuất | Nguồn | Tình trạng |
|---|---|---|---|
| N18 | Xóa mục khỏi template (vi sinh, HEPA…) → **xóa luôn đoạn trống đi kèm**; giữa câu Đánh giá và tiêu đề mục kế không có đoạn trống (tiêu đề `space_before` 6pt). Code mẫu `_drop_section_gaps` | 138143 (khoảng trắng 9.4 → 9.5), 132809 | Đã sửa theo báo lỗi của người dùng |
| N19 | Không để đoạn trống ngay trước đường kết thúc → tránh trang cuối chỉ có đường kết thúc (`_drop_trailing_empty`) | 138143, 134792 (VN 4 trang, EN 3) | AI sửa, chưa duyệt |
| N20 | Muốn bỏ trang thừa / lấp trang: chỉnh **kích thước ảnh, sơ đồ trong khung cho phép** (3 cột 4 – 5 cm, 2 cột 7 – 8 cm; sơ đồ 6,5 – 16 cm), không rút chữ. Vd 136247 ảnh 4,8 → 4,2 cm bỏ trang 3 trống; 138143 sơ đồ 13 → 15,8 cm, ảnh 5,0 cm cho 9.4 về trang 3 | 136247, 138143, 134792 | Đã chốt ở 138143 ("có thể tăng giảm kích thước layout và ảnh") |
| N21 | Tên cột bảng EN dài làm tăng số dòng → nới độ rộng cột (138143 "Post-Anesthesia Care Unit (PACU)" 1 dòng) hoặc rút ghi chú EN cho cùng số dòng với VN, để VN = EN ngắt cùng chỗ | 138143 | AI làm |

### Mục 6 / 14 / 15 — khách hàng

| Mã | Đề xuất | Nguồn | Tình trạng |
|---|---|---|---|
| N22 | **Tên EN khách hàng khi không có GCN**: theo PO / PYC bản EN; không có nguồn nào → AI dịch, nêu trong báo cáo cuối | 136839 (theo PO), 136247 (tự dịch) | AI suy luận — cần chốt |
| N23 | **Khách bổ sung PYC đổi tên** (vd liên danh): TBKQ VN/EN ghi tên mới; BCKQ + Excel giữ tên theo BBLM, ô "Đơn vị yêu cầu" BCKQ thêm dòng nghiêng `(Khách hàng bổ sung Phiếu yêu cầu & nhận thẩm định/ giám định với tên mới "…"; TBKQ ghi theo tên mới.)`, "Các bất thường" đánh `[X] Có` (theo bản người dùng sửa trên Z: không ghi ngày PYC bổ sung) | 138143 | Đã chốt (người dùng tự sửa trên Z) |

### Mục 8B — Excel tính toán

| Mã | Đề xuất | Nguồn | Tình trạng |
|---|---|---|---|
| N24 | Trước `strike_box`: **gỡ gộp, xóa giá trị và viền cả hàng** của vùng (C..M), không chỉ vùng gạch chéo — tránh sót số mẫu cạnh khung | E48 (134804) | Đã sửa trong `bsc_calc`; ghi để bộ dựng mới làm theo |
| N25 | Sửa Excel bằng openpyxl xong → **tính lại bằng Excel COM** trước khi đọc giá trị hoặc xuất PDF (openpyxl xóa giá trị đã tính); openpyxl làm mất ảnh "trong ô" (richData) → chèn ảnh nổi theo khung | 134804, E46a | Kỹ thuật — đề xuất ghi vào mục 11 |

### Mục 8C — tủ (BSC, tủ cấy)

| Mã | Đề xuất | Nguồn | Tình trạng |
|---|---|---|---|
| N26 | **Hệ số k inflow (Kanomax, PP2)** lấy theo nhãn NSX trên tủ (ESCO: "Multiply the SM average with correction factor of 0.457"; 136247 k = 0,435); số ghi tay trên phiếu chỉ đối chiếu (134804 phiếu 0,41 → dùng 0,49) | 134804, 136247 | Đã chốt ở 134804 |
| N27 | **HEPA ATI 2i quét nhiều lần**: mỗi lọc lấy lần quét **lớn nhất** (khớp phiếu ghi chép); phân lọc cấp / thải theo giờ quét + phiếu ghi chép; phiếu in "Alarm Exceeded: Yes" là dấu hiệu Không đạt, đối chiếu lại | 134792, 136839, 136247 | AI suy luận — cần chốt "lấy max" |
| N28 | **KH yêu cầu không thể hiện đánh giá** (BBLM ghi) → `has_assessment=False`: bảng 3 cột STT / Chỉ tiêu / Kết quả, bỏ cột mức + đánh giá, Ghi chú chỉ câu công nhận; thêm dòng vào bảng 8C | 134792 | Đã chốt |
| N29 | Chỉ tiêu vi sinh tủ cấy: STT riêng "Thử khả năng lọc vô trùng của tủ", dòng con "Tổng số vi sinh vật hiếu khí", VT1 – VTn theo sơ đồ trên phiếu + sơ đồ sau bảng Ghi chú; kết quả chờ KQTN (`………` vàng) | 134792 | Đã chốt |
| N30 | Nội dung yêu cầu (mục 3) theo **PYC/báo giá** khi BBLM/phiếu in sẵn loại tủ khác (134792: PYC "tủ cấy", BBLM in sẵn "tủ an toàn sinh học") | 134792 | Đã chốt |
| N31 | Kết quả **sát ngưỡng** (vd inflow 0,402 → in 0,40, ngưỡng ≥ 0,40): đánh giá theo giá trị in trên chứng thư, nêu trong báo cáo cuối | 136839 | Cần quyết: đánh giá theo số làm tròn hay số chưa làm tròn |
| N32 | **Tủ không đạt — có cấp tem không?** (136247 đã dựng tem cho tủ không đạt ồn + lọc thải) | 136247 | Cần quyết |
| N33 | Tem: serial `N/A` mà có mã thiết bị → SN/ID chỉ ghi mã thiết bị | 132488 (30/10) | Đã làm trong `seal_serial` — bổ sung câu vào 8C tem |

### Mục 8C — phòng sạch

| Mã | Đề xuất | Nguồn | Tình trạng |
|---|---|---|---|
| N34 | **Hiệu chính nhiệt độ, độ ẩm theo giấy hiệu chuẩn**: sai số = **nội suy tuyến tính** giữa 2 điểm hiệu chuẩn kẹp hai bên; ngoài khoảng → điểm biên (mục 15 đã có); giá trị = TB 3 lần sau hiệu chính. Excel: dòng đọc + dòng sau hiệu chính (`=B-FORECAST(...)`) + bảng hiệu chuẩn trên sheet; BCKQ viện dẫn phiếu "Tính toán hiệu chuẩn nhiệt độ, độ ẩm…". Thay câu 8C "132115: −0,1 °C / −1,9 %RH" (hằng số) | 138143 (chốt 04/11), 132809 | Đã chốt |
| N35 | Lấy bảng hiệu chuẩn từ thư viện `HieuChuan/` (`Core/Tools/calib_library.py`, `cal_table()`), không gõ tay; **cảnh báo khi thiết bị dùng đo có giấy hết hạn / không đạt** ở ngày đo (thư viện 05/11 phát hiện Testo 440 MT38-1 không đạt kiểm định, nhiều giấy hết hạn) | thư viện 05/11 | Đề xuất — chưa nối vào build |
| N36 | **ACH** = tổng lưu lượng HEPA / thể tích phòng; thể tích = diện tích × chiều cao theo giấy tính / BBLM; giấy tính tay ghi số tròn ("ACH: 20") → tính lại (21,2). Lưu lượng testo 420: một HEPA đo nhiều phần miệng cấp (ghi chú `1-1 … 3-3`) → cộng lại | 138143 (chốt thể tích 65 m³) | Đã chốt |
| N37 | **File máy đếm hạt**: AeroTrak `6ChannelFixed.xlsx` (hạt/m³ cộng dồn Σ), Lasair `database*.csv` (Data ID = số thứ tự dòng); file **zero** đầu loạt không tính; vị trí ↔ số file theo phiếu ghi chép; file không có trên phiếu → không dùng (133266 file 1562; 132809 file bụi sáng 29/10); phiếu ghi thứ tự lạ → theo giờ đo | 131671, 132809, 133266, 138143 (chốt theo giờ đo) | Đã chốt một phần |
| N38 | **Chỉ tiêu có mức nhưng không đánh giá do điều kiện hiện trường** (KH xác nhận có nguồn ồn khác): TBKQ chỉ ghi số, mức `-`, ghi chú ⁽ⁿ⁾ nêu lý do; câu Đánh giá chỉ nêu chỉ tiêu còn lại; Căn cứ bỏ chỉ tiêu đó; BCKQ "Chỉ tiêu không đánh giá: … — lý do"; Excel mức/đánh giá `-` | 138143 | Đã chốt |
| N39 | Mức lấy từ **đề cương / tài liệu KH trong scan** (vd đề cương BV Hoàn Mỹ) giống mức trong Excel người dùng: ghi "Mức quy định cung cấp bởi khách hàng", nêu tài liệu ở docstring | 138143 | AI làm |
| N40 | **Vi sinh không khí phòng sạch**: mức theo KH (131671 ≤ 200 CFU/m³ ghi trên phiếu), **không** ghi cứng "WHO TRS 961"; có câu Đánh giá vi sinh; blank ghi ở BCKQ | 131671, E44 | AI làm — template Core chưa sửa |
| N41 | Chênh áp: tên khu vực đối chiếu theo phiếu (không tự ghi "Bên ngoài"); giá trị = TB các lần đọc | 138143 ("Hành lang") | Đã chốt |
| N42 | Định dạng phiếu tính toán bụi: số đếm `#,##0`, in đậm cỡ hạt có giới hạn theo cấp sạch, hiện mọi cột cỡ hạt, cố định dòng tiêu đề, lặp tiêu đề khi in | 132115 (`format_excel_132115.py`) | Mới áp 132115 — hỏi áp chung |

### Mục 8C — loại đối tượng chưa có dòng trong bảng

| Mã | Đề xuất | Nguồn | Tình trạng |
|---|---|---|---|
| N43 | **Khí nén** (thêm dòng bảng 8C): `CompressedAirHandler` + `khinen_calc` (mẫu `build_133266.py`); quy trình `QUATEST 3 3030:2025, N2.TC.KN.01`; căn cứ `ISO 8573-1:2010` + ngưỡng KH / ngưỡng nội bộ ghi trên phiếu ghi chép; điểm sương theo ảnh màn hình CS DP510 đầu/cuối (`KQDo\KNn\Use`); hạt theo thể tích / thời gian mẫu KH yêu cầu (nhãn ghi tay trên phiếu in); dầu theo KQTN, KPH → `< LOD`; mục Đánh giá chỉ có khi có mức | 133266, 122266 | Đã chốt một phần (02/11) |
| N44 | **Vi sinh bề mặt (swab)** (thêm dòng bảng 8C): bìa qua `CleanroomHandler` + phụ lục dựng riêng (mẫu `build_137058.py`); quy trình `QUATEST 3 3030:2025, N2.TC.09`; **không** hiệu chỉnh Feller, **không** phiếu tính toán Excel; kết quả CFU/100 cm² ghi như KQTN; tổng số VSV hiếu khí, **nấm men, nấm mốc tách cột riêng**; ký hiệu theo N03; lưới ảnh 3 cột | 137058 | Đã chốt (05/11) |

---

## C. Đã có quy tắc nhưng Core chưa làm (script hồ sơ đang vá tay)

Không phải quy tắc mới — ghi để biết các script sau còn phải vá:

- In đậm giá trị **Không đạt** (mục 3): 136247, 138143 tự vá (`_bold_fail_values`) — E41.
- Tiêu đề mục kết quả phòng sạch keepNext, giữ trọn bảng nhỏ (mục 15): 131671, 132809, 133266, 138143 tự vá — E44.
- Footer phụ lục `{= {NUMPAGES} - 1}` thay `SECTIONPAGES`: 132809, 133266, 137058 tự vá — E46e.
- Lề ô `tcMar` làm dòng bảng cao hơn ~4,5 pt: 137058 tự gỡ — E51.
- Câu WHO TRS 961 ghi cứng ở bảng vi sinh template phòng sạch — E44 (N40).
