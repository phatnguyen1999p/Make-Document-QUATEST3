# Bản đồ quy trình N2.TC và ngưỡng đánh giá

> Nguồn: thư mục `QUY TRINH\` (15 file, đọc ngày 30/10/2026, Claude; cập nhật 30/10/2026 22:45 — thêm chốt Q5, UVC). Tài liệu tham chiếu cho Claude và Antigravity khi làm phiếu tính toán, BCKQ, TBKQ, tem tủ.
> Quy tắc: **ưu tiên mức do nhà sản xuất (NSX) công bố hoặc khách hàng (KH) cung cấp** (Excel `Data Do`, AGENTS mục 11). Ngưỡng dưới đây chỉ dùng khi không có mức đó. Quy trình nào không ghi mức thì không tự đặt mức.
> Ký hiệu nguồn: **[QT]** = ghi trong file quy trình; **[ẢNH]** = bảng dạng ảnh trong file quy trình; **[QCVN]** = PDF `01-btnmt-qc05.pdf`; **[CODE]** = chỉ có trong code `Core\`, quy trình không ghi.

---

## 1. Danh mục quy trình

| Mã | Tên quy trình | Nhóm (thư mục) | Hiệu lực | Đối tượng trong code (`dossier_type` / handler) | Ghi BBLM / TBKQ |
|---|---|---|---|---|---|
| N2.TC.BSC.01 | Đo đạc và đánh giá tủ an toàn sinh học cấp I | 4. Phòng sạch + Mapping + Tủ ATSH | 01/08/2025 | chưa có handler riêng | `QUATEST 3 3030:2025, N2.TC.BSC.01` |
| N2.TC.BSC.02 | Đo đạc và đánh giá tủ an toàn sinh học cấp II | 4 | 01/08/2025 | `bsc` / `BSCHandler`, `bsc_calc` | `QUATEST 3 3030:2025, N2.TC.BSC.02` |
| *(Dự thảo)* N2.TC.BSC | Đo đạc và đánh giá tủ ATSH cấp I & II | 4 | — | — | Nội dung **giống hệt BSC.02**, chỉ khác tiêu đề. Không dùng, dùng BSC.02 |
| N2.TC.LAF.01 | Đo đạc và đánh giá tủ sạch (LFC: LAF, clean bench, laminar flow hood/workstation; tủ cấy) | 4 | 01/08/2025 | `laf` / `LAFBSCHandler`; tủ cấy 132488/1 dùng `BSCHandler` + `inflow.skip`, `hepa.skip_exhaust` | `QUATEST 3 3030:2025, N2.TC.LAF.01` |
| N2.TC.T.01 | Đo đạc và đánh giá tủ hút khí độc | 4 | 01/08/2025 | `fumehood` / `FumeHoodHandler`, `fumehood_calc` | `QUATEST 3 3030:2025, N2.TC.T.01` |
| N2.TC.PS.01 | Đo một số thông số liên quan đến chất lượng phòng sạch | 4 | 01/08/2025 | `cleanroom` / `CleanroomHandler` (132115 dùng script riêng) | `QUATEST 3 3030:2025, N2.TC.PS.01` |
| N2.TC.MP.01 | Đo phân bố nhiệt độ, độ ẩm (mapping kho lạnh/kho lưu trữ) | 4 | 01/08/2025 | chưa có handler | `… N2.TC.MP.01` |
| N2.TC.KN.01 | Đo đạc và lấy mẫu khí có áp (tổng dầu, tiểu phân, điểm sương, vi sinh, thành phần khí) | 3. Khí (KKXQ) + Khí nén | 01/08/2025 | `compressed_air` / `CompressedAirHandler`, `khinen_calc`, `iso8573_cal` | `… N2.TC.KN.01` |
| N2.TC.K.01 | Đo các thông số khí tượng (nhiệt độ, độ ẩm, tốc độ gió) — KKXQ | 3 | 01/08/2025 | chưa có | `… N2.TC.K.01` |
| N2.TC.K.02 | Lấy mẫu, xác định bụi TSP và kim loại (Pb, As, Cd, Cr⁶⁺, Mn, Ni, Hg) — KKXQ | 3 | 01/08/2025 | chưa có | `… N2.TC.K.02` |
| N2.TC.K.03 | Lấy mẫu, xác định các yếu tố hóa học — KKXQ | 3 | 01/08/2025 | chưa có | `… N2.TC.K.03` |
| N2.TC.07 | Lấy mẫu, xác định yếu tố hóa học — môi trường lao động (MTLĐ), phương pháp hấp phụ | 3 | 29/12/2023 | chưa có | `… N2.TC.07` |
| N2.TC.08 | Lấy mẫu yếu tố hóa học MTLĐ bằng ống phát hiện nhanh (Kitagawa AP-20) | 3 | 29/12/2023 | chưa có | `… N2.TC.08` |
| N2.TC.09 | Lấy mẫu vi sinh vật trong môi trường (không khí, bề mặt) | 6. Vi sinh | 01/08/2025 | dùng chung với PS.01, KN.01 | `… N2.TC.09` |
| QCVN 05:2023/BTNMT | Quy chuẩn chất lượng không khí xung quanh (PDF) | 3 | — | — | căn cứ đánh giá K.02, K.03 |

Quy trình được viện dẫn nhưng **không có trong thư mục**: N2.TC.01 (vi khí hậu, ánh sáng nơi làm việc — PS.01 dùng cho nhiệt ẩm, độ sáng), N2.TC.02 (tiếng ồn nơi làm việc — PS.01 dùng cho độ ồn).

---

## 2. Tủ an toàn sinh học cấp II — N2.TC.BSC.02

Chỉ tiêu **bắt buộc**: inflow, downflow, hình thái dòng khí, rò rỉ HEPA. **Khuyến khích**: nồng độ hạt, độ sáng, độ ồn. Mô tả thêm: độ rung, UV. Tủ chạy ≥ 20 phút trước khi đo, UV tắt.

| Chỉ tiêu | Căn cứ | Ngưỡng [QT] | Cách đo / tính chính |
|---|---|---|---|
| Tốc độ dòng khí qua cửa làm việc (inflow) | EN 12469 G.3.2.2; NSF 49 6.10–6.11 | EN: **≥ 0,4 m/s**. NSF: A1 **> 0,38 m/s**, A2 **> 0,51 m/s**. Ưu tiên NSX | PP1 (ưu tiên) DIM TSI 8375/8380: 5 lần đo lưu lượng, lệch nhau ≤ 10 m³/h, V = Q/A. PP2 Kanomax (constricted window): điểm cách nhau ≤ 15 cm, mỗi điểm ≥ 15 s, V = V' × hệ số NSX (không có thì × H2/H1); biến thể NSF: 2 đường 25 % và 75 % độ mở, điểm cách mép ≥ 10 cm, cách nhau ≤ 10 cm. PP3 lưu lượng thải. PP2, PP3 chỉ dùng khi không làm được PP1, ghi lý do |
| Tốc độ dòng khí cấp qua HEPA (downflow) | EN 12469 G.3.2.1; NSF N-5.2 | Trung bình **0,25 – 0,50 m/s** và **mọi điểm trong Vtb ± 20 %**. Bảng 1 thêm theo NSX/NSF: trung bình ± 0,025 m/s; từng điểm ± 25 % hoặc ± 0,081 m/s | Lưới đo cao hơn mép dưới cửa kính 10 cm, cách vách 15 cm, điểm cách nhau 10–20 cm, mỗi điểm ≥ 15 s |
| Rò rỉ lọc HEPA (cấp, thải) | EN Annex D; NSF 5.22/6.3.2 | Scan trực tiếp màng lọc: **≤ 0,01 %** mọi điểm. Scan qua mặt nạ (không mở được / KH không cho): **≤ 0,005 %**. Lọc thải không tiếp cận được (khoan lỗ ống thải): **≤ 0,005 %** | Nồng độ phun > 10 µg/L, 23 psi, 2 hoặc 6 jet; C_upstream = 3375 × số jet / tổng lưu lượng (CFM) khi không đo trực tiếp được; scan 5 cm/s, đầu dò cách ≤ 2,5 cm. Kết quả 0 → ghi `< 0,0001 %` (AGENTS 8B) |
| Hình thái dòng khí | NSF 6.11 | Trong tủ, dọc tâm, cao 10 cm: khói đi xuống, không có điểm chết. Sau cửa kính 2,5 cm, cao 15 cm: đi xuống, không điểm chết, không thoát ra ngoài. Cách 5 cm dọc mép cửa: không thoát ra ngoài. Ngoài tủ, cách 4 cm dọc chu vi cửa: bị hút vào khe, không lan trên bề mặt làm việc | Smoke pencil |
| Cường độ sáng | EN A.2; NSF N-5.9.4 | EN: **> 750 lux**. NSF: **> 480 lux** + mức nền (thường 160 lux). Ưu tiên NSX | Điểm trên đường tâm, cách hai cạnh 15 cm, cách nhau ≤ 30 cm; lấy trung bình |
| Độ ồn | EN A.3; NSF N-5.11 | EN: **< 65 dBA**. NSF (field test): **< 70 dBA** | Máy cách mặt tủ 0,3 m, cao 0,38 m so với mặt tủ, giữa vùng mở cửa; đo nền và vận hành, mỗi lần 2–5 phút. Nền > 55 dBA → hiệu chỉnh theo bảng mục 9 |
| Nồng độ hạt bụi | ISO 14644-1:2015 | Cấp sạch NSX công bố (ISO 3 – 8) | TSI 9350, zero ≥ 250 L (0 hạt cả 6 kênh); đầu isokinetic giữa bề mặt làm việc, cao 30 cm; lấy 1 m³ |
| Cường độ UV-C | NSX | Theo NSX | Extech SDL470, giữa tủ, trung bình 1 phút |
| Độ rung | NSF N-5.10 | **≤ 0,05 mm** RMS, dải 10 – 1000 Hz (văn bản ghi "0,0005 m (0,05 mm)", xem mục 11) | RION VM-83/85 |

## 3. Tủ an toàn sinh học cấp I — N2.TC.BSC.01

Bắt buộc: inflow, rò rỉ HEPA, hình thái dòng khí. Khuyến khích: độ sáng, độ ồn.

| Chỉ tiêu | Ngưỡng [QT] | Ghi chú |
|---|---|---|
| Inflow (EN 12469) | **0,7 – 1,0 m/s** (văn bản: "> 0,7 m/s – 1,0 m/s") | Kanomax, ≥ 5 điểm: 4 góc cách cạnh 50–55 mm và 1 điểm giữa, mỗi điểm ≥ 15 s |
| Rò rỉ HEPA thải | Scan trên màng lọc ≤ 0,01 %; scan trên mặt nạ ≤ 0,005 % | Chỉ áp dụng tủ cấp I có lọc thải ra phòng |
| Hình thái dòng khí (NEBB) | Khói ở trong tủ, thoát lên theo ống thải; tại và ngoài vùng mở cửa khói bị hút vào, không cuốn vào vùng xoáy trần tủ | |
| Cường độ sáng (EN 12464-1, NSF 49) | **≥ 750 lux** | Rộng 0,9 m: ≥ 6 ô; 1,2 m: ≥ 8 ô; đo giữa ô và trên đường tâm |
| Độ ồn | **LAeq ≤ 65 dB** | Máy cách mặt tủ **1 m** (khác BSC.02 là 0,3 m); hiệu chỉnh nền như mục 9 |

## 4. Tủ sạch / LAF / tủ cấy — N2.TC.LAF.01

Bắt buộc: nồng độ hạt, downflow, rò rỉ HEPA. Khuyến khích: độ sáng, độ ồn. **Không có inflow, không có lọc thải.**

| Chỉ tiêu | Ngưỡng [QT] | Cách đo |
|---|---|---|
| Nồng độ hạt bụi (WHO-GMP Annex 6, ISO 14644-1) | Cấp sạch NSX công bố (tên bước 1); quy trình không ghi số cụ thể | Tủ sạch: 500 – 1000 L/điểm; LAF: cấp A, B 1000 L, cấp C, D 250 L; số điểm theo ISO 14644-1. Văn bản ghi đầu lấy mẫu "300 cm cách bề mặt" (xem mục 11) |
| Downflow | Trung bình **0,25 – 0,50 m/s** và mọi điểm trong **Vtb ± 20 %** | Tủ sạch: cách tấm phân phối 15 cm, cách vách 5 cm, điểm ≤ 10 cm. LAF: dưới HEPA 15 cm, cách mép 5 cm, điểm cách 20 cm. Mỗi điểm ≥ 15 s |
| Rò rỉ HEPA (NSF 49) | **≤ 0,01 %** mọi điểm | Upstream tính bằng công thức 3375 × jet / CFM, chế độ "set internal reference" |
| Cường độ sáng | **Theo công bố NSX** (không có số cố định) | Như BSC.02 |
| Độ ồn | **LAeq ≤ 65 dB** | Máy cách mặt tủ 1 m; hiệu chỉnh nền như mục 9 |

## 5. Tủ hút khí độc — N2.TC.T.01

Không áp dụng cho tủ hút có bộ lọc/hấp thu trước quạt. Tủ chạy ≥ 10 phút trước khi đo. Ưu tiên NSX; không có thì theo EN/ASHRAE/SEFA/TCVN.

| Chỉ tiêu | Căn cứ | Ngưỡng [QT] | Cách đo |
|---|---|---|---|
| Vận tốc dòng khí qua cửa làm việc (face velocity) | EN 14175-2 mục 8.1 | Theo quy định nước sở tại | — |
| | **SEFA 1-2010 mục 5.3.4** | **0,35 – 0,51 m/s** | ASHRAE 110: chia ô ≤ 0,09 m², cạnh ≤ 330 mm, đo giữa ô, ≥ 5 giá trị/điểm. Hoặc TCVN 6915: 3 vùng trái, giữa, phải |
| | TCVN 6914:2001 mục 5.2 | **0,35 – 0,61 m/s** (mục đo bằng flow hood ghi "0,35 0,60 m/s") | Flow hood TSI 8375: V = (Q/S)/3600 |
| Lưu lượng khí thoát | TCVN 6914:2001 mục 5.2 | **≥ 12 m³/phút** hoặc mức KH | TSI 8375 + hood, ≥ 5 lần, lấy trung bình |
| Cường độ sáng | SEFA 1-2010 mục 4.1.8 | **≥ 861 lux** | EN 14175-3 mục 9: ≥ 8 điểm/m² bề mặt thao tác, vẽ sơ đồ, lấy trung bình |
| Độ ồn | TCVN 6914:2001 | **≤ 60 dBA** | 3 điểm (trước, trái, phải), cách tủ 1 m, cao 1,2–1,5 m, mỗi điểm ≥ 3 phút; nền chỉ đo 1 điểm phía trước; lấy trung bình |
| | SEFA 1-2010 mục 12.1 | **≤ 85 dBA** | Tại vị trí tai người thao tác, 2–5 phút |

Quy trình T.01 **không** ghi bảng hiệu chỉnh ồn nền (BSC/LAF mới có).

## 6. Phòng sạch — N2.TC.PS.01

Mức đánh giá lấy theo cấp sạch/mức KH (Excel `Data Do`). Quy trình không đặt mức cho chênh áp, ACH, nhiệt ẩm, độ sáng, độ ồn.

| Chỉ tiêu | Ngưỡng / quy định [QT] | Cách đo / tính |
|---|---|---|
| Nồng độ hạt 0,1 – 5,0 µm | Mỗi vị trí không vượt giới hạn cấp sạch yêu cầu (ISO 14644-1 Bảng 1 hoặc WHO-GMP, mục 10) | Số điểm theo ISO 14644-1 Bảng A.1 (làm tròn lên). V_s = (20 / C_n,m) × 1000 L, ≥ 2 L, ≥ 1 phút/mẫu. Cao 1,2 m. Thiết bị: GMP C, D và ISO 7–9 → TSI 9350; GMP A, B → TSI 9350 + Lasair 110; ISO 1–6 → Lasair 110. Nhiều lần đo tại 1 vị trí → trung bình |
| Chênh áp | Theo mức KH | Đóng cửa; ống cách sàn ≥ 10 cm, tránh miệng cấp/hồi; đo 3 lần lấy trung bình |
| Lưu lượng gió cấp, ACH | Theo mức KH | Primary TSI 8375: mỗi miệng 3 lần, Q = Σ Q_itb. Secondary (chỉ khi miệng là HEPA): TSI 8345/Kanomax, cách lọc 15–30 cm, lưới ≥ 2 × 2, Q_i = v_itb × diện tích lọc. **ACH = ΣQ / thể tích phòng** |
| Rò rỉ HEPA | **Quy trình không ghi mức chấp nhận** (xem mục 12) | U_tb trong 10 – 100 µg/L, các lần đo lệch ≤ 15 % so với U_tb; scan 5 cm/s, cách bề mặt ≤ 3 cm |
| Vi sinh | Xem N2.TC.09 | |
| Nhiệt độ, độ ẩm | Theo mức KH | ≥ 1 điểm/vùng kiểm soát, ở độ cao bề mặt làm việc (N2.TC.01). Hiệu chính theo giấy hiệu chuẩn (AGENTS 8C) |
| Độ sáng | Theo mức KH | Tại vị trí làm việc (N2.TC.01) |
| Độ ồn | Theo mức KH | Cao 1,2 m, LAeq hoặc dải 1 octave (N2.TC.02) |

## 7. Mapping nhiệt độ, độ ẩm — N2.TC.MP.01

| Nội dung | Quy định [QT] |
|---|---|
| Ngưỡng | Dải nhiệt độ ghi trên nhãn dược phẩm / yêu cầu NSX (KH cung cấp) |
| Số đầu dò tối thiểu | < 2 m³: **10**; 2 – 20 m³: **16**; > 20 m³: **28** (hoặc theo KH, USP 1079.4) |
| Ghi dữ liệu | Logtag Haxo-8, 15 phút/lần, đặt 24 – 48 h |
| Kết quả | Hiệu chuẩn dữ liệu; trung bình, lớn nhất, nhỏ nhất; đồ thị |

## 8. Khí nén, khí có áp — N2.TC.KN.01

Quy trình **không ghi bảng cấp ISO 8573-1**; cấp được phân loại trong code (bảng dưới đánh dấu [CODE]).

| Nội dung | Quy định [QT] |
|---|---|
| Điều kiện an toàn | Áp tại điểm lấy ≤ 10 bar (lớn hơn → đề nghị KH giảm áp) |
| Tổng dầu (hạt + hơi) | Sau regulator < 0,5 bar, 1 L/phút. V_min = m / C (m = LOD 5 µg/mẫu). C yêu cầu 0,01 mg/m³ → **500 L / 500 phút**; 0,1 → 50 L / 50 phút; 1 → 5 L / 5 phút. KQTT = (KQTN dầu hạt + dầu hơi) / V_mẫu |
| Tiểu phân | Áp tại vị trí 5 – 7 bar. TSI 9350 qua HPD III, 55 – 60 L/phút; Lasair III 110 qua HPD II, 2 – 6 bar. **KQTT = KQ đo × (273 + T_khí) / 293** |
| Điểm sương | Xả 3 – 5 phút; đo 20 phút; độ dao động chấp nhận: > 0 °C → 2 °C (± 1,0); −40 – 0 °C → 4 °C (± 2,0); −80 – −40 °C → 6 °C (± 3,0). Trả trung bình cao nhất/thấp nhất; > 30 phút vẫn dao động → trả biểu đồ. Quy đổi về 7 bar, 20 °C theo ISO 8573-3 Annex B3 |
| Vi sinh | Áp ≤ 6 bar; Pinocchio 80 – 100 L/phút; 100 – 500 L. KQTT = CFU / V_mẫu (m³) |

ISO 8573-1:2010 [CODE `iso8573_cal.py`]:

| Cấp | Tổng dầu (mg/m³) | Điểm sương áp suất (°C) | Hạt 0,1–0,5 µm | 0,5–1 µm | 1–5 µm (hạt/m³) |
|---|---|---|---|---|---|
| 1 | ≤ 0,01 | ≤ −70 | ≤ 20 000 | ≤ 400 | ≤ 10 |
| 2 | ≤ 0,1 | ≤ −40 | ≤ 400 000 | ≤ 6 000 | ≤ 100 |
| 3 | ≤ 1 | ≤ −20 | — | ≤ 90 000 | ≤ 1 000 |
| 4 | ≤ 5 | ≤ +3 | — | — | ≤ 10 000 |
| 5 | — | ≤ +7 | — | — | ≤ 100 000 |
| 6 | — | ≤ +10 | — | — | — |

Có hạt > 5 µm thì ghi "Không phân loại".

## 9. Bảng hiệu chỉnh ồn nền [ẢNH trong BSC.01, BSC.02, LAF.01]

Chỉ áp dụng khi **ồn nền > 55 dBA**. ΔN = ồn vận hành − ồn nền.

| ΔN (dBA) | Trừ vào ồn vận hành |
|---|---|
| 0 – 2 | Không hiệu chỉnh — **phải giảm ồn nền** ("reduce background levels") |
| 3 | 3 |
| 4 – 5 | 2 |
| 6 – 10 | 1 |
| > 10 | 0 |

Số ồn ghi trên phiếu ghi chép là số đo **chưa hiệu chỉnh**. Bảng dùng số nguyên. **Chốt 30/10/2026:** ΔN làm tròn nửa lên về số nguyên rồi tra (2,6 → 3 → trừ 3; 5,4 → trừ 2; 5,5 → trừ 1). ΔN ≤ 2 → không hiệu chỉnh, code in `[WARNING] [Noise]`, báo người dùng. Code chung: `Core/Tools/noise_corr.py`.

## 10. Bảng tham chiếu cấp sạch và vi sinh

**ISO 14644-1:2015 Bảng 1** [ẢNH trong PS.01] — giới hạn hạt/m³, cộng dồn (≥ cỡ hạt):

| ISO | ≥ 0,1 µm | ≥ 0,2 µm | ≥ 0,3 µm | ≥ 0,5 µm | ≥ 1 µm | ≥ 5 µm |
|---|---|---|---|---|---|---|
| 1 | 10 | — | — | — | — | — |
| 2 | 100 | 24 | 10 | — | — | — |
| 3 | 1 000 | 237 | 102 | 35 | — | — |
| 4 | 10 000 | 2 370 | 1 020 | 352 | 83 | — |
| 5 | 100 000 | 23 700 | 10 200 | 3 520 | 832 | — |
| 6 | 1 000 000 | 237 000 | 102 000 | 35 200 | 8 320 | 293 |
| 7 | — | — | — | 352 000 | 83 200 | 2 930 |
| 8 | — | — | — | 3 520 000 | 832 000 | 29 300 |
| 9 (chỉ khi hoạt động) | — | — | — | 35 200 000 | 8 320 000 | 293 000 |

**WHO-GMP (TRS 961 Annex 6) Bảng 1** [ẢNH trong PS.01]:

| Cấp | Nghỉ ≥ 0,5 µm | Nghỉ ≥ 5,0 µm | Hoạt động ≥ 0,5 µm | Hoạt động ≥ 5,0 µm |
|---|---|---|---|---|
| A | 3 520 | 20 | 3 520 | 20 |
| B | 3 520 | 29 | 352 000 | 2 900 |
| C | 352 000 | 2 900 | 3 520 000 | 29 000 |
| D | 3 520 000 | 29 000 | không quy định | không quy định |

Cấp A lấy ≥ 1 m³/vị trí.

**ISO 14644-1 Bảng A.1 — số điểm đo tối thiểu** [ẢNH]: ≤ 2 m² → 1; 4 → 2; 6 → 3; 8 → 4; 10 → 5; 24 → 6; 28 → 7; 32 → 8; 36 → 9; 52 → 10; 56 → 11; 64 → 12; 68 → 13; 72 → 14; 76 → 15; 104 → 16; 108 → 17; 116 → 18; 148 → 19; 156 → 20; 192 → 21; 232 → 22; 276 → 23; 352 → 24; 436 → 25; 636 → 26; 1000 → 27; > 1000 → công thức A.1. Diện tích nằm giữa hai mốc thì lấy mốc lớn.

**Vi sinh không khí — N2.TC.09** [QT] (WHO-GMP Annex 6):

| Cấp | Giới hạn (CFU/m³) | Thể tích tối thiểu |
|---|---|---|
| A | < 1 | 1 000 L |
| B | ≤ 10 | 100 L |
| C | ≤ 100 | 10 L |
| D | ≤ 200 | 5 L |

N2.TC.09 khác: số vị trí không khí chủ động = N/3, bề mặt = 3 + N/3 (N = số điểm đếm hạt ISO 14644-1); nhà xưởng 60 – 100 L/phút trong 2 – 5 phút; đĩa thụ động ≤ 4 h, cách sàn 2 – 3 cm; cao lấy mẫu 1 – 1,2 m; bảo quản 2 – 8 °C, về PTN ≤ 24 h. Tính: chủ động KQTT (CFU/m³) = KQTN (sau hiệu chỉnh bảng Pr của NSX) × 1000 / V(L); thụ động KQTN < 1 → không phát hiện; tăm bông 100 cm², gạc ≥ 900 cm²; Rodac 25 cm², áp 10 s, KQTN < 1 → `< 1 CFU/25 cm²`.

## 11. Không khí xung quanh và môi trường lao động

**K.01 khí tượng**: Testo 625 (T, RH), Velocicalc 8345 (gió); cao 1,5 m; T/RH ổn định 3 – 5 phút, giữ 2 – 3 phút, cách người ≥ 0,5 m; gió đo 3 lần (20 s/lần, cách nhau ≤ 1 phút), lấy trung bình. Không có ngưỡng (QCVN 46:2012).

**K.02 TSP, kim loại** — lấy mẫu [QT]: TSP 500 L/phút × 60 phút (TCVN 5067:1995, giấy QC lệch 2 lần cân < 3 mg); As, Mn, Cd, Ni 500 L/phút × 1 h/24 h; Pb 24 h (NIOSH 7300, ICP-OES); Hg ống Carulite 200 mL/phút (NIOSH 6009).

**K.03 hóa chất KKXQ** — đánh giá theo QCVN 05:2023/BTNMT. Giá trị giới hạn [QCVN], đơn vị **µg/Nm³**:

| Thông số | 1 giờ | 8 giờ | 24 giờ | Năm | Lấy mẫu theo K.03 [QT] |
|---|---|---|---|---|---|
| SO₂ | 350 | — | 125 | 50 | MASA 704C, 0,2 L/phút, 30 phút |
| CO | 30 000 | 10 000 | — | — | PdCl₂, chai 1 L, 4–8 h |
| NO₂ | 200 | — | 100 | 40 | MASA 406, 0,4 L/phút, 30 phút |
| O₃ | 200 | 120 | — | — | UV photometric, đo trực tiếp |
| TSP | 300 | — | 200 | 100 | (K.02) |
| PM10 | — | — | 100 | 50 | |
| PM2,5 | — | — | 45 (từ 01/01/2026; trước đó 50) | 25 | |
| Pb | — | — | 1,5 | — | (K.02) |
| As | 0,03 | — | — | — | (K.02) |
| Cd | 0,4 | 0,2 | — | — | (K.02) |
| Cr⁶⁺ | 0,007 | — | 0,003 | — | |
| Mn (theo MnO₂) | 10 | — | 8 | — | (K.02) |
| Ni | — | — | 1 | — | (K.02) |
| Hg | — | — | 0,3 | — | (K.02) |
| HCl | — | — | 60 | — | NIOSH 7907, 2 L/phút, 30–60 phút |
| HNO₃ | 400 | — | 150 | — | NIOSH 7907 |
| H₂SO₄ | 300 | — | 50 | — | NIOSH 7908, 0,3 L/phút, 60 phút |
| HCN | 10 | — | — | — | NIOSH 7904, 1 L/phút, 60 phút |
| Benzene | 22 | — | — | — | NIOSH 1501, ORBO 32 |
| Toluene | 500 | — | — | — | NIOSH 1501 |
| Xylene | 1 000 | — | — | — | NIOSH 1501 |
| Styrene | — | — | 260 | — | NIOSH 1501, 1 L/phút, 24 h |
| Naphthalene | — | 500 | 120 | — | OSHA 35 / NIOSH S292 |
| Chloroform | — | — | 16 | — | NIOSH 1003 |
| Tetrachloethylene | — | — | 100 | — | NIOSH 1003 |
| Tổng hydrocarbon (CxHy) | 5 000 | — | 1 500 | — | NIOSH 1500/1501/1003 |
| Formaldehyde | 20 | — | — | — | NIOSH 3500, 1 L/phút, 60 phút |
| NH₃ | 200 | — | — | — | NIOSH 6015 |
| H₂S | 42 | — | — | — | MASA 701, 0,2 L/phút, 30 phút |
| n-heptane, n-hexane, n-octane, cyclohexane | QCVN 05:2023 **không có** giá trị riêng (chỉ tính vào CxHy) | | | | NIOSH 1500 |

**N2.TC.07 MTLĐ (hấp phụ)** và **N2.TC.08 (ống phát hiện nhanh)**: ngưỡng theo **QCVN 03:2019/BYT** (phụ lục từng chất ghi trong bảng của N2.TC.07; quy trình không chép giá trị). Lấy mẫu ngang tầm thở; mẫu thời điểm ca sáng và ca chiều mỗi ca ≥ 2 lần, mỗi lần ≥ 15 – 25 phút (cùng thời lượng); mẫu ca 8 h hoặc ≥ 80 % thời gian tiếp xúc (≥ 6,4 h). N2.TC.08 có danh mục ống Kitagawa kèm cột TLV (J/A/B — tham khảo, không thay QCVN 03).

---

## 12. Lệch giữa quy trình và code / AGENTS

**Người dùng chốt 30/10/2026 (đã sửa code):**
- Q1 → dùng bảng quy trình, ΔN làm tròn số nguyên, ΔN ≤ 2 cảnh báo (`noise_corr.py`, thay 3 cách cũ). Số ồn trên phiếu ghi chép là số **chưa hiệu chỉnh** → luôn áp bảng (132488 tủ 3: 63,5 → 60,5 dBA).
- Q2 → ± 20 %; **thông thường không xét độ đồng đều**, vẫn tính sẵn, bật bằng `downflow.eval_uniformity = True`.
- Q3 → `≥ 0,5 m/s` là mức công bố của một NSX (hồ sơ mẫu). Không có mức NSX thì dùng SEFA 0,35 – 0,51.
- Q4 → tạm thời ≤ 0,01 % cho mọi cách scan BSC/LAF.
- Q8 → phòng sạch mặc định ≤ 0,01 %; lọc H13 → ≤ 0,1 % (ISO 14644-3).
- Q5 → có catalogue NSX thì downflow = vận tốc công bố ± 20 %, độ sáng ≥ cường độ đèn công bố (132488/1 ESCO: 0,36 – 0,54 m/s; ≥ 1 110 lux); độ ồn NSX (buồng tiêu âm) không dùng, giữ ≤ 65 dBA. Không có NSX, KH không có mức → `no_eval`. TBKQ LAF đánh giá theo số liệu (`BSCHandler.result_rows`), không ghi cứng Đạt.
- UVC không có mức quy định → cột Đánh giá `-` (TBKQ, Excel `J122`), BCKQ xếp "Chỉ tiêu không đánh giá".

**Còn mở:** Q6 (dấu `>`/`≥` lux, `<`/`≤` ồn BSC.02), Q7 (ồn tủ hút TCVN ≤ 60 dBA vs SEFA ≤ 85 dBA, không có bảng hiệu chỉnh nền), Q9 (lưu lượng tủ hút ≥ 12 m³/phút chưa có trong code), `parse_fumehood_excel` mặc định `gte_05`.

Bảng gốc lúc rà soát (giữ để tra):

| # | Mức | Nội dung | Quy trình | Code / AGENTS hiện tại |
|---|---|---|---|---|
| Q1 | **Cao** | Hiệu chỉnh ồn nền, có **3 cách làm khác nhau** | Bảng mục 9: ΔN 0–2 → không hiệu chỉnh, phải giảm ồn nền; 3 → 3; 4–5 → 2; 6–10 → 1; > 10 → 0 | `bsc_calc.calc_noise`: ΔN < 3 → **trừ 3**, làm tròn 0,5 (3,5–5,49 → 2; 5,5–10,49 → 1). `fumehood_calc.calc_fumehood_noise` và `build_132488._noise_corr` (E22): ΔN < 3 → 0; khoảng (5; 6) → 0. Đề xuất: làm tròn ΔN về số nguyên rồi tra bảng; ΔN ≤ 2 thì báo người dùng (giảm ồn nền / đo lại / ghi chú). 132488 tủ 3 (ΔN = 2,6): nếu làm tròn thành 3 → trừ 3 dBA, khác kết quả đang ghi 63,5 |
| Q2 | **Cao** | Độ đồng đều downflow | BSC.02, LAF.01: mọi điểm trong **Vtb ± 20 %** | `bsc_calc.calc_downflow(tolerance_pct=10.0)` và ô Excel `E43` mặc định **10 %** → có thể đánh giá "Không đạt" mà theo quy trình là Đạt |
| Q3 | Trung bình | Vận tốc mặt tủ hút | T.01: SEFA **0,35 – 0,51**; TCVN **0,35 – 0,61** (mục flow hood ghi 0,35 – 0,60) | `fumehood_calc.calc_face_velocity(limit_type="gte_05")` và `FumeHoodHandler` mặc định **"≥ 0,5 m/s (*)"**, chú thích (*) ghi SEFA 1-2010 — mức ≥ 0,5 không có trong T.01 (132488 đã dùng 0,35 – 0,51) |
| Q4 | Trung bình | Rò rỉ HEPA khi scan qua mặt nạ / lọc thải không tiếp cận | BSC.01, BSC.02: **≤ 0,005 %** | Code chỉ có mức ≤ 0,01 %, không có trường ghi cách scan |
| Q5 | Trung bình | LAF/tủ cấy — mức độ sáng và downflow | LAF.01: độ sáng **theo NSX** (không số cố định); downflow **0,25 – 0,50 m/s ± 20 %** là mức chấp nhận của quy trình | `LAFBSCHandler` ghi cứng "≥ 750 lux ⁽¹⁾ EN 12469". AGENTS 8C (132488/1): KH không có mức thì `downflow`/`lux` dùng `no_eval` — trong khi LAF.01 có sẵn mức downflow |
| Q6 | Thấp | Dấu so sánh BSC cấp II | Bảng 1 BSC.02: lux **> 750**, ồn **< 65** | Code ghi **≥ 750**, **≤ 65,0** (BSC.01 và LAF.01 lại ghi ≥ 750, ≤ 65). Chỉ khác khi kết quả đúng bằng 750 / 65,0 |
| Q7 | Thấp | Ồn tủ hút | T.01 không có bảng hiệu chỉnh nền; TCVN ≤ 60 dBA (3 điểm, cách 1 m) khác SEFA ≤ 85 dBA (tai người thao tác) | Code luôn dùng ≤ 85 dBA và có hiệu chỉnh ΔN → cần ghi đúng vị trí đo tương ứng SEFA |
| Q8 | Thấp | Mức rò rỉ HEPA phòng sạch | PS.01 không ghi mức chấp nhận | Kiểm tra form phòng sạch đang ghi mức nào (thường ≤ 0,01 % theo ISO 14644-3 photometer) và xin người dùng chốt |
| Q9 | Thấp | Lưu lượng tủ hút ≥ 12 m³/phút (TCVN 6914) | Có trong T.01 | Chưa thấy trong `fumehood_calc` / `FumeHoodHandler` |

## 13. Lỗi chữ / mâu thuẫn trong file quy trình gốc (ghi nhận, không sửa)

- **BSC.02 độ rung:** "≤ 0,0005 m (0,05 mm)" — 0,05 mm = 0,00005 m. Thiết bị ghi lúc VM-83 (danh sách) lúc VM-85 (mục đo).
- **BSC.02, LAF.01 độ sáng:** mục đo ghi "Testo 540", danh sách thiết bị ghi Testo 440.
- **BSC.02:** NSF 49:**2022**; BSC.01, LAF.01, T.01: NSF 49:**2014**. BSC.01 ghi "ISO 12464-1:2011" (đúng là EN 12464-1).
- **LAF.01 nồng độ hạt:** "Vị trí đo: 300 cm cách bề mặt làm việc" — có thể là 30 cm (BSC.02 ghi 30 cm).
- **BSC.02 nồng độ hạt:** "bằng hoặc cao hơn cấp sạch do NSX công bố" — nên hiểu là đạt cấp sạch bằng hoặc tốt hơn mức NSX.
- **T.01:** viết "ASHREA" / "ASHARE" (đúng: ASHRAE); "0,35 0,60 m/s" thiếu dấu gạch và khác Bảng 1 (0,61).
- **KN.01:** tiêu đề "0,3 µm đến 0,5 µm" và "0,1 µm đến 0,5 µm" (có thể là đến 5,0 µm); còn viện dẫn biểu mẫu cũ "M01 của QTGĐ 07/N2-2022".
- **N2.TC.08:** biên bản viện dẫn "M02 - QTGD 07:2019" (cũ, nay là BM3030.02). N2.TC.07, N2.TC.08 header còn ghi "Tổng cục TCĐLCL", hiệu lực 29/12/2023.
- **PS.01:** viện dẫn N2.TC.01, N2.TC.02 nhưng hai file không có trong thư mục.
- **Dự thảo N2.TC.BSC:** trùng nội dung BSC.02.
