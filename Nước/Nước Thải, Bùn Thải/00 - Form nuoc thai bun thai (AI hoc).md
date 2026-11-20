# Form kết quả NƯỚC THẢI, NƯỚC DƯỚI ĐẤT, BÙN THẢI / CHẤT THẢI — AI học từ hồ sơ mẫu (16/11/2026)

Nguồn: `Nước\Nước Thải, Bùn Thải\` — 9 hồ sơ đã ban hành, đọc cả DOCX lẫn ảnh PDF. Phần chung với nước uống / sinh hoạt (bìa Table 0, bảng niêm phong, ảnh vị trí, cấu trúc BCKQ, `KPH` → `< LOD`, ⁽²⁾ giới hạn định lượng) xem `Nước\00 - Form ket qua nuoc (AI hoc).md`. Code `Core/` **chưa có** handler cho nhóm này.

| Hồ sơ | Loại | Căn cứ | Quy trình | Ghi chú |
|---|---|---|---|---|
| 100302 Nghị Phong | Nước dưới đất (giếng), 2 mẫu → `/1`, `/2` | `QCVN 09:2023/BTNMT – Quy chuẩn kỹ thuật quốc gia về chất lượng nước dưới đất` | `N2.TC.N.05` | 41 thông số, bảng chia nhóm |
| 106141 Cao ốc Khánh Hội | Nước thải sinh hoạt | `QCVN 14:2025/BTNMT - Quy chuẩn kỹ thuật quốc gia về nước thải sinh hoạt` (bảng 1, lưu lượng, cột B) | `N2.TC.N.06` | có niêm phong |
| 106222 Bosch | Nước thải KCN | Mức giới hạn KH cung cấp (hố xả ra Sonadezi) | `N2.TC.N.06` | 33 thông số |
| 106227 Việt Sơn | Nước thải đầu vào `/1` + đầu ra `/2` | `/1`: `-` (không mức); `/2`: `QCVN 40:2011/BTNMT – Quy chuẩn kỹ thuật quốc gia về nước thải công nghiệp`, cột A, Kq = 0,9, Kf = 1,0 | `N2.TC.N.06` | |
| 108037 JLL (Intel) | Nước thải vào KCNC | Tiêu chuẩn nước thải đầu vào NMXL tập trung Khu công nghệ cao (QĐ 257/QĐ-KCNC ngày 24/12/2020) | `N2.TC.N.06` | VN + EN |
| 109764 MT Kỷ Nguyên (LEGO) | Nước thải vào VSIP | Bảng nồng độ tối đa xả vào hệ thống thu gom VSIP (KH gửi kèm) | `N2.TC.N.06` | VN + EN |
| 107520 JLL (Intel) | Bùn thải, 3 ngày → `/1` 07/09, `/2` 08/09, `/3` 09/09 | `QCVN 07:2025/BNNMT - Quy chuẩn kỹ thuật quốc gia về ngưỡng chất thải nguy hại` | `N2.TC.B.01` | VN + EN, mỗi ngày 1 BCKQ |
| 110364, 110370 Phân viện KHHS | Chất thải rắn, **mẫu gửi** (4 / 5 mẫu) | `QCVN 07:2025/BNNMT` | `QUATEST3 3030:2025` (không mã N2) | mỗi mẫu 1 chứng thư, BCKQ gộp |

Nội dung yêu cầu: `Lấy mẫu và giám định chất lượng nước thải` / `… bùn thải`; nước dưới đất ghi `Lấy mẫu, thử nghiệm chất lượng nước ngầm`; mẫu gửi `Giám định mẫu chất thải`. EN: `Sampling and quality inspection of wastewater` / `… of sludge`.

## 1. Nước thải / nước dưới đất — phụ lục TBKQ

- **9.1**: như nước uống — bảng `Vị trí lấy mẫu | Ký hiệu mẫu | Chỉ tiêu thử nghiệm` (ô chỉ tiêu liệt kê tên thường + `(n thông số)`, vd `pH, BOD₅, COD, TSS, … (11 thông số)`), bảng niêm phong + ảnh túi mẫu, ảnh vị trí (chữ đỏ "VỊ TRÍ LẤY MẪU").
- **9.2 bảng kết quả** `STT | Thông số | Đơn vị | Phương pháp | Kết quả | Giá trị giới hạn cho phép ⁽¹⁾` (tên cột mức thay theo căn cứ: `Giới hạn tối đa`, `Giới hạn cho phép`, `Mức quy định`, `Giá trị giới hạn`). Nước thải dùng chữ **"thông số"** (108037 dùng "chỉ tiêu").
  - **Không có mức** (106227/1 đầu vào): bảng **5 cột** (bỏ cột mức), Căn cứ `-`, **không có Đánh giá**, Ghi chú chỉ ⁽²⁾ + VIMCERTS.
  - Nước dưới đất QCVN 09: dòng nhóm in nghiêng gộp ngang cả bảng — `Thông số cơ bản`, `Thông số ảnh hưởng sức khỏe con người`.
  - Nhóm chất (PCBs, thuốc BVTV clo / lân hữu cơ, DDT, xylene): 1 STT, tên nhóm rồi từng cấu tử `+ PCB-18`… trong cùng ô, kết quả từng dòng tương ứng, 1 mức chung.
  - Mức ghi theo tài liệu: `≤ 40`, `6 – 9`, `5,0 – 10`, `< 60`, `Không quy định`, `-` (không đánh giá). Coliform lớn ghi `1,3 × 10⁶`, `1 300`, `5 000` (cách nghìn).
  - Ô Phương pháp gộp dọc khi trùng; đo hiện trường `Đo tại hiện trường` (pH, nhiệt độ, TDS, clo dư); clo dư vượt thang ghi `> 4,4`.
- **Ghi chú**: `● Các thông số được chứng nhận VIMCERTS 078.` (EN `● Parameters are accredited to VIMCERTS 078.`) — **không có** câu VILAS 004 / VIAS 004 như nước uống; ⁽¹⁾ ghi nguồn mức, kèm `(thông tin cung cấp bởi đại diện <KH>)` khi KH chọn cột / hệ số; ⁽²⁾ giới hạn định lượng.
  - QCVN 40: `Giới hạn tối đa cho phép tính toán theo QCVN 40:2011/BTNMT; cột A; hệ số Kq = 0,9; Kf = 1,0` → mức = C × Kq × Kf (vd BOD₅ 30 × 0,9 = 27).
  - QCVN 14: `Giá trị giới hạn cho phép tại QCVN 14:2025/BTNMT, bảng 1, lưu lượng xả thải ≤ 2 000 m³/ ngày, cột B, nguồn nước tiếp nhận khác.`
- **Đánh giá**: `So sánh kết quả thử nghiệm các thông số của mẫu nước thải nêu trên với mức quy định tương ứng tại <căn cứ>:` → `+ Thông số không đánh giá` (khi có mức `-`) / `+ Thông số không đạt : sắt` / `+ Thông số đạt : các thông số còn lại`. EN: `Assessment:` / `Comparing the test result of parameters of wastewater sample mentioned above with the corresponding limits for …:` / `+ Non-compliant parameter : -` / `+ Compliant parameter : all parameters`.
- Thiết bị (mẫu cũ): `Máy đo pH Seven2Go S2, Mettler Toledo – USA`, `Máy đo TDS HI 98192, Hanna – USA`, `Máy đo pH HI98190, Hanna – USA`, `Máy đo Chlorine DR300, Hach – USA` (EN `Chlorine pocket colorimeter DR300, Hach – USA`, `pH meter Seven2Go S2, Mettler Toledo – USA`).
- Tên thông số EN (109764): Temperature, Colour, Biological oxygen demand (BOD₅) as O₂, Chemical oxygen demand (COD) as O₂, Suspended solids, Arsenic, Mercury, Lead, Cadmium, Chromium (VI)/(III), Copper, Zinc, Nickel, Manganese, Iron, Cyanide, Phenol, Mineral oil and fat, Animal-vegetable fat and oil, Residual Chlorine, Pesticides: Organic Phosphorous / Organic Chloride, Sulfide, Fluoride, Chloride, Ammonia (as N), Total nitrogen, Total phosphorous, Gross α / β activity; phương pháp hiện trường `On-site measurement`.

## 2. Bùn thải / chất thải (QCVN 07:2025/BNNMT) — phụ lục TBKQ

- **9.1 có lấy mẫu** (107520): bảng vị trí, ô chỉ tiêu 3 nhóm: `Tính chất nguy hại: tính kiềm/ tính axit (pH);` / `Thành phần nguy hại theo hàm lượng tuyệt đối: tổng xyanua, PCB;` / `Thành phần nguy hại theo nồng độ ngâm chiết: antimon, asen, …`. Ký hiệu mẫu theo KH (`NV4413.BT.1`).
- **9.1 mẫu gửi** (110364/110370): bảng `Tên mẫu theo khai báo` (`Chất thải rắn M1 – Vụ số: Tây Ninh`), `Số lượng mẫu`, `Ngày nhận mẫu`, `Người gửi mẫu`, `Mô tả tình trạng mẫu` (bao bì, niêm phong của ai, hình dạng màu sắc), `Chỉ tiêu yêu cầu thử nghiệm` + ảnh mẫu. Địa điểm = Trung tâm 49 Pasteur.
- **9.2 chia 3 mục**, mỗi mục bảng + Ghi chú riêng:
  - `9.2.1. Tính chất nguy hại` — `Tính chất nguy hại ⁽¹⁾ | Phương pháp | Kết quả | Ngưỡng CTNH ⁽²⁾`; pH `US EPA 9045D`, ngưỡng `pH ≥ 12,5 hoặc pH ≤ 2,0`. Ghi chú ⁽¹⁾ `Tính chất nguy hại được chứng nhận VIMCERTS 078.`, ⁽²⁾ `Ngưỡng chất thải nguy hại tại QCVN 07:2025/BNNMT, bảng 1.`
  - `9.2.2. Nồng độ ngâm chiết của các thành phần nguy hại` — `STT | Thành phần nguy hại ⁽¹⁾ | Phương pháp | Kết quả, mg/L | Ngưỡng CTNH ⁽³⁾, mg/L` (ngưỡng `≥ 15`, `≥ 5,0`…; Se `> 1`); phương pháp `US EPA 1311 & US EPA 200.7`… Ghi chú ⁽³⁾ `… tính theo nồng độ ngâm chiết tại QCVN 07:2025/BNNMT, bảng 2.`
  - `9.2.3. Hàm lượng tuyệt đối của các thành phần nguy hại` — `… | Kết quả, ppm | Ngưỡng CTNH ⁽⁴⁾, ppm` (tổng xyanua `≥ 590`, PCB `≥ 5,0`). Ghi chú ⁽⁴⁾ `… tính theo hàm lượng tuyệt đối tại QCVN 07:2025/BNNMT, bảng 3.` EN đơn vị `mg/kg`.
- **Đánh giá** (không có dấu `:` sau tiêu đề): `Độ pH không vượt ngưỡng nguy hại tương ứng nêu tại QCVN 07:2025/BNNMT.` rồi 2 khối `So sánh kết quả thử nghiệm nồng độ ngâm chiết / hàm lượng tuyệt đối …:` → `+ Chỉ tiêu cao hơn mức quy định : -` / `+ Chỉ tiêu thấp hơn mức quy định : Tất cả các chỉ tiêu` (dùng "cao hơn / thấp hơn", không "đạt / không đạt").
- **Phiếu tính toán** `tinh toan chat thai.xlsx` → PDF `KQ tinh toan.pdf`: header `<mã>` – ngày, tiêu đề `TÍNH TOÁN QUY ĐỔI HÀM LƯỢNG TUYỆT ĐỐI THEO MẪU RẮN KHÔ`, cột `Kí hiệu mẫu | Chỉ tiêu | kết quả, mg/kg | Độ ẩm, % | Hàm lượng chất rắn khô, % (=100 − độ ẩm) | Kết quả quy đổi về khối lượng khô, mg/kg`; giá trị `< LOD` giữ nguyên `<` (có số thì quy đổi = kết quả × 100 / % rắn khô — chưa thấy mẫu nào phát hiện); cuối `NBC: <tên>`. BCKQ viện dẫn `phiếu Tính toán hàm lượng tuyệt đối theo mẫu rắn khô số <mã> ngày …`.
- **BCKQ chất thải**: 4.6 `- KQTN hàm lượng tuyệt đối và nồng độ ngâm chiết của mẫu chất thải: xem tại phiếu KQTN số MTTN… ngày …;` + dòng phiếu tính toán; Đánh giá dùng bảng `Ký hiệu mẫu | Chỉ tiêu cao hơn mức quy định | Chỉ tiêu thấp hơn mức quy định` khi nhiều mẫu; mẫu gửi ghi thêm diễn biến (ký bổ sung giấy đăng ký, ngày gửi mẫu).

## 3. Đã chốt (người dùng 16/11/2026 — ghi AGENTS mục 3, 8C, 14)

| # | Điểm | Chốt |
|---|---|---|
| 1 | Câu công nhận | Theo **quy chuẩn** (chốt lại 16/11/2026): nước dùng `QCVN …/BYT` (uống, sinh hoạt, sạch, cấp sinh hoạt) → VIAS 004; **các mẫu nước còn lại** (thải, dưới đất, mặt…) và bùn / chất thải → `● Các thông số được chứng nhận VIMCERTS 078.` (không câu VIAS) |
| 2 | Thứ tự Ghi chú | VIMCERTS **luôn dòng đầu**, rồi ⁽¹⁾, ⁽²⁾ |
| 3 | Tên máy | Theo danh mục: `eq("hach_dr300")` `Thiết bị đo clo dư Hach DR300, USA`, `eq("hanna_hi98192")` `Thiết bị đo TDS Hanna HI 98192, USA`, `eq("hanna_hi98190")` `Thiết bị đo pH Hanna HI 98190, USA` |
| 4 | Đơn vị | Luôn `mg/L` |
| 5 | Từ gọi | Nước thải / nước dưới đất dùng "thông số" |
| 6 | Đề mục Đánh giá | Luôn `Đánh giá:` (có dấu hai chấm), cả bùn thải / chất thải |

`final-check` (`qa_check.py`, `is_vimcerts`) xét theo dòng "Nội dung yêu cầu" / "Căn cứ" của Table 0: sai câu → `[Missing Accreditation Note]` / `[Wrong Accreditation Note]`; VIMCERTS không ở dòng đầu → `[Accreditation Order]`. Câu VIAS viết gạch `–` vẫn được nhận.
Lưu ý khi dựng: các mẫu cũ ghi ký hiệu ghi chú dạng `(1) ...` — `final-check` coi `...` là chỗ trống chưa điền (`[Placeholder Found]`); mẫu cũ còn lỗi cỡ chữ 12,5pt, thiếu lặp tiêu đề bảng, đường kết thúc lệch → dựng mới theo AGENTS, không chép định dạng mẫu cũ.

Ghi chú kỹ thuật: dấu `≥` ở cột ngưỡng các mẫu cũ là ký tự font Symbol (`w:sym`, PDF trích ra ``) — khi dựng mới dùng ký tự Unicode `≥`.

Lỗi thấy trong mẫu (không chép theo): BCKQ 107520 ô Đơn vị ghi `Jones Lang 2aSalle (Việ1 Nam)`; `SWEWW 4110B` (đúng `SMEWW`); `Tổng Cyanua` / `Tổng xyanua` lẫn nhau; `Trichloroethylene (CH2Cl3)` ở hàng 1,1,1-trichloro… (QCVN 09) cần soát theo KQTN.
