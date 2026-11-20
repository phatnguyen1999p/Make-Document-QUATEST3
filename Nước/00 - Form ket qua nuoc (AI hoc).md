# Form kết quả giám định NƯỚC — ghi chép AI học từ hồ sơ mẫu (16/11/2026)

Nguồn: `Nước\` — 3 hồ sơ đã ban hành có PDF (113750 NU, 114119 NSH + NU mẫu gửi, 116644 NU) và mẫu trong các hồ sơ Hyosung đang làm (bản chép từ 133022 Thuận Hải, 137037 / 090965 Tân Hiệp). Đã đọc cả DOCX lẫn ảnh PDF (bìa có giấy tiêu đề, phụ lục, BCKQ scan, BBLM scan).
Code `Core/` **chưa có** handler cho nước — hồ sơ nước hiện làm tay trên DOCX mẫu.

## 1. Ba dạng hồ sơ

| Dạng | Căn cứ | Quy trình | Nội dung yêu cầu (mục 3) | Chỉ tiêu thường gặp |
|---|---|---|---|---|
| **NU** — nước uống (máy lọc, bình) | `QCVN 06-1:2010/BYT - Quy chuẩn kỹ thuật quốc gia đối với nước khoáng thiên nhiên và nước uống đóng chai` | `N2.TC.N.07` | `Lấy mẫu và giám định chất lượng nước uống` | 21 hóa lý (phụ lục II) + 05 vi sinh (phụ lục III); hoặc chỉ 05 vi sinh (phụ lục III); hoặc gói riêng KH chọn (116644: 15 hóa lý + 06 vi sinh) |
| **NSH** — nước sinh hoạt (vòi, bồn, bể) | `QCVN 01-1:2024/BYT - Quy chuẩn kỹ thuật quốc gia về chất lượng nước sạch sử dụng cho mục đích sinh hoạt` | `N2.TC.N.01` (BBLM 145067 ghi `N2.TC.N.02` — cần hỏi) | `Lấy mẫu và giám định chất lượng nước sinh hoạt` | 10 thông số nhóm A (+ Pseudomonas aeruginosa, Staphylococcus aureus nếu KH yêu cầu); đo hiện trường pH, nhiệt độ, clo dư, mùi |
| **Mẫu gửi** — KH tự mang mẫu (114119) | theo loại nước | dòng `Cơ sở tiến hành giám định: QUATEST3 3030:2025` thay dòng Quy trình | `Giám định mẫu nước` | như trên; Địa điểm = `Trung tâm Kỹ thuật Tiêu chuẩn Đo lường Chất lượng 3, 49 Pasteur, phường Sài Gòn, Thành phố Hồ Chí Minh` |

- **Nhiều mẫu / nhiều loại nước:** mỗi chứng thư một mẫu (hoặc một loại nước) — `/1`, `/2`, … (133022: N1–N6 NU → `/1`–`/6`, N7–N9 NSH → `/7`…). **BCKQ gộp 1 file**, số = mã trần.
- **Ký hiệu mẫu:** `<6 số> - NU`, `<6 số> - NSH`, hoặc `N1…Nn` (BBLM ghi `145067 – N1`) — theo BBLM (AGENTS mục 13).
- Chỉ có bản **tiếng Việt** trong các mẫu (chưa thấy TBKQ EN cho nước).

## 2. TBKQ (chứng thư) — trang bìa (Table 0, 13pt, có giấy tiêu đề)

```
1. Đơn vị yêu cầu giám định : <tên KH>
   Địa chỉ                  : <địa chỉ 2 cấp>
2. Văn bản yêu cầu giám định :
-  Giấy yêu cầu và nhận giám định số:  Q326Dxxxxxx  ngày  dd/mm/yyyy
3. Nội dung yêu cầu giám định : Lấy mẫu và giám định chất lượng nước uống | nước sinh hoạt
4. Thời gian giám định : Từ dd/mm/yyyy          (133022 ghi "Ngày dd/mm/yyyy")
5. Địa điểm giám định  : <nơi lấy mẫu theo BBLM>, <địa chỉ>
6. Quy trình thực hiện : QUATEST3 3030:2025, N2.TC.N.07 | N2.TC.N.01
7. Thiết bị sử dụng    : -   (không đo hiện trường)
                         - Máy đo pH Seven2Go S2, Mettler Toledo – USA;
                         - Máy đo Chlorine DR900, Hach – USA.
8. Căn cứ áp dụng/ đánh giá : <QCVN + tên đầy đủ>
9. Kết quả giám định : Xem Phụ lục đính kèm
GIÁM ĐỊNH VIÊN (dòng 2)          KT. GIÁM ĐỐC / PHÓ GIÁM ĐỐC
Nguyễn Xuân Nam | Phạm Quốc Nghĩa | Nguyễn Khánh Toàn      Lê Đăng Huyền
```
Trang A4, lề trái 3,0 / phải 1,5 / trên 1,2 / dưới 1,0 cm. Header phụ lục: `Phụ lục / CHỨNG THƯ GIÁM ĐỊNH / Số: <số> - Ngày dd/mm/yyyy`; footer `Trang x/y` (đếm riêng phụ lục).

## 3. TBKQ — phụ lục

**9.1 Nội dung thực hiện** (có lấy mẫu):
1. `- Vị trí lấy mẫu và chỉ tiêu thử nghiệm theo yêu cầu của <tên KH>, cụ thể như sau:` → bảng 3 cột `Vị trí lấy mẫu | Ký hiệu mẫu | Chỉ tiêu thử nghiệm` (tiêu đề nền xám, căn giữa; ô chỉ tiêu ghi gộp `21 chỉ tiêu hóa lý và 05 chỉ tiêu vi sinh tại QCVN 06-1:2010/BYT, phụ lục II và phụ lục III`, hoặc liệt kê `- 15 chỉ tiêu hóa lý: …` / `- 06 chỉ tiêu vi sinh: …`).
2. `- Thông tin về quá trình lấy mẫu và hình ảnh mẫu niêm phong:` → bảng `Người chứng kiến | Ký hiệu mẫu | Thời điểm lấy mẫu | Số niêm phong` (vd `Đại diện <KH> – anh <tên>` | `116644 - NU` | `18:25` | `KT371440`) + ảnh túi mẫu niêm phong. **Không niêm phong → ghi `: -` cuối dòng và bỏ bảng** (PDF 113750 đã bỏ bảng; DOCX còn sót bảng dòng `-`).
3. `- Hình ảnh vị trí lấy mẫu:` → **1 ảnh căn giữa** (~9 cm), ảnh có thể có chữ đỏ "VỊ TRÍ LẤY MẪU" người dùng đã đánh dấu sẵn (`0. Vị trí lấy mẫu.jpg`).

**9.1 dạng mẫu gửi** (114119): bảng thông tin 3 cột có gạch đầu dòng ▪ — `Tên mẫu theo khai báo`, `Số lượng mẫu`, `Ngày nhận mẫu`, `Người gửi mẫu`, `Mô tả tình trạng mẫu` (vd `Chất lỏng trong, không cặn chứa trong bình nhựa 5L`), `Chỉ tiêu yêu cầu thử nghiệm` → ảnh chai mẫu.

**9.2 Kết quả giám định** — bảng 6 cột, 13pt, tiêu đề nền xám + lặp tiêu đề:

| | NU (QCVN 06-1) | NSH (QCVN 01-1) |
|---|---|---|
| Tiêu đề | `Stt | Chỉ tiêu | Đơn vị | Phương pháp | Kết quả | Giới hạn tối đa ⁽¹⁾` | `STT | Tên thông số | Đơn vị | Phương pháp | Kết quả | Ngưỡng giới hạn cho phép ⁽¹⁾` |
| Vi sinh | `CFU/ 250 mL` (bào tử kị khí `CFU/ 50 mL`), mức `Không phát hiện` | `CFU/ 100 mL`, mức `< 1` |
| Clo | `Clor (Cl₂)` mức 5 | `Clo dư tự do` mức `Trong khoảng (0,2 – 1,0)`; pH `Trong khoảng (6,0 – 8,5)` |

- **Gộp dọc ô Phương pháp** khi các dòng liền nhau cùng phương pháp (`US EPA 200.8`, `SMEWW 4110B:2023`, `ISO 9308 – 1:2014 / Amd.1:2016`); gộp dọc ô mức `-` của các dòng liền nhau không có mức.
- Ký hiệu hóa học dùng chỉ số dưới / trên (`Cl₂`, `NO₃⁻`, `NH₄⁺`).
- Chỉ tiêu đo hiện trường: Phương pháp `Đo tại hiện trường`; Mùi: `Cảm quan tại hiện trường`, kết quả `Không có mùi lạ`.
- **KQTN ghi `KPH` → `< <giới hạn phát hiện>`** lấy ở cột "Giới hạn phát hiện" của phiếu KQTN (vd độ màu `< 5`, độ đục `< 0,5`, As `< 0,0005`); khi số là **giới hạn định lượng** → thêm ⁽²⁾ sau số (`< 0,4 ⁽²⁾`) và ghi chú (2). Vi sinh `＜1⁽ᶜ⁾` → `< 1`.
- Chỉ tiêu không có mức trong QCVN (KH chọn thêm: độ đục, màu, Zn, Fe, TDS, TSVSVHK…) → mức `-` và vào "Chỉ tiêu không đánh giá".

**Ghi chú** (11pt nghiêng, gạch chân nhãn):
```
⁽¹⁾ ... Mức giới hạn tối đa cho phép quy định tại phụ lục II và phụ lục III, QCVN 6-1:2010/BYT.   (chỉ vi sinh: "tại phụ lục III, …")
        | Ngưỡng giới hạn tối đa cho phép quy định tại QCVN 01-1:2024/BYT.
⁽²⁾ ... Giới hạn định lượng của phương pháp thử.        (chỉ khi có)
 ●  Phương pháp thử nghiệm được công nhận ISO 17025 – VILAS 004.
 ●  Lĩnh vực thực hiện được công nhận ISO 17020 – VIAS 004.     (đứng cuối)
```

**Đánh giá** (13pt; cụm từ in đậm nghiêng):
```
Đánh giá:
So sánh kết quả thử nghiệm các chỉ tiêu của mẫu nước uống nêu trên so với mức quy định tương ứng tại QCVN 06-1:2010/BYT, phụ lục II và phụ lục III:
   + Chỉ tiêu ***không đánh giá*** : độ đục, màu sắc, …        (chỉ khi có)
   + Chỉ tiêu ***không đạt***      : -
   + Chỉ tiêu ***đạt***            : Tất cả các chỉ tiêu | các chỉ tiêu còn lại
```
NSH: thay "chỉ tiêu" bằng **"thông số"**, căn cứ `QCVN 01-1:2024/BYT`. Đường kết thúc đôi ngay dưới câu cuối.

## 4. BCKQ (BM3030.03, nội bộ, 1 file cho cả hồ sơ)

- Bảng đầu: `Số: <mã trần>`; 1 Đơn vị; 2 Địa chỉ; 3 Nội dung (`Lấy mẫu, giám định chất lượng nước uống [và nước sinh hoạt]`); 4 Quá trình: 4.1 Địa điểm, 4.2 Thời gian `Ngày dd/mm/yyyy từ hh:mm đến hh:mm` (theo BBLM), 4.3 Thiết bị, 4.4 Người thực hiện (tên cuối).
- 4.5 `- Vị trí lấy mẫu và chỉ tiêu thử nghiệm được chỉ định bởi đại diện <KH>, cụ thể như sau:` → bảng `Vị trí thực hiện | Kí hiệu mẫu | Chỉ tiêu thử nghiệm` (đủ mọi mẫu); bảng niêm phong; ảnh. Mẫu gửi: `Số lượng mẫu`, `Ngày nhận mẫu`, `Người gửi mẫu`, bảng `Ký hiệu mẫu | Tên mẫu theo khai báo | Mô tả tình trạng mẫu | Chỉ tiêu`, `Hình ảnh mẫu gửi:`.
- 4.6 Kết quả đo đạc/ phân tích: `- KQTN mẫu <ký hiệu> : xem tại phiếu KQTN số MTTN…` (nhiều mẫu: `N1 ÷ N6 : xem tại phiếu VSTN…`, `MTTN….01-03`; mẫu QC `MTTN….04`; có thể thêm nhận xét in nghiêng như `(Kết quả Amoni < 0,1 mg/L phù hợp với bản chất nước khử ion)`); `- Kết quả đo tại hiện trường:` bảng `Kí hiệu mẫu | pH | Nhiệt độ [°C] | Clo dư [mg/L] | Mùi` (+ dòng `Thiết bị sử dụng | Mettler Toledo S2 | Hach DR900 | -`), không đo thì `: -`.
- Đánh giá: bảng `Ký hiệu mẫu | Chỉ tiêu đạt | Chỉ tiêu không đạt` hoặc 3 dòng `+` như TBKQ.
- Gửi mẫu thử nghiệm — mỗi phiếu KQTN một khối (`a.`, `b.`): `Phiếu KQTN số … ngày …`, Ngày gửi mẫu, Ngày yêu cầu cấp KQTN, Ngày nhận KQTN, trễ hạn (AGENTS mục 13); rồi Báo cáo của GĐV, Ngày báo cáo, Người báo cáo.
- Phiếu KQTN: `MTTN…` (hóa lý, có cả vi sinh nước), `VSTN…` (vi sinh). Cột "Giới hạn phát hiện" dùng cho `< LOD`.

## 5. Đã chốt (người dùng 16/11/2026 — đã ghi AGENTS mục 2, 8C, 14)

| # | Điểm | Chốt |
|---|---|---|
| 1 | `QUATEST3` dính liền | Ghi `QUATEST 3 3030:2025, N2.TC.N.0x` (có dấu cách) |
| 2 | Tên máy đo nhanh | Theo danh mục: `Thiết bị đo pH Mettler Toledo Seven2Go S2, USA` (`eq("mettler_s2")`), `Thiết bị đo clo dư Hach DR900, USA` (`eq("hach_dr900")`) |
| 3 | Chính tả tên chỉ tiêu, QCVN | **Tên chỉ tiêu ghi theo phiếu KQTN** (vd `Enterococci (Streptococci faecal)`, `Staphylococcus aureus`); QCVN thống nhất `QCVN 06-1:2010/BYT`, `QCVN 01-1:2024/BYT` |
| 4 | Đánh giá NSH | **Luôn có** đoạn Đánh giá khi có mức QCVN (cả mẫu gửi) |
| 5 | Clo dư hiện trường < 0,1 | Ghi `< 0,1` |
| 6 | Thời gian giám định | `Từ ngày dd/mm/yyyy` |
| 7 | Header | `Số: <số> - Ngày: dd/mm/yyyy` (có dấu hai chấm) — riêng hồ sơ nước |

Hệ quả: các mục 2, 3 ở trên mô tả **mẫu cũ**; khi làm hồ sơ mới áp bảng này (vd tên chỉ tiêu `Arsen (As)` của mẫu cũ đổi theo KQTN `Hàm lượng asen (As)`).

## 6. Hồ sơ Hyosung đang có trong `Nước\` (chưa làm)

145046, 145055 (NU 1 mẫu), 145051 (NSH + NU, 6 vị trí N1–N6), 145067 (NSH + NU, 9 mẫu N1–N9): có scan BBLM / Ghi chép / PYC / báo giá, ảnh theo vị trí; thư mục `thuc hien` chứa BBLM/Ghi chép DOCX in sẵn (chưa điền) và **file mẫu chép từ hồ sơ khác** (137037, 090965, 133022) — không phải kết quả của các hồ sơ này.
Đã thấy khi đọc: Ghi chép DOCX 145067 ghi địa điểm `HYOSUNG VINA INDUSTRIAL MACHIERY, Đường N3` trong khi BBLM DOCX ghi `HS HYOSUNG VIỆT NAM, Đường N2`; tên vị trí N5 khác nhau giữa hai file → khi làm phải đối chiếu scan thật (AGENTS mục 0).
