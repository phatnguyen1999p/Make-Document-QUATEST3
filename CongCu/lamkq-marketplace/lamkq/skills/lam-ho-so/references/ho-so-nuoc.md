# Hồ sơ nước, nước thải, bùn thải / chất thải

Học từ hồ sơ mẫu 16/11/2026 (113750, 114119, 116644; 100302, 106141, 106222, 106227, 107520, 108037, 109764, 110364, 110370). Quy tắc chính thức: `AGENTS.md` mục 2, 3, 8C ("Nước uống / NSH", "Nước thải / …"), 14. Bản đầy đủ (chỉ có trên máy N2-PHAT, thư mục dữ liệu không lên git): `Nước\00 - Form ket qua nuoc (AI hoc).md`, `Nước\Nước Thải, Bùn Thải\00 - Form nuoc thai bun thai (AI hoc).md`.

**Chưa có handler Core.** Mẫu code: `build_nuoc_mau.py` (4 hồ sơ ảo Q326D999911–999914) — bìa qua `CleanroomHandler`, phụ lục dựng lại; BCKQ qua `BCKQBuilder`, thay mục 4.5 / 4.6 / Đánh giá / Gửi mẫu. Hồ sơ thật: chép thành `build_<6 số>.py`, đổi số liệu theo BBLM / KQTN.

## Loại hồ sơ

| Loại | Căn cứ | Quy trình | Câu công nhận (Ghi chú) |
|---|---|---|---|
| Nước uống (máy lọc, bình) | `QCVN 06-1:2010/BYT` phụ lục II (hóa lý), III (vi sinh) | `N2.TC.N.07` | VILAS 004 rồi VIAS 004 (đứng cuối) |
| Nước sinh hoạt / sạch / cấp | `QCVN 01-1:2024/BYT` nhóm A | `N2.TC.N.01` | VILAS 004 rồi VIAS 004 |
| Nước thải | `QCVN 14:2025/BTNMT`, `QCVN 40:2011/BTNMT` (mức = C × Kq × Kf), mức KH (KCN, VSIP, KCNC) | `N2.TC.N.06` | `● Các thông số được chứng nhận VIMCERTS 078.` **dòng đầu** |
| Nước dưới đất | `QCVN 09:2023/BTNMT` | `N2.TC.N.05` | VIMCERTS dòng đầu |
| Bùn thải / chất thải | `QCVN 07:2025/BNNMT` | `N2.TC.B.01` (mẫu gửi: chỉ `QUATEST 3 3030:2025`) | VIMCERTS dòng đầu |

Câu công nhận xét theo **quy chuẩn**, không theo thư mục Z: nước `QCVN …/BYT` → VIAS 004; nước còn lại + bùn / chất thải → VIMCERTS 078 (người dùng chốt 16/11/2026). `final-check` kiểm (`qa_check.is_vimcerts`).

## Điểm người dùng đã chốt (16/11/2026)

- `QUATEST 3 3030:2025` có dấu cách (mẫu cũ `QUATEST3` dính liền — không theo).
- Tên máy theo danh mục: `eq("mettler_s2")`, `eq("hach_dr900")`, `eq("hach_dr300")`, `eq("hanna_hi98192")`, `eq("hanna_hi98190")`.
- Nước uống / sinh hoạt: **tên chỉ tiêu ghi theo phiếu KQTN** (vd `Hàm lượng asen (As)`, `Enterococci (Streptococci faecal)`).
- Luôn có đoạn `Đánh giá:` (có dấu hai chấm, cả bùn thải) khi có mức.
- Clo dư đo hiện trường đọc dưới 0,1 mg/L → `< 0,1`.
- Thời gian giám định `Từ ngày dd/mm/yyyy`.
- Header chứng thư nước **BYT**: `Số: <số> - Ngày: dd/mm/yyyy` (có dấu hai chấm — ngoại lệ AGENTS mục 2).
- Nước thải / nước dưới đất: luôn `mg/L`, dùng chữ **"thông số"**.
- Hồ sơ có 2 ảnh (mẫu niêm phong + vị trí lấy mẫu): giữ 2 ảnh trên dưới như mẫu cũ.

## Khuôn phụ lục TBKQ

- **9.1** `- Vị trí lấy mẫu và chỉ tiêu thử nghiệm theo yêu cầu của <KH>, cụ thể như sau:` → bảng `Vị trí lấy mẫu | Ký hiệu mẫu | Chỉ tiêu thử nghiệm`; `- Thông tin về quá trình lấy mẫu và hình ảnh mẫu niêm phong:` → bảng `Người chứng kiến | Ký hiệu mẫu | Thời điểm lấy mẫu | Số niêm phong` + ảnh (không niêm phong: `: -`, bỏ bảng); `- Hình ảnh vị trí lấy mẫu:` + 1 ảnh căn giữa. Mẫu khách gửi: bảng `Tên mẫu theo khai báo`, `Số lượng mẫu`, `Ngày nhận mẫu`, `Người gửi mẫu`, `Mô tả tình trạng mẫu`, `Chỉ tiêu yêu cầu thử nghiệm` + ảnh mẫu.
- **9.2** bảng `STT | Chỉ tiêu (Thông số) | Đơn vị | Phương pháp | Kết quả | Giới hạn … ⁽¹⁾`; gộp dọc ô phương pháp trùng nhau và ô mức chữ (`-`, `Không phát hiện`), **không** gộp mức số trùng nhau của 2 thông số khác nhau. Không có mức → bảng 5 cột, Căn cứ `-`, không Đánh giá.
- KQTN `KPH` → `< <giới hạn phát hiện>`; giới hạn định lượng → thêm ⁽²⁾ "Giới hạn định lượng của phương pháp thử.". Không đạt → in đậm. Nhóm chất (PCBs, thuốc BVTV): 1 STT, cấu tử `+ …` trong cùng ô.
- Đánh giá: `+ Chỉ tiêu (Thông số) không đánh giá / không đạt / đạt : …` (cụm từ đậm nghiêng); bùn thải: `Độ pH không vượt ngưỡng …` rồi `+ Chỉ tiêu cao hơn / thấp hơn mức quy định` cho ngâm chiết và hàm lượng tuyệt đối.
- Bùn thải 9.2.1 tính chất nguy hại (pH, `pH ≥ 12,5 hoặc pH ≤ 2,0`), 9.2.2 ngâm chiết (mg/L, bảng 2), 9.2.3 hàm lượng tuyệt đối (ppm, bảng 3) — quy đổi về mẫu rắn khô ở phiếu Excel (`% rắn khô = 100 − độ ẩm`, công thức động).
- Đoạn ngăn 1 pt giữa bảng kết quả và bảng Ghi chú (không có thì Word gộp 2 bảng); giữ tiêu đề 9.2 + 4 dòng đầu, và 2 dòng cuối + Ghi chú + Đánh giá đi liền (không giữ trọn cả bảng ngắn — để trống nửa trang).

## BCKQ

4.5 bảng vị trí / niêm phong / ảnh như TBKQ; 4.6 `- KQTN mẫu <ký hiệu> : xem tại phiếu KQTN số MTTN…` + bảng đo hiện trường `Kí hiệu mẫu | pH | Nhiệt độ [°C] | Clo dư [mg/L] | Mùi` (dòng `Thiết bị sử dụng`); bùn thải viện dẫn thêm `phiếu Tính toán hàm lượng tuyệt đối theo mẫu rắn khô số <mã> ngày …`; Đánh giá; Gửi mẫu thử nghiệm (mỗi phiếu KQTN một khối); trễ hạn theo `kqtn_late.tre_han_text`.

## Không chép từ mẫu cũ

Lỗi trong mẫu: `Streptococci feacal`, `Staphylococcus auresus`, `SWEWW`, `QCVN 06 - 1: 2010`, `Jones Lang 2aSalle (Việ1 Nam)`, ký hiệu ghi chú `(1) ...` (`final-check` coi `...` là chỗ trống), chữ 12,5pt, thiếu lặp tiêu đề bảng.
