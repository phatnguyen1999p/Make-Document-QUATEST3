---
name: hoan-tat-kqtn
description: Hoàn tất hồ sơ giám định môi trường N2 (QUATEST 3) khi đã có phiếu KQTN — đọc phiếu KQTN (PDF khoá), nhập kết quả vào file tính toán, điền kết quả/phân loại vào TBKQ VN-EN và BCKQ, cập nhật mọi ngày, gỡ tô vàng, chép KQTN vào Edoc. Dùng khi người dùng nói "đã có KQTN", "có kết quả rồi", "hoàn tất hồ sơ", "điền kết quả" cho một mã hồ sơ đã làm khung trước đó.
user-invocable: true
---

# Hoàn tất hồ sơ khi có KQTN

Hồ sơ đã được làm khung (skill `thuc-hien-chung-thu`), còn ô `………` tô vàng chờ kết quả. Việc
bây giờ là điền nốt — **không** làm lại khung. Quy trình chuẩn: `QUY-TRINH-XU-LY-HO-SO.md`
**Bước 6**; cách làm, script, bẫy: `CONG-CU-VA-WORKFLOW.md`. Trả lời bằng tiếng Việt.

## 1. Định vị và nắm hiện trạng (một lượt lệnh)

- Wildcard 6 số (không thấy → `0.VIMCERTS078`). Liệt kê file + thư mục con kèm ngày sửa.
- Tìm phiếu KQTN: `0. KQ-<mã>-<số phiếu>.pdf`, thư mục `*edoc*`, `*TBKQ*\Edoc`. Có thể nhiều phiếu
  (VSTN vi sinh, MTTN dầu hơi/môi trường, HDTN dầu hạt…) — mỗi phiếu có `.01`, `.02`… theo mẫu.
- `dump-all -OutFile _claude_tmp\<6 số>\map.txt`, grep `………`, `\.\.\.`, `/20\d\d` → danh sách ô
  còn trống, ngày đang ghi, header nào có ngày.

## 2. Đọc KQTN

- `pdf-render.ps1` (mặc định render hết trang). Kiểm tra `grep -a "/Count"` = số trang đã render.
- Mỗi mẫu 2 trang: trang lẻ tên mẫu + ngày nhận mẫu, trang chẵn bảng kết quả. Cắt dải đầu trang
  bằng `img-crop` (`-X 100 -Y 270 -W 1450 -H 420 -Scale 0.5`) rồi Read — tiết kiệm token.
- **Đếm số mẫu, đối chiếu ký hiệu mẫu trên BBLM/Ghi chép, kể cả blank** (Bẫy 16). Thiếu mẫu →
  kiểm tra lại số trang trước khi báo người dùng.
- Ghi lại: số phiếu, ngày phiếu, ngày nhận mẫu, kết quả từng chỉ tiêu, giới hạn phát hiện, ghi chú
  bất thường (vd đơn vị in sai "CFU/g") để nêu trong báo cáo.

## 3. File tính toán

Theo skill `nhap-tinh-toan-excel` / `CONG-CU` mục 7: thay hết vùng nhập liệu (file thường chép từ hồ
sơ cũ — Bẫy 18), đọc kết quả + phân loại, quét ô lỗi, đối chiếu bản scan tính toán.

## 4. Điền chứng thư và báo cáo

- `3.*` VN và EN: bảng kết quả, phân loại, kết luận. Định dạng số theo mẫu (VN dấu phẩy thập phân,
  EN dấu chấm; `< 10`, "Không có chỉ thị" / *No indication*…). Không đạt → in đậm.
  KQTN ghi KPH → cột kết quả TBKQ nhập dạng `< LOD` (vd `< 0,01`), **không nhập `KPH`** (QUY-TRINH Bước 6).
- `2.*`: số phiếu KQTN, số phiếu tính toán (chưa có → nối tiếp `/1`, `/2`… và báo), dòng KQTN mẫu
  blank, các mục "Gửi mẫu thử nghiệm" (a, b, c… mỗi phiếu một mục).
- Ngày — một quy ước cho cả hồ sơ (QUY-TRINH Bước 6):
  - ngày ban hành TBKQ = ngày phiếu KQTN (nhiều phiếu → muộn nhất), ở **mọi** header có ngày
    (trang đầu + phụ lục, VN + EN);
  - BCKQ: ngày KQTN = ngày nhận KQTN = ngày kết thúc = ngày báo cáo = ngày phiếu KQTN;
  - ngày yêu cầu cấp KQTN = "Ngày dự kiến trả kết quả" trên PYC (`*YCTN`, PDF khoá → render);
  - ngày gửi mẫu = "Ngày nhận mẫu" trên KQTN.
- Kiểm tra lại người báo cáo (Toàn/Nghĩa → `<tên> + Minh`; còn lại `<tên> + Minh + Nam`) và giờ thực hiện — mẫu hay
  sót giờ (vd "Bắt đầu từ :00").
- Một `plan.json` cho cả hồ sơ: mọi ô vừa điền đưa vào `unhl`; ô vẫn chưa có dữ liệu giữ `hl`.

## 5. Kết thúc

- Chép phiếu KQTN vào thư mục Edoc của hồ sơ: `<6 số> - KQTN <số phiếu>.pdf`.
- Dump lại, grep `………` + đếm `w:highlight` trong **mọi** entry (document + header) — còn tô vàng
  thì phải là ô cố ý chờ người dùng.
- `final-check -Old "<số phiếu cũ>,<mã hồ sơ mẫu cũ>,<tên người cũ>…"` CLEAN; `test-open`.
- Báo cáo: bảng kết quả tóm tắt; file đã sửa; ngày đã chọn và lý do; điểm chờ xác nhận (mẫu thiếu,
  đơn vị lạ, ngày mẫu ghi sẵn khác quy ước, thư mục con còn bản TBKQ cũ…).
