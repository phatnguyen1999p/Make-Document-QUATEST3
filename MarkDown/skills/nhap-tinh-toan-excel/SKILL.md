---
name: nhap-tinh-toan-excel
description: Nhập số liệu đo/KQTN của một hồ sơ N2 (QUATEST 3) vào file tính toán Excel — khí nén (1. tinh toan tong dau.xlsx, sheet Tinh_Toan_Khi_Co_Ap_V6 + DiemSuong) và vi sinh không khí (0. Tinh Toan VSKK.xlsx) — thay ảnh trong sheet, sửa ô lỗi công thức, rồi đối chiếu kết quả với bản scan và chứng thư. Dùng khi người dùng nói "file tính toán bị lỗi", "nhập liệu vào file tính toán", "tính toán tổng dầu / vi sinh", "thay ảnh trong file Excel".
user-invocable: true
---

# Nhập liệu file tính toán Excel

Chi tiết ô nhập / nguồn số liệu: `CONG-CU-VA-WORKFLOW.md` **mục 7** (7a khí nén, 7b vi sinh
không khí). Bẫy liên quan: 15 (`#VALUE!` vi sinh khí nén), 17 (ghi ô qua COM), 18 (file chép từ
hồ sơ cũ). Trả lời bằng tiếng Việt.

## Nguyên tắc

- **Không khôi phục bản gốc** chỉ vì thấy ô lỗi — mở bằng Excel COM, liệt kê ô lỗi + công thức,
  tìm nguyên nhân. Thường là lỗi công thức mẫu hoặc dữ liệu cũ, không phải file hỏng.
- Backup trước mỗi lần ghi: `E:\HoSo\_claude\backup\<tên file> - <mã>_<yyyyMMdd_HHmmss>.xlsx`.
- Chỉ ghi vào **vùng nhập liệu**; vùng kết quả/tổng hợp là công thức — không gõ đè. Sửa công thức
  chỉ khi công thức sai logic (vd Bẫy 15), và nói rõ trong báo cáo.
- Nguồn số theo thứ tự: bản scan tính toán đã ký → file nguồn của máy đo (`thuc hien\kq …`) →
  phiếu KQTN → BBLM/Ghi chép (lưu lượng, thời gian, giờ). Không lấy số từ chứng thư để nhập ngược.

## Các bước

1. **Dò file** (Excel COM, chỉ đọc): liệt kê mọi sheet, `UsedRange`, ô có công thức/giá trị trong vùng
   nhập, ô lỗi (`Text` bắt đầu `#`), shape/ảnh (`Shapes`, ô chữ `Picture` = ảnh trong ô) và merge area.
   So số phiếu, ngày, số đo đang có với hồ sơ hiện tại → biết ô nào còn của hồ sơ cũ.
2. **Chuẩn bị JSON UTF-8** `_claude_tmp\<6 số>\in.json` dạng `{ "<ô>": giá trị }` theo từng sheet (mẫu khí nén dùng khoá `s1` =
   sheet tính toán, `ds` = `DiemSuong`) — chuỗi tiếng
   Việt nằm trong JSON, script `.ps1` thuần ASCII (Bẫy 1). Mẫu: `E:\HoSo\_claude\kn-fill-excel.ps1`
   + `maps\kn-fill-Q326D100311.json`.
3. **Ghi** qua COM: `$rg=$ws.Range('<ô>'); $rg.Formula=<chuỗi InvariantCulture>` (Bẫy 17); ngày dùng
   `=DATE(y,m,d)`; dòng mẫu thiếu công thức → `Range('<dòng trên>').Copy(<dòng dưới>)`.
   `CalculateFull()`, `Save()`.
4. **Ảnh** (nếu mẫu có): thay bằng ảnh của hồ sơ này — ảnh trong ô chép bằng `Range.Copy` từ file
   nguồn cùng bố cục; shape nổi dùng Copy/Paste với `$x.Visible=$true`; ảnh file dùng
   `Shapes.AddPicture` giữ tỉ lệ vừa merge area. Xoá shape cũ trước khi chèn.
5. **Kiểm tra**: quét lại ô lỗi toàn workbook (phải 0); in các ô kết quả; đối chiếu từng số với bản
   scan và với `3.*` đang ghi. Lệch → báo người dùng (lệch ±1 do scan làm tròn hệ số là chấp nhận
   được, chứng thư theo scan).
6. **Báo cáo**: ô đã nhập (theo nhóm, nguồn), bảng so sánh Excel ↔ chứng thư, ảnh đã thay, phần
   chưa làm được, vị trí backup. Phát hiện quy tắc mới về file tính toán → cập nhật `CONG-CU` mục 7.
