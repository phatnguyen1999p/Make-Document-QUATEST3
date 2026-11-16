---
name: thuc-hien-chung-thu
description: Xử lý một hoặc nhiều hồ sơ giám định môi trường (vá BCKQ/TBKQ từ dữ liệu BBLM theo mã hồ sơ) cho phòng N2, QUATEST 3. Dùng khi người dùng đưa mã hồ sơ (6 số, hoặc Q326Dxxxxxx, hoặc mã ngắn kiểu 90965/093N2.25DG.V078) và nói "thực hiện"/"xử lý"/"làm" hồ sơ, kể cả khi đưa nhiều mã cùng lúc.
user-invocable: true
---

# Thực hiện chứng thư — hồ sơ giám định môi trường N2

Kích hoạt khi người dùng đưa **mã hồ sơ** và yêu cầu thực hiện/xử lý — dù gõ số trần
("106978"), mã đầy đủ ("Q326D106978"), hay mã kiểu VIMCERT ("093N2.25DG.V078").

## Bắt buộc trước khi làm bất cứ việc gì

Đọc và làm **đúng theo** hai tài liệu cùng thư mục `.claude/`:

- `../QUY-TRINH-XU-LY-HO-SO.md` — quy trình chuẩn (định vị hồ sơ, đọc BBLM, vá 2.*/3.*,
  đặt tên file, checklist bàn giao)
- `../CONG-CU-VA-WORKFLOW.md` — công cụ, script tại `E:\HoSo\_claude\`, các bẫy đã gặp,
  công thức, TBKQ dạng .xlsx (mục 6), file tính toán Excel theo loại hồ sơ (mục 7)

Skill liên quan: `hoan-tat-kqtn` (hồ sơ đã có khung, nay có KQTN), `nhap-tinh-toan-excel`
(nhập số liệu vào file tính toán khí nén / vi sinh không khí).

Không tự nghĩ ra cách khác. Hai file trên là nguồn sự thật; skill này chỉ là lối vào nhanh +
vài quy ước bổ sung phát sinh khi vận hành thực tế mà chưa kịp gộp vào đó.

## Trình tự một hồ sơ (tóm tắt — chi tiết ở QUY-TRINH)

1. `startup-check.ps1` nếu là đầu phiên; định vị thư mục bằng wildcard 6 số (không thấy → `0.VIMCERTS078`).
2. Đọc BBLM (và Ghi chép nếu BBLM ghi gộp vị trí) → khách hàng, ngày giờ, vị trí, ký hiệu mẫu,
   chỉ tiêu, người thực hiện. Liệt kê các file `0.*`–`4.*` và thư mục con có sẵn.
3. `dump-all -OutFile _claude_tmp<6 số>map.txt` + `set-cells -Dump` → biết mẫu đang mang gì của hồ sơ cũ.
4. Vá `2.*`, `3.*` (plan.json → `apply-all`, rồi `set-cells`), đổi tên file có đuôi mã hồ sơ,
   thay ảnh vị trí, ô chưa có dữ liệu để `………` tô vàng.
5. Chưa có KQTN → dừng ở khung; có KQTN → theo skill `hoan-tat-kqtn`.
6. `final-check` CLEAN, `test-open` mở được + đúng tổng trang; báo cáo cuối bằng tiếng Việt.

## Xử lý nhiều hồ sơ cùng lúc

Người dùng thường đưa 2+ mã hồ sơ trong một câu ("thực hiện cùng lúc 2 hồ sơ X và Y").
Mỗi hồ sơ **độc lập hoàn toàn** (thư mục khác nhau, không chung dữ liệu) → luôn xử lý bằng
`Agent` với `subagent_type: "fork"`, **một fork riêng cho mỗi hồ sơ**, gọi trong cùng một
lượt (nhiều Agent call trong một message) để chạy song song thật sự. Không xử lý tuần tự
trong cùng một phiên chính — vừa chậm vừa dễ lẫn context giữa các hồ sơ.

Khi giao việc cho fork, prompt phải nói rõ:
- Đường dẫn/mã hồ sơ cụ thể (dùng wildcard định vị, không liệt kê cả thư mục năm).
- File trung gian để riêng `E:\HoSo\_claude\_claude_tmp\<mã hồ sơ>\` — **không dùng chung**
  với fork khác đang chạy (đặc biệt `dump-all.ps1 -OutFile` phải trỏ vào thư mục riêng đó,
  vì mặc định của script là một file dùng chung `E:\HoSo\_claude\map.txt`).
- Tự làm hết các bước, không hỏi lại người dùng giữa chừng trừ khi thật sự bế tắc (không
  tìm thấy thư mục/mẫu tham chiếu, chữ viết tay không đọc nổi dù đã phóng to, thiếu dữ liệu
  không thể suy luận).
- Báo cáo cuối ngắn gọn: đã làm gì / file nào tạo-sửa-đổi tên / final-check sạch chưa /
  danh sách rõ ràng các điểm cần người dùng xác nhận.

Nếu một mã hồ sơ trong danh sách **không tìm thấy** thư mục (wildcard rỗng) → báo ngay cho
người dùng ở mã đó, đồng thời vẫn khởi chạy fork cho các mã tìm thấy — không chờ toàn bộ.

## Hồ sơ chạy dở bị gián đoạn (rate limit / lỗi khác)

Nếu một fork dừng giữa chừng (status `failed`, hoặc kết quả cụt ngang không phải báo cáo
hoàn chỉnh) → resume đúng fork đó bằng `SendMessage` (theo tên đã đặt lúc spawn), yêu cầu nó
**tự kiểm tra lại trạng thái file thật** trước khi làm tiếp (đối chiếu bản ghi nhớ đã làm gì,
kiểm tra backup trong `E:\HoSo\_claude\backup\` nếu nghi ngờ file dở dang, mở thử bằng
Word/Excel COM) — không được vá chồng thêm một lượt nữa lên trạng thái không rõ ràng.

## Vòng phản hồi với người dùng

Sau khi fork báo cáo, chuyển tiếp cho người dùng bằng tiếng Việt, ngắn gọn, nêu rõ:
việc đã xong / cảnh báo quan trọng (dữ liệu sai khách hàng còn sót, file chưa nên gửi
khách...) / các điểm chờ xác nhận. Khi người dùng phản hồi xác nhận hoặc bổ sung dữ liệu,
`SendMessage` thẳng tới đúng fork đó (resume theo tên) để nó tự vá tiếp — không tự vá thay
ở phiên chính, vì fork mới có đủ context chi tiết (chỉ số run, tên file, cấu trúc bảng đã
dò) của đúng hồ sơ đó.

## Khi phát hiện quy tắc/bẫy mới trong lúc làm

Nếu trong lúc xử lý phát hiện một pitfall mới (kiểu Bẫy 3 header, hay quy tắc TBKQ
liên kết công thức ở mục 6) mà `CONG-CU-VA-WORKFLOW.md` chưa ghi — cập nhật thẳng vào file
đó (không chỉ nói miệng), để lần sau không lặp lại lỗi tương tự. Quy ước về **việc gì / thứ tự** →
`QUY-TRINH-XU-LY-HO-SO.md`; **cách làm / script / bẫy** → `CONG-CU-VA-WORKFLOW.md`. Không ghi trùng
một nội dung ở cả hai file — file kia chỉ trỏ sang.

## Ngôn ngữ

Luôn trả lời người dùng bằng **tiếng Việt**.
