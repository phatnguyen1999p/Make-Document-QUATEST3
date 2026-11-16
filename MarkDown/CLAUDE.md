# CLAUDE.md — Hồ sơ môi trường 2026 (phòng N2, QUATEST 3)

> **LỖI THỜI (ghi chú Claude, cập nhật 30/10/2026):** tài liệu của quy trình cũ (vá DOCX bằng PowerShell trên máy `n2-hangnln`). Nội dung còn hiệu lực đã gộp vào `THAM-KHAO-QUY-TRINH-CU.md`. **Không làm theo file này** — quy tắc chính là `AGENTS.md` + `PROJECT_CONTEXT.md` ở thư mục gốc. Lưu ý: quy tắc Z: đã đổi — **được lưu hồ sơ lên Z** theo `AGENTS.md` mục 0 (file AI tạo được ghi đè, file người dùng chỉ đọc).

Vai trò: chuyên gia kỹ thuật hỗ trợ lập chứng thư kết quả lĩnh vực quan trắc môi trường.

## Bắt buộc đọc trước khi làm

- `.claude/QUY-TRINH-XU-LY-HO-SO.md` — **làm gì, theo thứ tự nào**: các bước, quy ước, checklist
- `.claude/CONG-CU-VA-WORKFLOW.md` — **làm bằng cách nào**: script, công thức, các bẫy

Làm **đúng theo đó**, không tự nghĩ ra cách khác. Hai file trên là nguồn duy nhất cho từng nội
dung — file này chỉ nêu những gì phải biết trước khi mở chúng.

## Thư mục

| Đường dẫn | Nội dung |
|---|---|
| `Z:\9.NV2 - MOI TRUONG\1.Ho So Moi Truong\2026` | Thư mục tổng hồ sơ 2026 (~479 thư mục con, mỗi thư mục một hồ sơ) |
| `…\2026\0.VIMCERTS078` | **Hồ sơ Vimcerts** — không nằm ở thư mục tổng. Không thấy mã ở 2026 → tìm tiếp ở đây trước khi báo không có |
| `E:\HoSo\_claude\` | Script dùng chung (`*.ps1`, `pdf-img.js`), `backup\`, `maps\` |
| `E:\HoSo\_claude\_claude_tmp\<6 số>\` | File trung gian của từng hồ sơ — **không để file rác trong thư mục hồ sơ** |
| `Input\<tên hồ sơ>\TBKQ\` | **Thư mục xuất kết quả TBKQ:** Toàn bộ file chứng thư (TBKQ .docx/.pdf) của hồ sơ xuất vào thư mục con `TBKQ` ngay trong thư mục `Input` tương ứng |

Mất ổ mạng: `net use Z: \\192.168.2.9\n2$ /persistent:yes`

## Máy và cách chạy lệnh

`n2-hangnln`: PowerShell 5.1, **Node.js v24**, **không có Python**, không có `pdftoppm`.

- Chạy script: **ghi ra `.ps1` thuần ASCII rồi chạy bằng `-File`**
  `powershell.exe -NoProfile -ExecutionPolicy Bypass -File "<script>.ps1" -Code Q326Dxxxxxx`
- **Không đặt tiếng Việt hay ký tự ngoài ASCII trong `.ps1`** (kể cả `…`) — dữ liệu truyền qua
  file JSON UTF-8. Chi tiết: Bẫy 1–2 trong `CONG-CU-VA-WORKFLOW.md`.
- Quyền truy cập `Z:` **mặc định đã cấp** — không hỏi lại.

## Bắt đầu phiên

1. Chạy `E:\HoSo\_claude\startup-check.ps1` — kiểm tra công cụ, dọn `_to_delete_*` và
   `_claude_tmp` cũ hơn 30 ngày. Báo cái nào thiếu.
2. Báo sẵn sàng, **chờ người dùng đưa mã hồ sơ**. Không tự đoán hồ sơ.

## Tiết kiệm token

- **Không quét toàn thư mục năm.** Vào thẳng thư mục hồ sơ bằng wildcard 6 số.
- **Không quét thư mục tổng để tìm hồ sơ mẫu.** Mẫu tham chiếu, nếu cần, người dùng đã đặt sẵn
  trong chính thư mục hồ sơ đang làm.
- Không tính lại khi đã có bản scan/file tính toán; chỉ đọc PDF thật sự cần; đọc `map.txt`
  và xlsx theo range; gộp nhiều bước vào một script.
- `/clear` giữa hai hồ sơ.
