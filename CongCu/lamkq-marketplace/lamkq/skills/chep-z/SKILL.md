---
name: chep-z
description: Chép kết quả một hồ sơ "Làm KQ" (ThucHien, TBKQ, pdf) lên thư mục hồ sơ trên ổ mạng Z an toàn, theo sổ file AI tạo (ai_files.json). Dùng khi hồ sơ đã dựng + QA xong, hoặc người dùng nói "chép lên Z", "đưa lên Z", "đồng bộ Z".
---

# Chép hồ sơ lên Z

Quy tắc gốc: `AGENTS.md` mục 0. Trên Z **không xóa gì**, không tạo thư mục hồ sơ mới, không chép vào thư mục kiểu `… (1)`.

## Điều kiện trước khi chép

- Đang giữ khóa hồ sơ (`context.py khoa <6 số> --ai <nhãn>`); khóa bận → không chép.
- Đã QA xong: `final-check` ra `[CLEAN]`, PDF trong `pdf\` mới hơn DOCX/XLSX nguồn.
- Không phải chạy mẫu/thử (`--sample-*`, mã `Q326D9999xx`).

## Cách chép (mặc định)

```
python Core/Tools/ai_files.py liet-ke <6 số>      # xem tình trạng
python Core/Tools/ai_files.py chep <6 số>         # chạy thử
python Core/Tools/ai_files.py chep <6 số> --ghi   # chép thật
```

- Bản trên Z **có trong sổ** → ghi đè được (kể cả do phiên / máy khác tạo).
- `[Z EDITED]` → có người sửa tay sau AI: **không ghi đè**. Chép bản Z về `Input\<6 số>\`, đưa chỉnh sửa vào script hồ sơ, dựng lại, so chữ, rồi mới chép (hoặc hỏi người dùng).
- `[Z Duplicate]` → có thư mục trùng `… (n)`: báo người dùng, không chép vào đó.
- `ERROR 32` / file đang mở → báo người dùng đóng file trên Z rồi chép lại riêng file đó.

## Dự phòng bằng robocopy

Chỉ khi `ai_files.py` không dùng được:

```
robocopy "Input\<6 số>\<ThucHien|TBKQ|pdf>" "<thư mục hồ sơ trên Z>\<cùng tên>" /E /XD _checkpoint _optimized /XF ~$* /R:1 /W:2
```

Exit 0–3 là thành công. **Cấm `/MIR`, `/PURGE`** (hook `chan_lenh_z` của plugin chặn). Trước khi chép so giờ sửa từng file trên Z; chép xong chạy `python Core/Tools/ai_files.py ghi <6 số>` để đưa vào sổ.

## Lưu ý nhiều máy

Sổ `ai_files.json` nằm trong `Input\<6 số>\ThucHien\_checkpoint\` của **máy đã dựng**. Máy khác mở lại hồ sơ đó sẽ không có sổ → mọi file trên Z hiện `[Z EDITED]`. Gặp trường hợp này: hỏi người dùng, không ghi đè (việc V9 trong `Context/Viec/`).
