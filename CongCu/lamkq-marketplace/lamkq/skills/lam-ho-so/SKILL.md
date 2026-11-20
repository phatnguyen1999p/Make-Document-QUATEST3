---
name: lam-ho-so
description: Quy trình làm một hồ sơ giám định "Làm KQ" (QUATEST 3) — BCKQ, TBKQ VN/EN, phiếu tính toán Excel, tem tủ, PDF; gồm cả hồ sơ nước uống, nước sinh hoạt, nước thải, nước dưới đất, bùn thải / chất thải. Dùng khi người dùng giao một hồ sơ theo 6 số (vd "làm hồ sơ 139329", "dựng lại 137058", "KQTN về rồi"), hoặc hỏi tiến độ một hồ sơ.
---

# Làm một hồ sơ "Làm KQ"

Nguồn quy tắc duy nhất là `AGENTS.md` ở gốc dự án (ưu tiên cao nhất). Skill này chỉ là thứ tự các bước, **không** thay thế AGENTS.md — gặp chỗ khác nhau thì theo AGENTS.md.

## Thứ tự

1. **Đầu phiên** — hook của plugin đã chạy `python Core/Tools/context.py dau-phien` và in **nhãn phiên** (`<máy>-<HHMM>-<4 ký tự>`). Không thấy bản tóm tắt trong ngữ cảnh thì tự chạy lệnh đó. Không đọc `ContextLuuTru/`.
2. **Khóa hồ sơ** (AGENTS mục 0, 16): `python Core/Tools/context.py khoa <6 số> --ai <nhãn>`.
   - `[KHÓA BẬN]` → dừng, báo người dùng; không dựng, không chép Z.
   - Hồ sơ mới: `moi <6 số> --khach "…" --doi-tuong "…" --ai <nhãn>`. Hồ sơ đã lưu trữ: `mo <6 số> --ai <nhãn>` rồi đọc file lưu trữ nó trỏ tới.
   - Đọc **chỉ** `Context/HoSo/<6 số>.md`.
3. **Lấy đầu vào** — hồ sơ trên Z: tìm thư mục bằng wildcard 6 số (AGENTS mục 0), chép scan / ảnh / `Data Do` / KQTN về `Input\<6 số>\` (chỉ đọc). Không thấy hoặc trùng nhiều thư mục → hỏi.
4. **Đọc hồ sơ** (mục 13): BBLM trước, rồi PYC, phiếu ghi chép, nhãn tủ trong ảnh, KQTN. Lệch thông tin khách hàng giữa các nguồn → **không tự sửa**, ghi vào `Cần người dùng quyết`, không áp tiền lệ hồ sơ khác.
   - **Đối chiếu khách hàng với mã số thuế** (mục 15, bắt buộc trước khi dựng): lấy MST trên PYC / GCN / PO / báo giá, chạy `python Core/Tools/tax_helper.py --mst <MST> --name "<tên trên PYC/BBLM>" --addr "<địa chỉ trên PYC/BBLM>"`. Chỉ khác cách viết → chuẩn hóa; lệch số nhà / tên đường / tên công ty / MST ngừng hoạt động → đưa vào câu hỏi gộp. Địa chỉ MST là trụ sở (mục 1 Table 0), Địa điểm giám định vẫn theo BBLM. Không có MST → hỏi người dùng. Ghi MST + kết quả vào khối Nguồn của script hồ sơ.
5. **Dựng** — hồ sơ thường: `Dossier` + `DossierRunner` (mục 10); hồ sơ đặc thù: `build_<6 số>.py` theo mẫu gần nhất (docstring 4 khối: Nguồn → Quyết định → Suy luận → Chạy). Luôn `network_dir=""`.
6. **Gom câu hỏi một lần** (mục 15): sau lần dựng đầu, liệt kê mọi điểm cần người dùng quyết, bỏ điểm đã có quy tắc mặc định (ngày ban hành mục 2, tên máy mục 14), hỏi gộp.
7. **QA** (mục 4): đóng mọi file của hồ sơ trước khi chạy; `python Core\Tools\qa_check.py final-check <path> --code <mã>` phải `[CLEAN]`; xuất PDF lại sau mỗi lần dựng; soát PDF bằng mắt (bìa 1 trang, VN = EN số trang và chỗ ngắt).
8. **Chép lên Z** → dùng skill `chep-z`.
9. **Ghi lại và trả khóa** (mục 16): sửa 3 dòng đầu `Trạng thái` / `Bước tiếp` / `Cần người dùng quyết` + dòng trên cùng bảng Nhật ký của `Context/HoSo/<6 số>.md`; rồi `context.py tra <6 số> --ai <nhãn>`.
10. **Báo cáo cuối**: file đã tạo/sửa (đã lên Z chưa), ô còn trống, điểm cần người dùng xác nhận.

## Ghi nhớ đi theo plugin

Các ý dưới đây trước chỉ nằm trong bộ nhớ Claude của máy N2-PHAT; đưa vào đây để máy khác cũng biết. Đã có trong AGENTS.md thì theo AGENTS.md.

- Không có KQTN → ngày ban hành = ngày hiện tại (`CERT_DATE = date.today()`), tự cập nhật cả khi dựng lại, không hỏi lại (mục 2).
- Hỏi gộp một lần mọi điểm cần người dùng quyết (mục 15).
- KQTN vi sinh ghi `< 1` → mọi nơi ghi `< 1`, không ghi `0`, cả cột CFU/đĩa của phiếu tính toán (mục 13).
- Khách hàng quen: xem `references/khach-quen.md`.
- Hồ sơ nước / nước thải / nước dưới đất / bùn thải / chất thải (chưa có handler Core, mẫu `build_nuoc_mau.py`): đọc `references/ho-so-nuoc.md` trước khi dựng.
