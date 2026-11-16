# Quy trình xử lý hồ sơ

Làm gì, theo thứ tự nào, khi người dùng đưa mã hồ sơ.
Script / công thức / bẫy kỹ thuật: xem `CONG-CU-VA-WORKFLOW.md` cùng thư mục.
Đường dẫn, cấu hình máy, quy tắc chạy script: xem `CLAUDE.md` ở thư mục 2026.

---

## 1. Đầu vào và định vị hồ sơ

Người dùng đưa **mã hồ sơ**, thường chỉ 6 chữ số cuối (vd `108033`).

| Thành phần | Ví dụ | Ý nghĩa |
|---|---|---|
| `Q3` | Q3 | Tiền tố cố định |
| 2 chữ số | `26` | Năm (26 = 2026) |
| 1 chữ cái | `D` | Ký hiệu phân loại |
| 6 chữ số | `108033` | Số hồ sơ |

Đã gặp: `Q326D108033`, `Q326C087322`, `Q326A064001`, `Q326A073488`.

Tìm thư mục bằng **wildcard 6 số**, không liệt kê cả thư mục năm:

```powershell
(Get-Item 'Z:\9.NV2 - MOI TRUONG\1.Ho So Moi Truong\2026\*108033*').FullName
# khong thay -> ho so Vimcerts:
(Get-Item 'Z:\9.NV2 - MOI TRUONG\1.Ho So Moi Truong\2026\0.VIMCERTS078\*110364*').FullName
```

Không thấy ở cả 2026 lẫn `0.VIMCERTS078` → tìm tiếp trong thư mục khách quen **`2026\0.1.Tân Hiệp\`**
(hồ sơ tuần định kỳ Nhà máy nước Tân Hiệp, vd `0.1.T*\*104561*`). Lưu ý script: Bẫy 12 trong `CONG-CU-VA-WORKFLOW.md`.

Làm **trực tiếp trên thư mục đó**, không copy đi đâu. Không tìm thấy, hoặc trùng nhiều thư mục
→ dừng, hỏi người dùng.

---

## 2. Cấu trúc thư mục & quy ước đặt tên

| Tiền tố | Loại | Vai trò |
|---|---|---|
| `0.*` | Dữ liệu gốc | BBLM, ghi chép, phiếu KQTN, số liệu thô, **bản scan kết quả tính toán**, ảnh vị trí |
| `1.*` (hoặc `b.*`) | File tính toán | Thường là Excel |
| Bước 2 | Báo cáo kết quả (ThucHien) | Báo cáo nội bộ (BM3030.03) — tên `BCKQ - ...` (không đặt `2. BCKQ`), lưu `/ThucHien` |
| Bước 3 | Thông báo kết quả (TBKQ) | Chứng thư trả khách (VN/EN) — lưu `/TBKQ`, bản PDF lưu `/pdf` |
| Edoc | Thư mục scan trình ký | Bản scan tách (BBLM, Ghi Chép) — lưu `/TBKQ/Edoc` |

Một số hồ sơ đánh số khác (vd `1. BCKQDG` + `2. BKQ`) — theo đúng cách đặt sẵn của hồ sơ đó.

**Chỉ thao tác trên `0.*` → `4.*`.** Không mở các file rời khác (PYC, Payment, ảnh ngoài `0.*`…)
trừ khi thật sự cần số liệu từ đó.

### Đặt tên file

Mỗi file `1.*`, `2.*`, `3.*` **hễ bắt đầu vá cho hồ sơ hiện tại** → đổi tên **ngay**, thêm đuôi
mã hồ sơ đầy đủ, **không chờ nội dung xong** (kể cả khi còn ô trống/tô vàng chờ KQTN):

```
BCKQ - 1 - phong sach.docx      ->  BCKQ - 1 - phong sach - Q326D108033.docx
3. TBKQ - vn - 2 - khi nen.docx ->  3. TBKQ - vn - 2 - khi nen - Q326D108033.docx
```

Quy tắc: `<tên gốc> - <mã hồ sơ>.<đuôi>`. Giữ nguyên tên file `0.*`.
File trong `4. Edoc\` hoặc `TBKQ\E doc\` đặt `<6 số> - <tên>.pdf`, vd `108033 - BBLM.pdf`, `108033 - Ghi Chep.pdf`.
Đuôi mã hồ sơ chỉ đánh dấu file **thuộc hồ sơ nào**, không có nghĩa đã xong.

### Tự động tách file scan trình ký (E doc)

- Từ các file scan PDF đầu vào trong thư mục `Input\<tên hồ sơ>\`, hệ thống tự động nhận diện header và tách thành 2 file PDF chuẩn trình ký:
  + `<6_số> - BBLM.pdf`: Toàn bộ các trang Biên bản giám định hiện trường (BM3030.02 / Biên bản lấy mẫu/thực hiện).
  + `<6_số> - Ghi Chep.pdf`: Toàn bộ các trang Ghi chép giám định tại hiện trường (BM3030.01), các bảng thông số đo đạc, và tất cả các trang không có header rõ ràng (phiếu in nhiệt đếm hạt bụi, scan lọc HEPA, phiếu đo độ ồn, phụ lục...).
- **Vị trí lưu trữ bắt buộc:**
  + Lưu duy nhất vào thư mục con `TBKQ\E doc\` nằm trong thư mục hồ sơ: `Input\<tên hồ sơ>\TBKQ\E doc\` (tuyệt đối không lưu trùng tại thư mục gốc `Input\<tên hồ sơ>\`).
  + Đồng bộ sang `4. Edoc` trên ổ mạng (`Z:\...\4. Edoc\`) khi kết nối được mạng.
- **Thực thi:**
  + Tự động chạy trong `DossierRunner` (Step 0) khi xử lý hồ sơ.
  + Hoặc chạy độc lập qua lệnh: `python Core/run.py --split-scan "Input\<tên hồ sơ>"` hoặc `python Core/run.py --split-all-scans`.

### Cấu trúc thư mục chuẩn & Phân định xuất file

Áp dụng làm **nguyên tắc chung cho tất cả hồ sơ**:
- **Bước 2 — `/ThucHien`:** Toàn bộ file DOCX Báo cáo kết quả nội bộ (`BCKQ - ...docx`, không đặt tiền tố `2. `) và file tính toán Excel (`... - TinhToan....xlsx`).
- **Bước 3 — `/TBKQ`:** Toàn bộ file DOCX Thông báo kết quả khách hàng (`TBKQ - viet - ...docx`, `TBKQ - en - ...docx`), Tem tủ (`tem tu...docx`), sơ đồ layout (`layout_....png`).
- **Scan PDF — `/TBKQ/Edoc`:** Toàn bộ file scan PDF tách tự động (`<mã> - BBLM.pdf`, `<mã> - Ghi Chep.pdf`).
- **Tập trung PDF — `/pdf`:** Toàn bộ file PDF xuất ra từ DOCX (`BCKQ - ...pdf`, `TBKQ - viet - ...pdf`, `TBKQ - en - ...pdf`, `tem tu - ...pdf`...) được gom chung vào duy nhất thư mục `/pdf` ở gốc hồ sơ (`Input\<tên hồ sơ>\pdf\`). Tuyệt đối không để PDF rải rác trong `ThucHien/` hay `TBKQ/`.

### Người ký TBKQ (giám định viên)

Tên ghi ở TBKQ xác định theo **người thực hiện** trên BBLM:

| Người thực hiện | TBKQ ghi |
|---|---|
| Toàn | Nguyễn Khánh Toàn |
| Nghĩa | Phạm Quốc Nghĩa |
| Mọi trường hợp khác (Huy, Nam, ghép nhiều người…) | Nguyễn Xuân Nam |

Không lấy tên theo chữ ký trên BBLM, không giữ tên của mẫu cũ.

### Số chứng thư TBKQ

- Hồ sơ **chỉ có 1 TBKQ** → ghi mã hồ sơ trần, **không thêm `/1`**: `Q326D113750`.
- Hồ sơ **có nhiều TBKQ** (vd mỗi mẫu một bản) → thêm `/1`, `/2`…: `Q326D104563/1`,
  `Q326D104563/2`. Không dùng dạng `/3.1`, `/3.2` của mẫu cũ.

---

## 3. Nguyên tắc xuyên suốt

- **Dữ liệu gốc là chuẩn:** mọi con số phải truy ngược được về `0.*` hoặc bản scan tính toán.
- **Đọc nội dung trước khi kết luận:** file mẫu trong thư mục hồ sơ có thể mang mã hồ sơ khác mà
  vẫn **đúng loại chứng thư** — phải mở đọc nội dung, không loại bỏ chỉ vì thấy mã lạ.
- **Tận dụng file mẫu:** luôn sửa từ mẫu có sẵn **trong chính thư mục hồ sơ đang làm**, vá thẳng
  vào file mẫu đó (kể cả khi nhân bản ra nhiều bản). Không cần giữ lại bản gốc chưa sửa.
  Không quét thư mục tổng 2026 để tìm mẫu.
- **Giữ nguyên format:** với `2.*`/`3.*` tuyệt đối không đổi đầu mục, bố cục, font, kiểu bảng.
  Được **thêm dòng** cho đủ số vị trí/mẫu; **không tự thêm cột** so với mẫu — thấy mẫu thiếu
  thì nêu trong báo cáo cuối để người dùng quyết.
- **Số lượng hình ảnh mapping vị trí lấy mẫu (sơ đồ quan trắc) và hình ảnh thực hiện có thể thay đổi theo hồ sơ (tăng hoặc giảm):**
  không cố định theo số khung có sẵn của mẫu cũ. Hồ sơ có nhiều ảnh/sơ đồ hơn mẫu → thêm dòng/ô/khung ảnh (cùng cỡ và kiểu chú thích như khung có sẵn); hồ sơ có ít ảnh/sơ đồ hơn mẫu → xoá bớt khung/ô/dòng thừa, không để khung trống thừa hoặc sót ảnh cũ. Cách làm: `CONG-CU-VA-WORKFLOW.md` mục 4 và `QUY-TAC-FONT-VA-THIET-BI.md` Phần 4.
- **Quy tắc chuyển trang (bảng có câu kết luận kèm theo):** bảng kết quả có câu kết luận/đánh giá đi kèm (`Đánh giá:` / `Assessment:` + câu kết luận, cùng bảng `Ghi chú:` / `Note:` ngay dưới bảng) thì **câu kết luận phải nằm cùng trang với bảng** — tuyệt đối không để câu kết luận bị đẩy lẻ loi sang đầu trang tiếp theo (thiết lập `keepNext` cho dòng cuối bảng + bảng Ghi chú + dòng `Đánh giá:`, `keepLines` cho câu kết luận, hoặc điều chỉnh ngắt trang). Xem `QUY-TAC-FONT-VA-THIET-BI.md` Phần 5.
- **Xoá sạch dấu vết hồ sơ cũ:** mã hồ sơ, ngày tháng, số phiếu KQTN, tên/địa chỉ khách cũ,
  số liệu cũ — kể cả ở header/footer và ở phần diễn giải (nguyên nhân trễ hạn, ghi chú…).
- **Thiếu thì để trống + tô vàng:** điền `………`, highlight vàng, không bịa, không giữ giá trị
  hồ sơ cũ; liệt kê trong báo cáo cuối.
- **Kết quả không đạt → in đậm** giá trị đó (cả ở `2.*` và `3.*`). Ngược lại, nếu mẫu cũ in đậm
  sẵn một giá trị mà hồ sơ mới lại đạt → **gỡ in đậm** (`set-cells.ps1` có `unbold`).
- **Cụm từ đánh giá kết luận → in đậm và nghiêng:** Các cụm từ kết luận đánh giá sau mục `Đánh giá:` / `Assessment:` phải được **in đậm và in nghiêng** (`bold=True, italic=True`):
  + Tiếng Việt: **_phù hợp_**, **_không phù hợp_**
  + Tiếng Anh: **_comply_**, **_complied_**, **_did not comply_**
- **Ảnh vị trí lấy mẫu & hình ảnh thực hiện:** luôn thay bằng ảnh `0.*` của **chính hồ sơ đang làm**; không giữ ảnh
  chép từ mẫu cũ. Không có ảnh mới → để trống (ghi đè ảnh trắng). **Giữ nguyên kích thước
  ảnh gốc** (không resample/giảm độ phân giải, không crop, không méo) — **chỉ thu gọn kích thước hiển
  thị cho vừa khung**, giữ tỉ lệ. Cách làm: `CONG-CU-VA-WORKFLOW.md` mục 4a.
- **Khách hàng yêu cầu không đưa hình ảnh:** không chèn ảnh; ở BCKQ bỏ khung ảnh trống và ghi ngay
  sau đầu mục ảnh: "Khách hàng yêu cầu không chụp/đưa hình ảnh vị trí lấy mẫu."
  Ở TBKQ xoá hẳn đầu mục ảnh cùng khung ảnh (`del-tables.ps1 -WithHeading`), không ghi chú thay thế.
- **Báo cáo cuối:** file nào đã tạo/sửa, ô nào để trống, điểm nào cần người dùng xác nhận.

---

## 4. Các bước thực hiện

### Bước 0 — Dữ liệu gốc (`0.*`)

- Đọc `0. BBLM.pdf` (biên bản giám định hiện trường) **trước tiên**: khách hàng, địa chỉ,
  ngày/giờ, số vị trí từng khu vực, ký hiệu mẫu, chỉ tiêu, người thực hiện.
- **Thiết bị sử dụng: KHÔNG lấy từ BBLM** đưa vào BCKQ và TBKQ. Ô/dòng "Thiết bị sử dụng" để `………`
  tô vàng cho người dùng tự nhập (kể cả khi BBLM có ghi thiết bị, kể cả khi mẫu cũ ghi sẵn thiết bị
  → xoá, thay bằng `………` tô vàng). Liệt kê các ô này trong báo cáo cuối.
- **Tên vị trí lấy mẫu: ưu tiên lấy theo BBLM.** BBLM là nguồn chuẩn. Chỉ khi hồ sơ nhiều mẫu mà
  BBLM **không thể hiện hết** (vd chỉ ghi gộp "10 mẫu", "1 → 10") thì mới lấy tiếp từ phiếu
  **Ghi chép** (BM3030.01). Ghi chép cũng không có/chưa điền → **yêu cầu người dùng bổ sung Ghi
  chép**, không tự suy tên vị trí từ tên ảnh hay từ hồ sơ mẫu.
- **Ký hiệu mẫu: ưu tiên ghi đúng theo BBLM** (vd BBLM ghi N1…N15 → ghi N1…N15), không lấy theo
  tên file ảnh hay ký hiệu của hồ sơ mẫu cũ (vd NU1…). BBLM không ghi ký hiệu → mới theo Ghi
  chép / quy ước mẫu, và nêu rõ trong báo cáo cuối.
- BBLM/Ghi chép là bản scan **chữ viết tay** → đọc ngay trong CLI bằng `pdf-img.js` +
  `img-crop.ps1` (`CONG-CU-VA-WORKFLOW.md` mục 5). Chữ không chắc → để trống + tô vàng, hỏi
  người dùng, không đoán.
- Phiếu KQTN dạng PDF khoá quyền copy → render bằng `pdf-render.ps1` rồi đọc ảnh.
- Chỉ đọc thêm file `0.*` nào thực sự cần. **Không chỉnh sửa** `0.*`.

### Bước 1 — Kết quả tính toán (`1.*`)

Thứ tự ưu tiên nguồn kết quả:

1. **Bản scan kết quả tính toán đã in/ký** → dùng thẳng, **không tính lại**.
2. **File Excel tính toán đã điền** → đọc kết quả, đối chiếu 2–3 giá trị mốc với `0.*`;
   lệch thì báo người dùng trước khi sửa. Giữ nguyên công thức và cấu trúc bảng tính.
3. **Chưa có gì** → mới tính từ dữ liệu thô (công thức: `CONG-CU-VA-WORKFLOW.md` mục 3).

**Cảnh báo:** file tính toán mang tên hồ sơ hiện tại vẫn có thể là **bản chép từ hồ sơ khác**
(mtime trước ngày đo là dấu hiệu). Luôn đối chiếu vài số với dữ liệu gốc trước khi dùng.

File tính toán Excel (khí nén `1. tinh toan tong dau.xlsx`, vi sinh không khí `0. Tinh Toan VSKK.xlsx`…):
thay **hết** vùng nhập liệu bằng số của chính hồ sơ này rồi đọc kết quả — ô nào, nguồn nào:
`CONG-CU-VA-WORKFLOW.md` mục 7.

### Bước 2 — Báo cáo kết quả (`/ThucHien`, tên `BCKQ - ...`, không đặt `2. BCKQ`)

- Vá từ file mẫu theo nguyên tắc mục 3.
- **"Người thực hiện" / "Người báo cáo"**: lấy theo tên trên BBLM của hồ sơ hiện tại, **chỉ ghi
  tên cuối**, nhiều người nối bằng ` + ` (vd "Nguyễn Văn Phát, Trần Minh" → `Phát + Minh`).
  Không giữ tên từ mẫu cũ.
  - **Người báo cáo:** luôn có thêm **`+ Minh`** sau tên người thực hiện (không trùng lặp nếu Minh đã có).
    Hồ sơ do **Toàn** hoặc **Nghĩa** thực hiện → `<tên> + Minh` (vd `Toàn + Minh`);
    các trường hợp còn lại → `<tên> + Minh + Nam` (vd `Quân + Minh + Nam`, `Huy + Minh + Nam`).
  - "Người thực hiện" luôn chỉ là tên trên BBLM, **không** thêm "Nam".
- Đổi tên file ngay khi bắt đầu vá (mục 2).

### Bước 3 — Chứng thư (`/TBKQ` và `/pdf`)

- Như Bước 2, cộng thêm:
- **Bắt buộc sửa cả header/footer** — mã hồ sơ, số chứng thư và ngày nằm ở `header1`…`header4`
  tuỳ mẫu (Bẫy 3). Bỏ sót là chứng thư mang mã của khách khác.
- **Mỗi mẫu một tờ riêng** khi khách yêu cầu, hoặc khi mẫu chứng thư chỉ có chỗ cho một mẫu:
  nhân bản đủ số bản, vá riêng từng bản. Các bản tách **giữ nguyên vị trí** ngang hàng trong
  thư mục hồ sơ, **không** tạo thư mục con riêng.
- Số chứng thư: theo mục 2 "Số chứng thư TBKQ" — 1 TBKQ ghi mã trần; nhiều TBKQ ghi `<mã hồ sơ>/<n>` theo thứ tự mẫu (bản tiếng Anh thêm `/E`, vd
  `Q326D111235/2` và `Q326D111235/2/E`).
- **Ngày ban hành TBKQ (Ngày chứng thư ở Header & Trang bìa)**:
  - Khi người dùng yêu cầu điền ngày hiện tại (hoặc ngày ban hành cụ thể) -> cập nhật đồng bộ cho cả **bản tiếng Việt** và **bản tiếng Anh**:
    + Bản tiếng Việt: `Số: <Số chứng thư> - Ngày dd/mm/yyyy` (vd: `Số: Q326D122266/1 - Ngày 19/10/2026`).
    + Bản tiếng Anh: `No: <Số chứng thư> – Date: dd/mm/yyyy` (vd: `No: Q326D122266/1/E – Date: 19/10/2026`).
    + Cập nhật tại **tất cả header** (header trang đầu và header phụ lục) và dòng số/ngày ở trang bìa, **gỡ bỏ `………` và gỡ bỏ hoàn toàn màu tô vàng (`highlight`)**.
  - Khi chưa có KQTN / chưa có ngày ban hành cụ thể -> ngày ban hành để `………` tô vàng (Bước 6).
- Bản chứng thư chỉ làm một phần chỉ tiêu (vd chỉ đo bụi) → **xoá hẳn các mục không thực hiện**
  cùng bảng ghi chú của chúng (`del-tables.ps1 -WithHeading`), không để bảng trống.
- **Thư mục xuất file chứng thư (TBKQ):** Toàn bộ file kết quả TBKQ (tiếng Việt, tiếng Anh, .docx và .pdf) được xuất trực tiếp vào **thư mục con `TBKQ` trong chính thư mục `Input` của hồ sơ đang thực hiện** (đường dẫn: `Input\<tên hồ sơ>\TBKQ\`, ví dụ: `Input\122266 - Phòng sạch\TBKQ\`, `Input\122266 - Khí nén\TBKQ\`). Tự động tạo thư mục `TBKQ` nếu chưa có.

### Bước 4 — Thư mục scan trình ký (`/TBKQ/Edoc`)

Hồ sơ **chưa có** thư mục `4.*` → tự tạo `4. Edoc\`, không hỏi lại.
Chép bản scan `0.*` của **chính hồ sơ này** vào, đặt tên `<6 số> - <tên>.pdf`.
Trong `4. Edoc\` không được để file của hồ sơ khác.

### Bước 5 — Kiểm tra cuối

1. Rà lại tên file theo mục 2 (`rename-done.ps1` bắt sót).
2. `final-check.ps1 -Code … -Old "<mã cũ>,<ngày cũ>,<tên khách cũ>,…"` → phải CLEAN.
3. `test-open.ps1 -Num <6 số>` — mở thử mọi docx bằng Word, bắt lỗi "unreadable content";
   số trang in ra phải khớp tổng trang ghi cứng ở footer (Bẫy 11).
4. Báo cáo cuối theo mục 3.

### Bước 6 — Hoàn tất khi có KQTN

Người dùng báo "đã có KQTN" → khung hồ sơ đã làm trước, chỉ còn điền kết quả và ngày.

1. Tìm phiếu KQTN (thường `0. KQ-<mã>-<số phiếu>.pdf`, hoặc trong thư mục edoc). PDF khoá →
   `pdf-render.ps1`. **Đếm số mẫu trên KQTN, đối chiếu BBLM/Ghi chép, kể cả mẫu blank** (Bẫy 16).
2. Nhập kết quả vào file tính toán Excel (`CONG-CU` mục 7), đọc kết quả + phân loại, đối chiếu scan.
3. Điền kết quả, phân loại, kết luận vào `3.*` (VN và EN) và các ô KQTN/KQTT ở `2.*`; không đạt → in đậm.
   - **KQTN ghi KPH → cột kết quả TBKQ nhập giá trị LOD hoặc LOQ** (theo phiếu KQTN) dạng
     `< LOD`, ví dụ `< 0,01` (EN: `< 0.01`); **không nhập chữ `KPH`** vào cột kết quả của TBKQ.
   - **Kết quả vi sinh ghi giống như KQTN** (vd `2,1 × 10²`, `1,3 × 10⁶`, `< 10`, `< 1`), không đổi sang
     số nguyên (`210`, `1 300 000`). Bản EN chỉ đổi dấu thập phân (`2.1 × 10²`). Kết quả phải quy đổi qua
     file tính toán (VSKK CFU/đĩa → CFU/m³) thì ghi theo file tính toán.
4. Ngày — một quy ước cho cả hồ sơ:
   - **Ngày ban hành TBKQ** = ngày trên phiếu KQTN (nhiều phiếu → ngày muộn nhất), ghi ở **mọi**
     header có ngày (header trang đầu và header phụ lục, VN và EN — Bẫy 3). Mẫu ghi sẵn ngày khác →
     vẫn thay, nêu trong báo cáo cuối.
   - BCKQ: số phiếu + ngày KQTN, ngày nhận KQTN, ngày kết thúc hồ sơ, ngày báo cáo = ngày KQTN.
   - "Ngày yêu cầu cấp KQTN" = "Ngày dự kiến trả kết quả" trên PYC (thư mục `*YCTN`).
   - "Ngày gửi mẫu" = "Ngày nhận mẫu" trên phiếu KQTN.
   - Phiếu tính toán chưa có số → nối tiếp số đã có (`/1`, `/2` → `/3`…), nêu trong báo cáo.
5. Gỡ tô vàng mọi ô vừa điền (`apply-all` `unhl`); rà lại còn `………` nào (dump + grep).
6. Chép phiếu KQTN vào thư mục Edoc (`<6 số> - KQTN <số phiếu>.pdf`), rồi làm lại Bước 5.

---

## 5. Checklist bàn giao

- [ ] Đúng thư mục hồ sơ theo mã (kể cả `0.VIMCERTS078` nếu là hồ sơ Vimcerts)
- [ ] Đã đọc BBLM: khách hàng / ngày giờ / vị trí / ký hiệu mẫu / chỉ tiêu / người thực hiện
- [ ] Số liệu lấy từ bản scan hoặc file tính toán **của đúng hồ sơ này** (đã đối chiếu)
- [ ] "Người thực hiện" / "Người báo cáo" theo BBLM; người báo cáo: Toàn/Nghĩa → `<tên> + Minh`, còn lại → `<tên> + Minh + Nam`
- [ ] `2.*` và `3.*` đúng format, sạch dấu vết hồ sơ cũ (kể cả header/footer và phần diễn giải)
- [ ] `3.*` đủ số tờ theo số mẫu; số chứng thư `/1`, `/2`… (EN thêm `/E`)
- [ ] Kết quả không đạt đã in đậm; giá trị đạt không còn in đậm sót từ mẫu cũ
- [ ] Ảnh vị trí, sơ đồ mapping và ảnh thực hiện đã tăng/giảm đúng số lượng của hồ sơ hiện tại (hoặc để trống/xoá theo quy định), nguyên ảnh gốc chỉ thu gọn hiển thị vừa khung, không lẫn ảnh cũ hay thừa khung trống
- [ ] Bảng kết quả có câu kết luận đi kèm: câu kết luận (`Đánh giá:` / `Assessment:`) nằm cùng trang với bảng, không bị rớt lẻ sang trang mới
- [ ] Cụm từ đánh giá trong câu kết luận ("phù hợp", "không phù hợp", "comply", "complied", "did not comply") đã được in đậm và in nghiêng
- [ ] Số chứng thư ở toàn bộ các header khớp nhau; Số hồ sơ trong thân văn bản (Table 0) ghi mã trần, KHÔNG có đuôi /1, /2, /E
- [ ] Toàn bộ các văn bản TBKQ có đường kết thúc đôi (thinThick, 4.5pt, centered) ở cuối cùng của văn bản, căn giữa trang, không để sót đường nối mồ côi ở giữa văn bản
- [ ] "Thiết bị sử dụng" ở BCKQ/TBKQ để `………` tô vàng (không lấy từ BBLM, không giữ của mẫu cũ)
- [ ] Cột "Ngưỡng giới hạn cho phép": Nước thải/Chất thải/Bùn thải → tô vàng cả cột;
      Nước uống/Nước sinh hoạt → không tô
- [ ] Mọi file `1.*`, `2.*`, `3.*` đã thêm đuôi mã hồ sơ
- [ ] `4. Edoc` (chưa có thì tự tạo) đủ bản scan và **chỉ** có file của hồ sơ này
- [ ] `final-check.ps1` CLEAN và `test-open.ps1` mở được hết
- [ ] Đã nêu rõ ô để trống và điểm cần người dùng xác nhận
