# Antigravity Workspace Rules - Làm Kết Quả Quan Trắc Môi Trường

> **LỖI THỜI (ghi chú Claude, cập nhật 30/10/2026):** file này là **bản chép cũ (25/10/2026) của `AGENTS.md`**, không còn được cập nhật. **Không làm theo file này** — dùng `AGENTS.md` (phiên bản 30/10/2026 22:45 trở về sau) và `PROJECT_CONTEXT.md` ở thư mục gốc. Lưu ý: quy tắc Z: đã đổi — **được lưu hồ sơ lên Z** theo `AGENTS.md` mục 0.

## 1. Cấu trúc Thư mục Hồ sơ và Nguyên tắc Phân định Lưu trữ Kết quả

Áp dụng làm **nguyên tắc chung bất biến cho tất cả hồ sơ** khi xử lý dữ liệu đầu vào trong thư mục `Input\<tên hồ sơ>` (ví dụ: `Input\043915`, `Input\131638`, `Input\132918`, `Input\122266 - Phòng sạch`...):

- **Bước 2 — Thư mục `/ThucHien` (`Input\<tên hồ sơ>\ThucHien\`):**
  - Chứa toàn bộ các file kết quả nội bộ của **Bước 2** trong quy trình, bao gồm:
    + Báo cáo kết quả giám định nội bộ (BM3030.03): file `.docx` tên chuẩn `BCKQ - <stt> - <loại> - <Mã_hồ_sơ>.docx` (**tuyệt đối không đặt tiền tố `2. `**).
    + Các file bảng tính Excel tính toán kết quả nội bộ: `<Mã_hồ_sơ> - TinhToan....xlsx` (ví dụ: `Q326D132918 - TinhToanBSC.xlsx`, `Q326D043915 - TinhToanFumehood.xlsx`).

- **Bước 3 — Thư mục `/TBKQ` (`Input\<tên hồ sơ>\TBKQ\`):**
  - Chứa toàn bộ các file chứng thư trả khách hàng của **Bước 3** trong quy trình, bao gồm:
    + Thông báo kết quả (TBKQ) bản tiếng Việt: `TBKQ - viet - ... - <Mã_hồ_sơ>.docx`.
    + Thông báo kết quả (TBKQ) bản tiếng Anh: `TBKQ - en - ... - <Mã_hồ_sơ>.docx`.
    + Tem kiểm định / dán tủ: `tem tu - ... - <Mã_hồ_sơ>.docx` (hoặc `<Mã_6_số> - tem tu.docx`).
    + Các sơ đồ layout vi tính và hình ảnh đồ họa: `layout_downflow.png`, `_diagrams/`...

- **Tách scan PDF — Thư mục `/TBKQ/Edoc` (`Input\<tên hồ sơ>\TBKQ\Edoc\`):**
  - Chứa toàn bộ các file PDF tách tự động từ bản scan gốc tài liệu hiện trường:
    + `<Mã_6_số> - BBLM.pdf`: Biên bản giám định hiện trường (BM3030.02).
    + `<Mã_6_số> - Ghi Chep.pdf`: Ghi chép giám định hiện trường (BM3030.01) và các phụ lục in nhiệt.
  - Tên thư mục chuẩn là `Edoc` (viết liền, không có khoảng cách).

- **Tập trung file PDF xuất ra — Thư mục `/pdf` (`Input\<tên hồ sơ>\pdf\`):**
  - **Để chung toàn bộ các file PDF xuất ra trong duy nhất 1 thư mục con `/pdf`** nằm ở thư mục gốc của hồ sơ (ngang hàng với `/ThucHien` và `/TBKQ`).
  - Gồm toàn bộ file PDF xuất ra từ DOCX và XLSX:
    + `BCKQ - <stt> - <loại> - <mã>.pdf`
    + `TBKQ - viet - ...pdf`, `TBKQ - en - ...pdf`
    + `tem tu - ...pdf` (nếu có)
    + **Các file PDF xuất ra từ file bảng tính Excel tính toán:** `<Mã_hồ_sơ> - TinhToan....pdf` (ví dụ: `Q326D131638 - TinhToanBSC.pdf`, `Q326D043915 - TinhToanFumehood.pdf`, `Q326D133266- KN1 - TinhToanKhiNen.pdf`...).
  - Không để các file PDF xuất ra rải rác trong `ThucHien/` hay `TBKQ/`.

- **Tuyệt đối KHÔNG lưu file kết quả ở thư mục gốc `Input\<tên hồ sơ>\`**:
  - Thư mục gốc `Input\<tên hồ sơ>\` chỉ lưu trữ tài liệu đầu vào (ảnh chụp hiện trường `Photo/`, các file scan gốc).
  - Toàn bộ thư mục `ThucHien`, `TBKQ`, `TBKQ/Edoc`, `pdf` nếu chưa có sẽ được hệ thống tự động tạo mới (`os.makedirs`).
  - Sao chép đồng bộ sang thư mục `Output\` ở gốc workspace và ổ mạng Z: theo đúng cấu trúc tương ứng.



## 2. Quy tắc Số chứng thư và Ngày ban hành (Header & Trang bìa)



- **Số chứng thư:** Hồ sơ 1 bản ghi mã trần (vd `Q326D122266`), hồ sơ nhiều bản ghi thêm `/1`, `/2`... (bản tiếng Anh thêm `/E`, vd `Q326D122266/1/E`).



- **Đồng bộ Số chứng thư giữa các header:** Toàn bộ các header trong cùng một tài liệu (header trang đầu Section 0 và header phụ lục Section 1) bắt buộc phải có cùng một Số chứng thư, tuyệt đối không được lệch nhau.



- **Quy tắc Số hồ sơ trong thân văn bản (Body Text & Bảng biểu):**



  - Số hồ sơ (Số giấy tiếp nhận / Inspection application no.) trong thân văn bản (Table 0 dòng 4 hoặc các đoạn văn) **bắt buộc ghi mã trần (mã gốc), tuyệt đối KHÔNG có đuôi `/1`, `/2`, `/3`, `/E`, v.v.** (Ví dụ chuẩn: `Q326D122266`, không ghi `Q326D122266/2` hay `Q326D122266/2/E`).



  - **Không in đậm mã hồ sơ** (`bold=False`) tại ô Table 0 dòng 4 trong thân văn bản.



  - **Định dạng Mã hồ sơ và Ngày trong Table 0 dòng 4:** Mã hồ sơ và ngày nằm liền mạch trên cùng một dòng/ô hợp nhất, cách nhau đúng 1 khoảng space:



    + Tiếng Việt: `<bare_code> ngày <date>` (Ví dụ chuẩn: `Q326D132918 ngày 20/10/2026`).



    + Tiếng Anh: `<bare_code> dated on <date>` (Ví dụ chuẩn: `Q326D132918 dated on 20/10/2026`).



    + Tuyệt đối không để ngày bị tách rời sang cột xa hoặc có khoảng trống trắng lớn.



  - Đuôi `/1`, `/2`, `/E` CHỈ dành riêng cho Số chứng thư trên Header và tiêu đề chứng thư.



- **Ngày ban hành TBKQ:**



  - Định dạng chuẩn: `dd/mm/yyyy` (vd: `20/10/2026`).



  - Cú pháp tiếng Việt: `Số: <Số chứng thư> - Ngày dd/mm/yyyy`



  - Cú pháp tiếng Anh: `No: <Số chứng thư> – Date: dd/mm/yyyy`



  - Khi cập nhật ngày (ngày hiện tại hoặc theo KQTN), cập nhật đồng bộ ở **mọi header** (header trang đầu và header phụ lục) và dòng số/ngày ở trang bìa, **gỡ bỏ hoàn toàn `………` và màu tô vàng (`highlight`)**.







## 3. Quy chuẩn Định dạng Bảng, Font chữ và Đánh giá (Áp dụng cho toàn bộ TBKQ)



- **Cỡ chữ (Font size) chuẩn trong TBKQ:**



  - **Font size 13 (Times New Roman):** Áp dụng cho **toàn bộ thân văn bản** bao gồm: các đoạn văn, tiêu đề các mục lớn/nhỏ (Mục 1, 2, 2.1, 2.2...), câu kết luận đánh giá (`Đánh giá:` / `Assessment:`), nội dung và thông tin trong Table 0.



  - **Font size 13 (in nghiêng):** Áp dụng cho **các ghi chú/chú thích khác** ngoài bảng kết quả, bao gồm: chú thích hình ảnh (ví dụ: `Hình 1: Vị trí chụp hình thực tế...`), chú thích sơ đồ layout (ví dụ: `Sơ đồ 1: Sơ đồ đo...`).



  - **Font size 10 (Times New Roman):** Áp dụng riêng cho **các bảng ghi chú nằm ngay trực tiếp dưới bảng kết quả** (`Ghi chú:` / `Note:` của các bảng đo đạc/lấy mẫu).



- **AutoFit to Window cho Bảng chứa Layout và Bảng chứa Hình ảnh:**



  - Các bảng chứa hình ảnh thực hiện và bảng chứa sơ đồ bố trí điểm đo (layout) **bắt buộc thiết lập AutoFit là `AutoFit to Window`** (`table.AutoFitBehavior(2)` / `wdAutoFitWindow` qua Word COM và OpenXML: `<w:tblW w:w="5000" w:type="pct"/>`, `<w:tblLayout w:type="autofit"/>`, không viền, các ô chia đều theo tỷ lệ phần trăm), đảm bảo hiển thị vừa vặn, cân đối trong lề trang in và không bị tràn lề.



- **Quy chuẩn chèn ảnh dùng grid 2 cột (2 ảnh trái phải - Áp dụng tất cả trường hợp):**



  - **Kích thước ngang của ảnh:** Khi chèn ảnh sử dụng bảng/grid 2 cột (2 ảnh trái phải cạnh nhau), kích thước ngang (`width`) của mỗi ảnh **bắt buộc để trong khoảng 7 cm – 8 cm** (chuẩn khuyến nghị: **7,5 cm** / `Cm(7.5)`).



  - **Tuyệt đối không crop ảnh:**



    + Giữ nguyên 100% tỷ lệ gốc (aspect ratio) của hình ảnh, không cắt xén, không làm méo hình.



    + Khi chèn ảnh qua code, **chỉ thiết lập thuộc tính `width`** (ví dụ: `r.add_picture(path, width=Cm(7.5))`), **tuyệt đối không đặt thuộc tính `height`** để Word/python-docx tự động tính toán chiều cao tương ứng theo tỷ lệ ảnh gốc.



    + Xóa bỏ toàn bộ các thuộc tính cắt xén `<a:srcRect>` (`sr.attrib.clear()`) trên đối tượng ảnh trong Word, đảm bảo hình ảnh hiển thị toàn vẹn.



    + Tuyệt đối không crop ảnh trong các bước xử lý hình ảnh (PIL/Pillow).



  - **Áp dụng cho toàn bộ các trường hợp:** Áp dụng trên 100% các loại hồ sơ (BSC, LAF, Fume Hood, Phòng sạch, Khí nén...) và toàn bộ các văn bản (TBKQ tiếng Việt & tiếng Anh, BCKQ nội bộ), bao gồm cả ảnh chụp hiện trường và sơ đồ đo đạc/layout.



  - **Định dạng bảng ảnh 2 cột:** Căn giữa trang (`WD_TABLE_ALIGNMENT.CENTER`), không viền, AutoFit to Window (`set_table_autofit_window`, mỗi ô chiếm 50% = 2500 twips/pct), các hàng chống tách trang (`add_cant_split(row)`), hình ảnh căn giữa ô (`WD_ALIGN_PARAGRAPH.CENTER`), nếu có chú thích thì đặt ngay dưới ảnh với font 13 in nghiêng.



- **Quy chuẩn chèn ảnh dùng grid 3 cột (khi có từ 5 ảnh trở lên, ví dụ BSC, LAF, Fume Hood...):**



  - **Kích thước ngang của ảnh:** Khi chèn ảnh sử dụng bảng/grid 3 cột (3 ảnh mỗi hàng), kích thước ngang (`width`) của mỗi ảnh **bắt buộc để trong khoảng 4 cm – 5 cm** (chuẩn khuyến nghị: **4,8 cm** / `Cm(4.8)`).



  - **Tuyệt đối không crop ảnh:** Chỉ thiết lập thuộc tính `width` (ví dụ: `r.add_picture(path, width=PHOTO_WIDTH_3COL)`), tuyệt đối không đặt thuộc tính `height`, xóa bỏ toàn bộ `<a:srcRect>`.



  - **Định dạng bảng ảnh 3 cột:** Căn giữa trang (`WD_TABLE_ALIGNMENT.CENTER`), không viền, AutoFit to Window (`set_table_autofit_window`, mỗi ô chiếm ~33.3% = 1666 twips/pct), các hàng chống tách trang (`add_cant_split(row)`), hình ảnh căn giữa ô (`WD_ALIGN_PARAGRAPH.CENTER`).



- **Merge ô STT cho các chỉ tiêu có nhiều dòng con (Table Kết quả đo):** Đối với các chỉ tiêu có dòng tiêu đề và nhiều dòng con bên dưới (ví dụ: Đo độ ồn, Mức rò rỉ lọc HEPA, Kiểm tra hình thái dòng khí, Nồng độ hạt bụi...), ô STT ở cột đầu tiên (No. / STT) bắt buộc phải merge dọc (vertical merge `<w:vMerge>`) từ dòng đầu xuống hết các dòng con của chỉ tiêu đó, căn giữa theo cả chiều ngang và chiều đứng (`valign="center"`, `align=WD_ALIGN_PARAGRAPH.CENTER`), không để ô STT của các dòng con bị trống hoặc có đường kẻ ngăn cách.



- **Chiều cao dòng (Row Height) cho Bảng kết quả:**



  - Toàn bộ các bảng chứa kết quả đo đạc (nồng độ bụi/hạt, vận tốc gió, vi khí hậu, độ rọi, dòng khí, khí nén, thông số phòng...) phải có chiều cao dòng (`row height`) tối thiểu là **0,6 cm**, ưu tiên sử dụng trong khoảng **0,7 cm – 0,8 cm** (`w:hRule="atLeast"`, tương đương 400 – 455 twips, ví dụ: 440 twips = 0,78 cm).



- **Bảng và Kết luận:** Bảng kết quả có câu kết luận đi kèm (`Đánh giá:` / `Assessment:`) thì câu kết luận và bảng `Ghi chú:` / `Note:` phải nằm cùng trang với dòng cuối của bảng (`keepNext`), không để câu kết luận rớt sang trang mới.



- **Lặp lại tiêu đề bảng khi sang trang (Repeat Header Rows):** Các bảng kết quả/dữ liệu dài khi ngắt sang trang mới **bắt buộc phải lặp lại (các) dòng tiêu đề cột** (`<w:tblHeader/>` trong `<w:trPr>`), giúp người đọc dễ dàng theo dõi thông số các cột trên từng trang.



- **Cụm từ đánh giá:** Phải **in đậm và in nghiêng** (`bold=True, italic=True`):



  - Tiếng Việt: ***phù hợp***, ***không phù hợp***



  - Tiếng Anh: ***comply***, ***complied***, ***did not comply***



- **Chỉ tiêu không thực hiện:** Xóa bỏ hoàn toàn mục và bảng liên quan (kèm bảng ghi chú của nó), không để bảng trống.



- **Căn chỉnh ngang hàng chữ ký người phê duyệt và giám định viên (Table 0):**



  - Khi người phê duyệt bên phải có 2 dòng chức danh (`KT. GIÁM ĐỐC` và `PHÓ GIÁM ĐỐC`, hoặc `FOR DIRECTOR` và `VICE DIRECTOR`):



  - Chức danh bên trái (`GIÁM ĐỊNH VIÊN` hoặc `INSPECTOR`) **bắt buộc phải nằm ở dòng thứ 2 để ngang hàng tuyệt đối với `PHÓ GIÁM ĐỐC` / `VICE DIRECTOR`** (dòng 1 bên trái để trống/spacer), không để `GIÁM ĐỊNH VIÊN` ngang hàng với `KT. GIÁM ĐỐC`.



  - **Khoảng trống ký tên chuẩn 3 cm:** Đoạn chứa họ tên người ký (`p_name`) bắt buộc thiết lập khoảng cách trước là đúng **3 cm** (`space_before = Cm(3.0)` hoặc `Pt(85.05)`, tương đương ~1701 twips) để đảm bảo không gian ký tươi và đóng dấu tròn rộng rãi, đồng đều và cân đối tuyệt đối giữa hai bên.



  - Họ tên người ký ở hai bên phải thẳng hàng ngang nhau.







## 4. Kiểm tra chất lượng (QA)



- Sử dụng công cụ `Core\Tools\qa_check.py`:



  - `python Core\Tools\qa_check.py final-check <path> [--code <Mã_hồ_sơ>]`: Tự động kiểm tra:



    + Tính nhất quán và đồng bộ của Số chứng thư ở toàn bộ các header.



    + Đảm bảo số hồ sơ trong thân văn bản là mã trần, không bị dính đuôi `/1`, `/2`, `/E`.



    + Phát hiện mã hồ sơ lạ/cũ còn sót lại từ file mẫu.



    + Kiểm tra 0 placeholder sót lại, 0 highlight vàng sót lại, giữ nhãn ô Table 0 nguyên vẹn.



    + Đảm bảo các bảng kết quả/dữ liệu có thuộc tính lặp lại hàng tiêu đề (`<w:tblHeader/>`).



    + Căn chỉnh ngang hàng chức danh chữ ký Table 0: `INSPECTOR` ngang hàng tuyệt đối với `VICE DIRECTOR`, và `GIÁM ĐỊNH VIÊN` ngang hàng với `PHÓ GIÁM ĐỐC` (dòng 2).



    + Đảm bảo khoảng trống ký tên Table 0 đạt chuẩn 3 cm (`Cm(3.0)` / ~1701 twips).



    + Đảm bảo ảnh trong bảng grid 2 cột có kích thước ngang 7 - 8 cm và tuyệt đối không crop ảnh.



    + Đảm bảo các chỉ tiêu có nhiều dòng con được merge ô STT dọc (<w:vMerge>) và căn giữa hoàn toàn theo cả chiều dọc và chiều ngang.



    + Đảm bảo kết quả đạt `[CLEAN]`.



  - `python Core\Tools\qa_check.py test-open <file.docx>`: Kiểm tra mở qua Word COM và xác nhận số trang chuẩn xác.







## 5. Quy tắc Đường kết thúc văn bản (Ending Line)



- Toàn bộ các báo cáo/chứng thư kết quả (TBKQ) khi kết thúc văn bản **bắt buộc luôn có đường kết thúc đôi (`thinThick` compound line, 4.5pt strokeweight, màu đen)** nằm ở cuối cùng của văn bản, căn giữa trang (`align center` / `wp:align center` qua hai lề `margin`).



- **Đường kẻ nằm ngang tuyệt đối & Không đè chữ:** Chiều cao trục Y phải bằng 0 (`cy="0"` / `cy=0`) để đường kẻ nằm ngang thẳng tuyệt đối, không bị xiên lệch. Đường kết thúc được đặt trong đoạn kết thúc riêng biệt (dedicated paragraph không chứa text) nằm ngay dưới câu kết luận cuối cùng với khoảng cách thẳng đứng phù hợp (~7-14pt dưới dòng chữ cuối cùng), **tuyệt đối không để đường kẻ đè/chắn lên bất kỳ chữ nào** trong câu đánh giá.



- Tuyệt đối không để sót các đường nối mồ côi (`orphaned connector`) ở các đoạn văn giữa văn bản.



- **Không áp dụng Ending Line cho BCKQ:** Tuyệt đối **KHÔNG thêm đường Ending Line vào Báo cáo kết quả giám định nội bộ (BCKQ)**.



- Công cụ QA Check (`qa_check.py`) tự động kiểm tra sự hiện diện của đường kết thúc chuẩn (`thinThick`, centered) ở cuối tài liệu và phát hiện đường nối mồ côi.







## 6. Quy tắc Đồng bộ Tên vị trí, Thiết bị và Số liệu (Tiếng Việt & Tiếng Anh)



- **Đồng bộ tên vị trí thực hiện (quan trắc):** Tên các vị trí thực hiện/đo đạc/lấy mẫu giữa bản tiếng Việt và bản tiếng Anh **bắt buộc phải tương ứng và giống nhau** (đồng nhất số lượng, thứ tự vị trí và mô tả vị trí theo bản dịch chuẩn, ví dụ: VT1, VT2... hoặc tên phòng/vị trí cụ thể). Tuyệt đối không để lệch tên vị trí, thiếu vị trí hoặc thứ tự vị trí khác nhau giữa hai bản.



- **Đồng bộ danh mục thiết bị:** Danh mục thiết bị sử dụng giữa các bảng tiếng Việt và tiếng Anh **bắt buộc phải giống nhau** (cùng chủng loại thiết bị, số lượng và quy cách tương ứng theo bản dịch chuẩn). Tuyệt đối không để bản tiếng Việt có thiết bị này mà bản tiếng Anh lại thiếu hoặc liệt kê thiết bị khác.



- **Đồng bộ số liệu đo đạc:** Toàn bộ các số liệu điền vào giữa bảng tiếng Việt và bảng tiếng Anh **phải hoàn toàn giống nhau** về giá trị đo đạc.



- **Quy chuẩn dấu thập phân:**



  - **Bản tiếng Việt:** Bắt buộc sử dụng **dấu phẩy (`,`)** khi biểu diễn số thập phân (ví dụ: `23,5`, `0,1`, `15,0`).



  - **Bản tiếng Anh:** Bắt buộc sử dụng **dấu chấm (`.`)** khi biểu diễn số thập phân (ví dụ: `23.5`, `0.1`, `15.0`).



- **Đồng bộ ký hiệu vị trí khi đánh số (KN1, KN2... / VT1, VT2...):**



  - Khi một hồ sơ có các vị trí được đánh số phân biệt (ví dụ: `KN1`, `KN2`, `VT1`, `VT2`...), vị trí đầu tiên **bắt buộc phải ghi rõ ràng là `KN1` (hoặc `VT1`), tuyệt đối KHÔNG ghi chung chung là `KN` (hoặc `VT`)**.



  - Đảm bảo tính nhất quán 100% giữa:



    + Ký hiệu được đánh dấu/chú thích trên hình ảnh hiện trường (ví dụ: trong ảnh có mũi tên chỉ `KN1`).



    + Chú thích ảnh (caption): `Vị trí lấy mẫu KN1` (không được ghi `Vị trí lấy mẫu KN`).



    + Bảng thông tin vị trí đo (Table 1 trong TBKQ & BCKQ).



    + Các bảng kết quả chi tiết (Hàm lượng dầu, Điểm sương, Nồng độ hạt bụi).



    + Bảng tổng hợp kết quả KQTT trong BCKQ (Table 3).



    + File bảng tính Excel tính toán (`Q326D... - KN1 - TinhToanKhiNen.xlsx` ô H3 phải là `KN1`).







## 7. Quy tắc Báo cáo kết quả giám định nội bộ (BCKQ - BM3030.03)



- **Mục đích & Biểu mẫu:**



  - Báo cáo kết quả giám định nội bộ (BCKQ) sử dụng biểu mẫu `BM3030.03` (Hiệu lực: `01/8/2025`).



  - Quy ước đặt tên file chuẩn (Không đặt tiền tố `2. `): `BCKQ - <stt> - <loại> - <Mã_hồ_sơ>.docx` (Ví dụ chuẩn: `BCKQ - 1 - phong sach - Q326D122266.docx`, `BCKQ - 2 - khi nen - Q326D122266.docx`, `BCKQ - 1 - fumehood - Q326D043915.docx`, `BCKQ - 1 - bsc - Q326D132918.docx`). Tuyệt đối không thêm tiền tố `2. `.



- **Quy chuẩn Bảng 0 của Báo cáo kết quả (Table 0 Formatting):**



  - **Sử dụng font chữ 13 (Times New Roman 13pt):** Áp dụng cho toàn bộ các ô, nhãn mục và nội dung điền trong Bảng 0 (Mã hồ sơ, Đơn vị yêu cầu, Địa chỉ, Nội dung yêu cầu, Quá trình thực hiện, Địa điểm thực hiện, Thời gian thực hiện, Thiết bị sử dụng, Người thực hiện).



  - **Format bảng Align Top:** Toàn bộ các ô (`cell`) trong Bảng 0 bắt buộc phải thiết lập căn lề theo phương đứng là căn đỉnh lề trên (**Align Top** / `<w:vAlign w:val="top"/>`), tuyệt đối không để căn giữa (`center`) hoặc mặc định lệch hàng.



  - **Thêm khoảng cách đoạn văn (Spacing Before & After):** Toàn bộ các đoạn văn (`paragraph`) bên trong các ô của Bảng 0 bắt buộc phải thiết lập:



    + **Layout / Spacing Top (Before): 3pt** (`space_before = Pt(3)` / `<w:spacing w:before="60"/>`).



    + **Layout / Spacing After: 3pt** (`space_after = Pt(3)` / `<w:spacing w:after="60"/>`).



    + Khoảng cách dòng đơn (`line_spacing = 1.0` / `<w:line="240" w:lineRule="auto"/>`).



  - **Áp dụng đồng bộ cho toàn bộ form báo cáo kết quả:** Bắt buộc áp dụng quy tắc định dạng Bảng 0 này trên 100% các form BCKQ (Phòng sạch, Khí nén, Tủ an toàn sinh học BSC, Tủ LAF, Tủ hút, v.v.).



- **Sử dụng Font chữ 13 cho toàn bộ Báo cáo kết quả (Global Font Size 13):**



  - **Toàn bộ văn bản và bảng biểu trong BCKQ bắt buộc sử dụng thống nhất font Times New Roman, cỡ chữ 13pt (`Pt(13)`):**



    + Toàn bộ thân văn bản: Tiêu đề các mục lớn/nhỏ (Mục 1, 2, 3, 4, 4.6, 5...), các đoạn văn mô tả khu vực, câu kết luận/đánh giá, các dòng tham chiếu kết quả.



    + Toàn bộ bảng biểu: Table 0 (thông tin chung), Table 1 (danh mục vị trí/thiết bị/nhiệm vụ), Table 2 (ghi chú), Table 3, Table 4, Table 5 (đánh giá).



    + Toàn bộ phần Báo cáo của GĐV về thời gian thực hiện: Ngày kết thúc hồ sơ thực tế, Trễ hạn, Thời gian trễ hạn, Nguyên nhân trễ hạn, Bất thường trong giám định, Ngày báo cáo, Người báo cáo.



    + Tuyệt đối không để sót các đoạn văn hoặc bảng dùng font 10pt, 11pt hay 12pt trong BCKQ.



- **Vị trí lưu trữ bắt buộc:**
  - **File `.docx`:** Nằm trong thư mục Bước 2 `/ThucHien`: `Input\<tên hồ sơ>\ThucHien\BCKQ - <stt> - <loại> - <Mã_hồ_sơ>.docx`.
  - **File `.pdf` xuất ra:** Nằm trong thư mục `/pdf`: `Input\<tên hồ sơ>\pdf\BCKQ - <stt> - <loại> - <Mã_hồ_sơ>.pdf`.
  - Đồng bộ sao chép sang thư mục `Output\` ở workspace và thư mục hồ sơ trên ổ mạng (`Z:\...\<Mã_hồ_sơ>...\ThucHien\` hoặc `<Mã_6_số> - Thuc hien\`).



- **Quy tắc Người thực hiện và Người báo cáo:**



  - **Người thực hiện:** Lấy từ Biên bản lấy mẫu / thực hiện (BBLM), ghi tên giám định viên trực tiếp đo đạc tại hiện trường (ví dụ: `Phát`, `Toàn`, `Nghĩa`...).



  - **Người báo cáo:**



    + Nếu người thực hiện là `Toàn` hoặc `Nghĩa`: Người báo cáo ghi `<Tên> + Minh` (Ví dụ: `Toàn + Minh`, `Nghĩa + Minh`).



    + Đối với các giám định viên khác (như `Phát`, v.v.): Người báo cáo ghi `<Tên> + Minh + Nam` (Ví dụ: `Phát + Minh + Nam`).



- **Đồng bộ nội dung, thời gian và số liệu với TBKQ:**



  - **Số báo cáo:** Ghi theo số chứng thư của hồ sơ (`Q326D.../1`, `Q326D.../2`...).



  - **Thời gian thực hiện:** Ghi ngày và khung giờ đo đạc thực tế (ví dụ: `Ngày 06/10/2026 từ 08:00 đến 17:00`).



  - **Thiết bị sử dụng:** Đồng nhất 100% với danh mục thiết bị đã ghi trên TBKQ.



  - **Khu vực / Vị trí thực hiện:** Liệt kê đầy đủ các phòng/điểm đo cùng diện tích, cấp sạch và các chỉ tiêu thực tế thực hiện.



  - **Hình ảnh & Sơ đồ:**



    + Chèn ảnh chụp thực tế tại hiện trường từ thư mục ảnh của hồ sơ (`Photo Thuc Hien\`, `Vi tri khi nen\`), tối ưu kích thước ảnh (~600x450 hoặc 800x600 px) để tài liệu gọn gàng, sắc nét.



    + Sơ đồ đo đạc/lấy mẫu hiển thị đúng vị trí và bố cục của hồ sơ.



  - **Tham chiếu kết quả (Mục 4.6 hoặc mục tương ứng trong BCKQ):**



    + **Tuyệt đối KHÔNG viện dẫn TBKQ hoặc Chứng thư giám định trong BCKQ:**



      * Theo quy trình thực hiện tại QUATEST 3, Báo cáo kết quả giám định nội bộ (BCKQ - BM3030.03) là tài liệu được lập và trình duyệt **TRƯỚC** "Chứng thư giám định" (Thông báo kết quả - TBKQ). Lúc lập BCKQ thì chưa có Chứng thư giám định / TBKQ.



      * Do đó, trong toàn bộ BCKQ (cả thân văn bản và bảng biểu), **BẮT BUỘC CHỈ VIỆN DẪN CÁC PHIẾU TÍNH TOÁN KẾT QUẢ / PHIẾU GHI CHÉP**, tuyệt đối **KHÔNG ĐƯỢC VIỆN DẪN TBKQ HAY CHỨNG THƯ GIÁM ĐỊNH**.



      * **Bảng tổng hợp KQTT (Khí nén - Table 3):** Cột KQTT ghi rõ theo từng vị trí: `Xem tại phiếu Kết quả tính toán số <Mã_hồ_sơ> - <Ký_hiệu> ngày dd/mm/yyyy` (Ví dụ: `Xem tại phiếu Kết quả tính toán số Q326D133266 - KN1 ngày 20/10/2026`). Tuyệt đối không ghi `TBKQ số...`.



      * **Chỉ tiêu nồng độ hạt bụi:** `Xem tại phiếu Tính toán kết quả nồng độ hạt bụi số <Mã_hồ_sơ> ngày dd/mm/yyyy`.



      * **Chỉ tiêu điểm sương:** `Xem tại phiếu Tính toán kết quả nhiệt độ điểm sương số <Mã_hồ_sơ> ngày dd/mm/yyyy`.



      * **Tủ an toàn sinh học (BSC), Tủ hút, Tủ LAF:** `Xem chi tiết tại phiếu Tính toán các thông số tủ số <Mã_hồ_sơ> ngày dd/mm/yyyy.` (không kèm `; TBKQ số...`).



      * **Phòng sạch:** Liệt kê theo từng chỉ tiêu tới phiếu tính toán tương ứng (nồng độ bụi: `phiếu Tính toán kết quả nồng độ hạt bụi...`, lưu lượng/HEPA: `phiếu Tính toán vận tốc gió và lưu lượng qua lọc HEPA...`, chênh áp/vi khí hậu: `phiếu Ghi chép giám định tại hiện trường...`), không ghi `Thông báo kết quả số...`.



    + **Quy tắc ghi nhận Phiếu kết quả đo độ ồn (và các chỉ tiêu có Phiếu đo riêng từ máy/phần mềm):**



      * **Tuyệt đối KHÔNG ghi cụt lủn** như: `xem tại phiếu Kết quả đo độ ồn...` (do không rõ vị trí đo và không rõ nơi lưu trữ tài liệu).



      * **Bắt buộc ghi rõ các vị trí đo và trạng thái in kèm theo hồ sơ (Cách 1 - Chuẩn QUATEST 3):**



        > `Phiếu kết quả đo độ ồn ở các vị trí <liệt kê tên các vị trí, ví dụ: VT1, VT2,... hoặc Tên vị trí cụ thể> (thuộc hồ sơ số <Mã_hồ_sơ> ngày dd/mm/yyyy) được in kèm theo hồ sơ.`



      * Tên và số lượng các vị trí đo phải đồng bộ 100% với Biên bản lấy mẫu/thực hiện (BBLM), sơ đồ và các hình ảnh chụp hiện trường.



      * Bản in phiếu kết quả đo đạc từ thiết bị bắt buộc phải được in ra và kẹp liền kèm theo tập hồ sơ giấy nộp lưu trình duyệt.



  - **Đánh giá (Assessment):**



    + Khớp chính xác với kết luận trong TBKQ: Phân định rõ từng phòng/vị trí về chỉ tiêu đạt và chỉ tiêu không đạt (ví dụ: phòng sạch đạt nồng độ bụi, HEPA, chênh áp nhưng không đạt số lần trao đổi khí ACH).



    + Ghi đầy đủ căn cứ đánh giá: Tiêu chuẩn ISO tương ứng và Mức quy định của khách hàng.



  - **Theo dõi KQTN & Báo cáo GĐV:**



    + Ghi nhận đầy đủ mã phiếu KQTN (MTTN..., HDTN...), ngày gửi mẫu, ngày yêu cầu cấp và ngày nhận KQTN.



    + Ngày kết thúc hồ sơ thực tế & Ngày báo cáo: Đồng bộ với ngày ban hành trên TBKQ (`dd/mm/yyyy`).



    + Mục trễ hạn, bất thường: Tích chọn `không` (hoặc nêu rõ nếu có).







## 8. Quy tắc Tự động Tách File Scan PDF Dữ liệu Đầu vào (E doc)



- **Mục đích & Phân loại tài liệu scan:**



  - Toàn bộ file scan tài liệu gốc nhận từ hiện trường (gồm BBLM, Ghi chép, phiếu đo đạc, máy đếm hạt, v.v.) tự động tách thành 2 file PDF chuẩn trình ký (Edoc) phân theo header trang tài liệu:



    1. `<Mã_6_số> - BBLM.pdf`: Dành riêng cho **Biên bản giám định hiện trường** (Biểu mẫu `BM3030.02`, Biên bản lấy mẫu/thực hiện có tiêu đề `BIÊN BẢN GIÁM ĐỊNH HIỆN TRƯỜNG`).



    2. `<Mã_6_số> - Ghi Chep.pdf`: Dành riêng cho **Ghi chép giám định tại hiện trường** (Biểu mẫu `BM3030.01` có tiêu đề `GHI CHÉP GIÁM ĐỊNH TẠI HIỆN TRƯỜNG`), các phiếu ghi chép thông số (thông số tủ, phòng...), và **toàn bộ các trang không có header rõ ràng** (phiếu in nhiệt đếm hạt bụi, kiểm tra rò rỉ màng lọc HEPA, phiếu đo độ ồn, bảng phụ lục kết quả...).



- **Vị trí lưu trữ bắt buộc:**



  - 2 file PDF sau khi xuất ra **bắt buộc lưu duy nhất vào thư mục con `TBKQ/Edoc/`** nằm ngay trong thư mục hồ sơ: `Input\<tên hồ sơ>\TBKQ\Edoc\` (ví dụ: `Input\132918\TBKQ\Edoc\132918 - BBLM.pdf`, `Input\132918\TBKQ\Edoc\132918 - Ghi Chep.pdf`).



  - **Tuyệt đối không lưu trùng tại thư mục gốc `Input\<tên hồ sơ>\`**.



  - Tự động sao chép sang thư mục `4. Edoc` trên ổ mạng (`Z:\...\<Mã_hồ_sơ>\4. Edoc\` hoặc `TBKQ\Edoc\`) khi kết nối ổ mạng.



- **Áp dụng tự động cho toàn bộ hồ sơ:**



  - Quy tắc này áp dụng tự động cho **mọi hồ sơ** thông qua `DossierRunner` (tự động thực thi ngay ở Step 0 trước khi sinh tài liệu) hoặc chạy độc lập qua công cụ `Core/Tools/scan_splitter.py`:



    + Tách scan cho 1 hồ sơ: `python Core/run.py --split-scan "Input\<tên hồ sơ>"`



    + Quét và tách toàn bộ các hồ sơ: `python Core/run.py --split-all-scans`



    + Chạy trực tiếp qua Tool: `python Core/Tools/scan_splitter.py "Input\<tên hồ sơ>"`







## 8. Quy tắc cho Tủ an toàn sinh học (BSC) & Kiểm tra rò rỉ màng lọc HEPA



- **Thể hiện độ rò rỉ lọc HEPA khi quét là 0,0000%:** Khi kết quả quét rò rỉ màng lọc HEPA (Downflow Filter hoặc Exhaust Filter) của BSC là `0,0000%` (hoặc `0`, `0.0`, `0%`):



  - **Bản tiếng Việt:** Bắt buộc thể hiện trên văn bản (TBKQ tiếng Việt, BCKQ, Bảng tính Excel) là: `< 0,0001%`.



  - **Bản tiếng Anh:** Bắt buộc thể hiện trên văn bản (TBKQ tiếng Anh) là: `< 0.0001%`.



  - Đánh giá vẫn là `Đạt` / `Pass` (do `< 0,0001%` nhỏ hơn rất nhiều so với mức quy định `≤ 0,01%`).



  - Nếu kết quả đo rò rỉ > 0 và ≥ 0,0001%: Thể hiện giá trị đo thực tế với 4 chữ số thập phân (ví dụ: `0,0006 %`, `0,0021 %`, `0,2578 %`) với dấu phẩy ở bản tiếng Việt và dấu chấm ở bản tiếng Anh.



  - Cột Mức quy định trong bảng TBKQ: Tiếng Việt là `≤ 0,01% ⁽¹⁾`, Tiếng Anh là `≤ 0.01% ⁽¹⁾`.







---







## 9. Quy tắc Đồng bộ Kích thước và Độ dài Bảng kết quả (Tiếng Anh & Tiếng Việt)

- **Bối cảnh & Nguyên tắc cốt lõi:**

  - Do đặc thù văn bản kỹ thuật, câu chữ tiếng Anh thường ngắn gọn hơn tiếng Việt (ít dòng hơn, từ ngữ cô đọng, không kèm giải thích song ngữ trong ngoặc đơn như tiếng Việt).

  - Nếu chỉ áp dụng thuộc tính tự co giãn hoặc chiều cao mặc định tối thiểu (400 twips), bảng kết quả tiếng Anh sẽ bị ngắn hơn đáng kể so với bản tiếng Việt (chênh lệch từ 5 – 10 cm). Điều này dẫn tới hiện tượng nghiêm trọng: các hàng bị dồn ép, vị trí ngắt trang bị lệch hoàn toàn (ví dụ: một chỉ tiêu bị chẻ đôi giữa 2 trang ở bản tiếng Anh trong khi bản tiếng Việt nằm trọn vẹn ở trang sau), hoặc số trang tổng thể của 2 bản bị lệch nhau (ví dụ: tiếng Việt 4 trang, tiếng Anh 3 trang).

  - **Quy tắc bắt buộc:** Khi một hồ sơ có cả 2 phiên bản tiếng Việt và tiếng Anh (hoặc khi xuất bản tiếng Anh tương ứng), độ dài và kích thước của bảng kết quả giữa 2 bản **bắt buộc phải tương tự nhau**, đảm bảo tính cân đối thị giác, đồng nhất cấu trúc trang và đồng bộ 100% vị trí ngắt trang giữa 2 ngôn ngữ.

- **Quy chuẩn Chiều cao dòng (Row Height) đồng bộ:**

  - Bắt buộc áp dụng chiều cao dòng tối thiểu tương ứng theo số dòng nội dung thực tế của hàng trong bản tiếng Việt:

    + **Hàng tiêu đề cột / Hàng 1 dòng ngắn:** Chiều cao tối thiểu **440 twips** (~0,78 cm).

    + **Hàng 2 dòng** (ví dụ: Cường độ sáng/Lux, Lọc khí cấp, Lọc khí thải, Chu vi cửa mở, các kích thước hạt bụi...): Chiều cao tối thiểu **560 – 580 twips** (~1,0 cm).

    + **Hàng 3 dòng** (ví dụ: Tốc độ trung bình dòng khí qua cửa làm việc, Tốc độ dòng khí cấp qua lọc HEPA...): Chiều cao tối thiểu **720 – 760 twips** (~1,3 cm).

    + **Hàng 4 dòng trở lên** (ví dụ: Vị trí đo kiểm tra hình thái dòng khí bên trong không gian tủ): Chiều cao tối thiểu **960 – 1000 twips** (~1,7 cm).

- **Đồng bộ nội dung chi tiết (Không lược bỏ gạch đầu dòng con trong tiếng Anh):**

  - Tuyệt đối không được tóm tắt hay lược bỏ các gạch đầu dòng con mô tả vị trí/kết quả trong bản tiếng Anh.

  - Ví dụ chỉ tiêu Kiểm tra hình thái dòng khí (Airflow pattern test):

    + Bản tiếng Anh bắt buộc phải có đầy đủ 3 gạch đầu dòng con bên trong tủ:

      `Inside cabinet workspace\n- Along center line, above work opening\n- Behind work sash\n- Along edge of work sash`

      (tương ứng 100% với bản tiếng Việt: `Bên trong không gian tủ\n- Dọc đường trung tâm, phía trên khoảng mở làm việc\n- Phía sau cửa mở làm việc\n- Dọc theo mép cửa mở làm việc`).

    + Dòng chu vi cửa mở ghi rõ: `Area along the outer perimeter of the work opening` (tương ứng với `Khu vực dọc theo chu vi phía bên ngoài của vùng mở cửa tủ`).

- **Chống chia cắt chỉ tiêu qua trang (`keepNext` / `cantSplit`):**

  - Các dòng con của cùng một chỉ tiêu lớn (ví dụ: Chỉ tiêu 6 của BSC, Chỉ tiêu 4 của Tủ hút) phải luôn đi liền với nhau, thiết lập `keep_with_next = True` trên cả dòng tiêu đề chỉ tiêu và dòng con đầu tiên để ngăn Word tự động ngắt dòng con rớt sang trang khác.

  - Bảng kết quả phải ngắt trang tại cùng một vị trí hàng giữa bản tiếng Anh và bản tiếng Việt (ví dụ: với BSC, cả 2 bản đều ngắt sạch sẽ sau Mục 5, toàn bộ Mục 6 và Mục 7 nằm trọn vẹn ở trang phụ lục tiếp theo).

- **Đồng bộ Tổng số trang tài liệu:**

  - Tổng số trang của bản tiếng Anh và tiếng Việt phải khớp nhau 100% (ví dụ: BSC không đo bụi cùng 3 trang, BSC có đo bụi cùng 4 trang, Tủ hút cùng 4 trang).

  - Số trang hiển thị trên Footer phụ lục (`Trang X/Y` và `Page X/Y`) phải hoàn toàn trùng khớp và chính xác.

- **Phạm vi áp dụng:**

  - Áp dụng trên 100% các loại hồ sơ (BSC, Tủ LAF, Tủ hút Fume Hood, Phòng sạch, Khí nén...).



---



## 10. Kiến trúc Module Tái Sử Dụng (QUATEST 3 Framework)



Toàn bộ project đã được module hóa và tối ưu triệt để thành các package chuyên biệt trong thư mục `Core/` để **tái sử dụng 100% cho mọi hồ sơ tiếp theo mà không cần phải viết code xử lý Word mới**:







```



Core/



├── Models/              # Data models: Dossier, ClientInfo, Equipment, PersonnelInfo, MonitoringLocation



├── Styling/             # Typography, Table AutoFit, Row Height (0.7-0.8cm), tblHeader, Ending Line



├── Builders/            # Bộ sinh tài liệu: BCKQBuilder, TBKQBuilder



│   └── handlers/        # Handlers chuyên biệt: BSCHandler, LAFBSCHandler, CleanroomHandler, CompressedAirHandler



├── Exporters/           # PDFExporter (Word COM, tự động dispatch file vào Input/<hồ sơ>/TBKQ/)



├── Services/            # DossierRunner (Pipeline tự động hóa từ dữ liệu -> DOCX -> PDF -> QA check)



├── Templates/           # Thư mục template chuẩn (BCKQ phong sach/khi nen/bsc, TBKQ VN/EN, tem tu)



└── Tools/               # Công cụ bổ trợ: bsc_calc, qa_check, calc_engine, iso8573_cal, docx_patcher...



```







### Cách sử dụng khi có hồ sơ mới:



Khi có hồ sơ mới, chỉ cần khai báo đối tượng `Dossier` (hoặc cung cấp file `profile.json` hay file Excel tính toán) và gọi `DossierRunner`:







```python



from Core import Dossier, ClientInfo, Equipment, PersonnelInfo, MonitoringLocation, DossierRunner







# 1. Khai báo thông tin hồ sơ



dossier = Dossier(



    code="Q326Dxxxxxx",



    cert_date="dd/mm/yyyy",



    sample_date="dd/mm/yyyy",



    dossier_type="fumehood",  # 'bsc', 'laf', 'fumehood', 'cleanroom', 'compressed_air'



    client=ClientInfo(name_vn="...", address_vn="..."),



    personnel=PersonnelInfo(executor="Phát"),  # Tự động tính GĐV ký và Người báo cáo



    equipments=[Equipment(name_vn="...")],



    locations=[MonitoringLocation(name_vn="...")],



    results_data={...}



)







# 2. Thực thi toàn trình sinh BCKQ + TBKQ VN/EN + Tem tủ + File Excel tính toán + Export PDF + Verify QA



runner = DossierRunner(dossier)



runner.run_all()



```







### Chạy nhanh qua CLI:



- Chạy profile Tủ an toàn sinh học (BSC) mẫu: `python Core/run.py --sample-bsc`



- Chạy profile BSC trực tiếp từ file Excel tính toán: `python Core/run.py --bsc-excel "path/to/TinhToanBSC.xlsx"`



- Chạy profile Tủ hút khí độc (Fume Hood) mẫu (Có đánh giá): `python Core/run.py --sample-fumehood`



- Chạy profile Tủ hút khí độc form Không đánh giá: `python Core/run.py --sample-fumehood --khong-danh-gia`



- Chạy profile Tủ hút khí độc xuất đồng thời cả 2 form (Có đánh giá & Không đánh giá): `python Core/run.py --sample-fumehood --both-forms`



- Chạy profile Tủ hút khí độc từ file Excel tính toán: `python Core/run.py --fumehood-excel "path/to/TinhToanFumehood.xlsx" [--khong-danh-gia | --both-forms]`



- Chạy profile LAF mẫu: `python Core/run.py --sample-laf`



- Chạy profile Khí nén mẫu: `python Core/run.py --sample-khinen`



- Chạy profile Khí nén thực tế (Q326D133266 - KN1 & KN2): `python Core/run.py --sample-khinen-133266`



- Chạy hồ sơ Khí nén trực tiếp từ file hoặc thư mục Excel tính toán: `python Core/run.py --khinen-excel "path/to/TinhToanKhiNen.xlsx"` hoặc `python Core/run.py --khinen-excel "Input\<tên hồ sơ>\TBKQ"`



- Chạy profile từ file JSON cấu hình: `python Core/run.py --profile profile.json`


## 11. Quy tắc Giữ nguyên Công thức Tính toán trong Excel & Kiểm tra Lệch Số liệu Hiện trường (Verification & Mismatch Warning)

- **Nguyên tắc Cốt lõi — Giữ nguyên 100% Công thức Động trong Excel:**
  - Toàn bộ các file bảng tính Excel tính toán kết quả nội bộ (`TinhToanBSC.xlsx`, `TinhToanFumehood.xlsx`, `TinhToanKhiNen.xlsx`...):
    + Các ô tính giá trị trung bình (`=AVERAGE(...)`), giá trị cực đại/cực tiểu (`=MAX(...)`, `=MIN(...)`), công thức nhân hệ số k (`=L22*L23`), công thức tính hiệu độ ồn (`=I69-I70`), công thức tính diện tích và vận tốc gió (`=G12/A12/3600`), các công thức hàm điều kiện đánh giá (`=IF(...)`):
    + **BẮT BUỘC giữ lại dạng công thức động Excel (Dynamic Formulas)**, **tuyệt đối không ghi đè bằng giá trị tĩnh (hardcoded static values)**.
    + Khi mở file hoặc chạy qua Excel COM (`excel.CalculateFull()`), Excel sẽ tự động đánh giá và tính toán độc lập một lần nữa trên chính các điểm đo hiện trường.

- **Giá trị trên file Scan Hiện trường chỉ dùng để Tham khảo / Đối chiếu:**
  - Các số liệu ghi chép trung bình, số nhẩm tay của giám định viên trên bản scan hiện trường (BM3030.01) đóng vai trò là số liệu tham chiếu (Reference Value).
  - Không được dùng số nhẩm tay trên scan để ghi đè công thức trong Excel. Nếu giám định viên tính nhẩm sai hoặc làm tròn lệch ở hiện trường, công thức Excel chuẩn xác vẫn phải được bảo toàn.

- **Cơ chế Cảnh báo Lệch Số liệu (Data Mismatch Warning):**
  - Hệ thống tự động chạy bộ tính toán độc lập từ các điểm đo chi tiết (raw readings) và đối chiếu với số liệu tham khảo trên scan.
  - Ngưỡng phát hiện sai lệch:
    + **Vận tốc gió (Inflow, Downflow, Face velocity...):** Nếu lệch $|V_{tính} - V_{scan}| > 0,01\text{ m/s}$ -> Bật cảnh báo `[WARNING] [Data Mismatch]`.
    + **Độ rọi (Lux):** Nếu lệch $|Lux_{tính} - Lux_{scan}| > 1\text{ lux}$ -> Bật cảnh báo `[WARNING] [Data Mismatch]`.
    + **Độ ồn (Noise):** Nếu lệch $|Noise_{tính} - Noise_{scan}| > 0,1\text{ dBA}$ -> Bật cảnh báo `[WARNING] [Data Mismatch]`.
    + **UVC:** Nếu lệch $|UVC_{tính} - UVC_{scan}| > 1\ \mu\text{W/cm}^2$ -> Bật cảnh báo `[WARNING] [Data Mismatch]`.
    + **Đánh giá (Pass/Fail):** Nếu kết quả tính toán độc lập kết luận khác với kết luận trên scan (ví dụ: tính toán ra Không đạt nhưng scan ghi Đạt) -> Bật cảnh báo cực kỳ nghiêm trọng `[CRITICAL WARNING] [Assessment Mismatch]`.
  - Toàn bộ cảnh báo phải được in nổi bật trong khung cảnh báo trên màn hình console và ghi vào kết quả kiểm tra để người phụ trách hồ sơ nắm rõ ngay lập tức.
