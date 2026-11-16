# Quy tắc định dạng Font chữ, Kiểu chữ và Trình bày Thiết bị sử dụng

Tài liệu bổ sung quy tắc bắt buộc khi lập Hồ sơ / Biểu mẫu / Báo cáo / Chứng thư giám định môi trường (Phòng N2 - QUATEST 3).

---

### PHẦN 2: QUY TẮC ĐỊNH DẠNG FONT CHỮ VÀ KIỂU CHỮ (FONT & STYLING)
1. **Font chữ & Cỡ chữ**:
   - Sử dụng thống nhất font chữ **Times New Roman**, cỡ chữ **13pt** (`Pt(13)`) cho toàn bộ các nội dung văn bản, thông tin khách hàng, số liệu, tiêu đề và bảng biểu.
2. **Quy cách in đậm / in thường**:
   - **Không in đậm** (`bold=False`) các thông tin điền mới (Tên đơn vị yêu cầu, Địa chỉ, Địa điểm thực hiện, Thời gian thực hiện, Thiết bị sử dụng, Quy trình thực hiện, Nội dung yêu cầu, Tên chỉ tiêu trong bảng, Số liệu đo...). Giữ đúng phong cách biểu mẫu hành chính gốc của QUATEST 3.
   - Chỉ in đậm (`bold=True`) tiêu đề chính của các bảng đo đạc hiện trường đặt phía trên bảng (và các giá trị kết quả **Không đạt** trong BCKQ/TBKQ theo quy định đánh giá).
   - **Quy cách in đậm và nghiêng cho câu đánh giá / kết luận (Assessment Phrasing)**:
     + Trong các câu đánh giá/kết luận đi kèm bảng kết quả đo đạc (sau dòng tiêu đề `Đánh giá:` / `Assessment:`):
       * Bản tiếng Việt: Bắt buộc **in đậm và in nghiêng** (`bold=True, italic=True`) đối với các cụm từ: **_phù hợp_**, **_không phù hợp_**.
       * Bản tiếng Anh: Bắt buộc **in đậm và in nghiêng** (`bold=True, italic=True`) đối với các cụm từ: **_comply_**, **_complied_**, **_did not comply_**.
     + Các phần còn lại của câu đánh giá giữ nguyên định dạng in thường (font Times New Roman 13pt).
3. **Ngày hiệu lực biểu mẫu (`HL:...`)**:
   - Cụm từ ngày hiệu lực ở góc trên bên phải bảng Header Table 0 (ví dụ: `HL: 29/09/2026`, `HL:01/8/2025`) là **ngày hiệu lực cố định của form biểu mẫu**, tuyệt đối **KHÔNG** được thay đổi theo ngày làm việc tại hiện trường.
4. **Ngày ban hành chứng thư (Date / Ngày)**:
   - Khi cập nhật ngày ban hành (theo ngày hiện tại hoặc ngày chỉ định dạng `dd/mm/yyyy`):
     + **Áp dụng đồng bộ cho cả bản tiếng Việt và tiếng Anh** (VN: `Số: ... - Ngày dd/mm/yyyy`, EN: `No: ... – Date: dd/mm/yyyy`).
     + Thay thế tại **mọi vị trí header** (header trang bìa và header phụ lục) cũng như dòng tiêu đề trên trang bìa.
     + Gỡ bỏ placeholder `………` và gỡ bỏ hoàn toàn định dạng tô nền vàng (`highlight`), giữ nguyên font chữ Times New Roman 12pt/13pt theo style của biểu mẫu.
5. **Quy tắc Số chứng thư trên Header vs Số hồ sơ trong thân văn bản**:
   - **Đồng bộ Số chứng thư trên các header:** Mọi header trong cùng một tài liệu (Section 0 và Section 1) bắt buộc phải có cùng một Số chứng thư (vd: `Số: Q326D122266/2 - Ngày 20/10/2026` hoặc `No: Q326D122266/2/E – Date: 20/10/2026`), tuyệt đối không được lệch nhau.
   - **Số hồ sơ trong thân văn bản (Body Text & Bảng Table 0):**
     + Chỉ ghi **mã trần (mã gốc)** của hồ sơ (ví dụ: `Q326D122266`).
     + **Tuyệt đối KHÔNG có đuôi `/1`, `/2`, `/3`, `/E`, v.v.** ở thân văn bản (dòng Số giấy tiếp nhận / Application No. của Table 0 hoặc các đoạn văn). Các đuôi phân nhánh `/1`, `/2`, `/E` chỉ được phép xuất hiện tại Số chứng thư trên Header và tiêu đề trang bìa.
     + Giữ nguyên vẹn nhãn của ô trước mã hồ sơ (ô `Giấy yêu cầu và nhận giám định số:` hoặc `Inspection application no. :`), không ghi đè mã hồ sơ làm mất nhãn ô.
6. **Quy tắc Đường kết thúc văn bản (Ending Line)**:
   - Toàn bộ các báo cáo/chứng thư kết quả (TBKQ) khi kết thúc văn bản **bắt buộc luôn có đường kết thúc đôi (`thinThick` compound line, 4.5pt strokeweight, màu đen)** nằm ở cuối cùng của văn bản, căn giữa trang (`align center` / `wp:align center` qua hai lề `margin`).
   - Đường kết thúc được neo vào đoạn kết luận cuối cùng (hoặc đoạn văn bản cuối cùng của tài liệu) với khoảng cách thẳng đứng phù hợp (~7-14pt dưới dòng chữ cuối cùng), không tạo đoạn trống thừa làm phát sinh thêm trang mới.
   - Tuyệt đối không để sót các đường nối mồ côi (`orphaned connector`) ở các đoạn văn giữa văn bản.
   - Công cụ QA Check (`qa_check.py`) tự động kiểm tra sự hiện diện của đường kết thúc chuẩn (`thinThick`, centered) ở cuối tài liệu và phát hiện đường nối mồ côi.

---

### PHẦN 3: QUY TẮC TRÌNH BÀY THIẾT BỊ SỬ DỤNG (EQUIPMENT)
1. **Danh mục thiết bị và quy trình theo lĩnh vực**:
   3.1. **Khí nén (N2.TC.KN.01)**:
     - Có chỉ tiêu đo hạt bụi trong khí nén:
       - Nếu yêu cầu đo hạt bụi kích thước đến 0,1 µm (`0,1um`, `0.1um`, `0,1-5,0um`...) -> `Máy đo nồng độ hạt bụi Lasair III 110 và phụ kiện HPD II, USA;`
       - Nếu đo nồng độ hạt thông thường (≥ 0,3 µm) -> `Máy đo nồng độ hạt AeroTrak 9350 và phụ kiện HPD III, USA;`
     - Có chỉ tiêu đo điểm sương -> `Thiết bị đo nhiệt độ điểm sương;`
     - Có chỉ tiêu lấy mẫu vi sinh khí nén -> `Thiết bị lấy mẫu vi sinh khí nén SAS Pinocchio Super II.`
     - Có chỉ tiêu lấy mẫu tổng dầu -> `Dụng cụ lấy mẫu tổng dầu (Regulator, Charcoal tube, Cassette).`
     - Quy trình áp dụng: `QUATEST 3 3030:2025; N2.TC.KN.01` (bổ sung `N2.TC.09` nếu có vi sinh).

   3.2. **Phòng sạch (N2.TC.PS.01 & N2.TC.09)**:
     - Đo nồng độ hạt bụi thông thường (≥ 0,3 µm, ISO 7-9, Grade C, D) -> `Thiết bị đo nồng độ hạt TSI AeroTrak 9350, USA;`
     - Đo hạt bụi siêu mịn (kích thước đến 0,1 µm, ISO 1-6, Grade A, B) -> `Máy đo nồng độ hạt bụi PMS Lasair III 110, USA;`
     - Đo nhiệt độ, độ ẩm -> `Thiết bị đo nhiệt độ, độ ẩm trong không khí Testo 625, Germany;`
     - Đo chênh lệch áp suất -> `Thiết bị đo chênh lệch áp suất TSI 8375, USA;`
     - Đo độ ồn -> `Thiết bị đo độ ồn Rion NL 53, Japan;`
     - Đo độ sáng -> `Thiết bị đo độ sáng Testo 440, Germany;`
     - Đo lưu lượng khí -> `Thiết bị đo lưu lượng TSI 8375, USA;`
     - Lấy mẫu vi sinh trong không khí -> `Thiết bị lấy mẫu vi sinh trong không khí Buck – Bio Culture B30120, USA.`
     - Quy trình áp dụng: `QUATEST 3 3030:2025; N2.TC.PS.01` (ghép `N2.TC.09` nếu có vi sinh, `N2.TC.01`, `N2.TC.02` nếu có vi khí hậu/ồn).

   3.3. **Tủ an toàn sinh học (BSC), Tủ sạch (LAF), Tủ hút (N2.TC.BSC.01/02, N2.TC.LAF.01, N2.TC.T.01)**:
     - Tốc độ gió vào (inflow) -> `Thiết bị đo lưu lượng TSI 8375/8380, USA;` hoặc `Máy đo tốc độ gió Kanomax Climomaster 6501, Japan;`
     - Tốc độ gió cấp (downflow) -> `Máy đo tốc độ gió Kanomax Climomaster 6501, Japan;`
     - Kiểm tra rò rỉ lọc HEPA -> `Máy tạo sol khí ATI aerosol generator, model 6D và quang kế aerosol photometer, model 2i, USA;`
     - Kiểm tra hình thái dòng khí -> `Máy tạo khói khí dung (Smoke pencil / Aerosol generator);`
     - Cường độ sáng -> `Máy đo độ rọi Testo 540 / Testo 440, Germany;`
     - Đo độ ồn -> `Máy đo độ ồn Rion NL-53, Japan;`
     - Đo độ rung -> `Máy đo độ rung RION VM-85, Japan;`
     - Đo cường độ UVC -> `Máy đo cường độ bức xạ UVC Extech SDL470, USA;`
     - Đo nồng độ hạt bụi trong tủ -> `Máy đếm hạt TSI AeroTrak 9350, USA.`
     - Quy trình áp dụng: `QUATEST 3 3030:2025; EN 12469:2000; N2.TC.BSC.02` (tủ BSC) hoặc `N2.TC.LAF.01` (tủ LAF) hoặc `N2.TC.T.01` (tủ hút).

   3.4. **Mapping nhiệt độ / Kho lạnh (N2.TC.MP.01)**:
     - Thiết bị ghi nhiệt ẩm tự động -> `Thiết bị tự ghi nhiệt độ độ ẩm Haxo-8 (mapping);`
     - Thiết bị kiểm tra đối chứng -> `Thiết bị đo nhiệt độ độ ẩm Testo 625, Germany.`
     - Quy trình áp dụng: `QUATEST 3 3030:2025; N2.TC.MP.01`.

   3.5. **Vi khí hậu & Môi trường lao động (N2.TC.01 đến N2.TC.08)**:
     - Vi khí hậu (nhiệt, ẩm, gió) -> `Máy đo nhiệt độ, độ ẩm Testo 625, Germany;`, `Máy đo tốc độ gió Kanomax Climomaster 6533-2G, Japan;`
     - Đo độ rọi -> `Máy đo độ sáng Testo 440 Kit Lux, Testo – Germany;`
     - Đo độ ồn -> `Thiết bị đo mức âm thanh Rion NL-53, Japan;`
     - Đo bụi MTLĐ -> `Bơm lấy mẫu bụi cá nhân SKC / Buck Genie và cassette màng lọc;`
     - Đo khí CO2 / TVOC -> `Thiết bị đo nồng độ CO2 trong không khí testo 440, Germany;`, `Thiết bị đo nồng độ TVOC Drager Xam 8000, Germany;`
     - Test nhanh hơi khí độc -> `Bơm hút khí Gastec / Kitagawa và ống phát hiện nhanh chuyên dụng.`

   *Tuyệt đối không đưa các thiết bị của chỉ tiêu không được yêu cầu.*
2. **Xuống dòng riêng biệt**:
   - Mỗi thiết bị đo đạc / lấy mẫu phải nằm trên một dòng riêng biệt để người đọc dễ kiểm soát danh mục thiết bị mang theo.
3. **Ký tự đầu dòng**:
   - Mỗi đầu dòng thiết bị bắt đầu bằng đúng 1 dấu gạch ngang (`- Tên thiết bị...`).
   - Xóa bỏ triệt để bullet / numbering tự động (`w:numPr`) của Word để tránh lỗi phát sinh trùng lặp 2 dấu gạch (`- - `).
4. **Căn lề và thụt đầu dòng**:
   - Thụt lề dòng đầu tiên (`first_line_indent`) và thụt lề trái (`left_indent`), lề phải (`right_indent`) của tất cả các dòng thiết bị phải set về đúng **0cm** (không thụt lề).

---

### PHẦN 4: QUY TẮC HÌNH ẢNH MAPPING VỊ TRÍ LẤY MẪU VÀ HÌNH ẢNH THỰC HIỆN (IMAGES & LAYOUT DIAGRAMS)
1. **Số lượng hình ảnh thay đổi linh hoạt theo hồ sơ (tăng hoặc giảm)**:
   - Số lượng **hình ảnh mapping / sơ đồ vị trí lấy mẫu (quan trắc)** (`Sơ đồ các vị trí quan trắc`, `Layout of monitoring points`, hình ảnh mapping vị trí lấy mẫu/đo đạc...) và **hình ảnh quá trình thực hiện / vị trí lấy mẫu** (`Hình ảnh quá trình thực hiện`, `Hình ảnh vị trí thực hiện`, `Hình ảnh vị trí lấy mẫu`, `Photos of the measurement process`, `Photo of performance location`...) trong BCKQ và TBKQ **không cố định** theo số khung có sẵn của biểu mẫu mà **thay đổi (tăng hoặc giảm) theo thực tế từng hồ sơ**.
2. **Trường hợp số lượng hình ảnh tăng (nhiều hơn biểu mẫu)**:
   - Chủ động thêm dòng / thêm ô / nhân bản khung ảnh và dòng chú thích tương ứng (giữ nguyên định dạng bảng không viền, kích thước khung hiển thị, khóa tỉ lệ `LockAspectRatio` và kiểu chữ chú thích của mẫu chuẩn) để chèn đầy đủ toàn bộ hình ảnh thực hiện và sơ đồ mapping vị trí lấy mẫu của hồ sơ.
3. **Trường hợp số lượng hình ảnh giảm (ít hơn biểu mẫu)**:
   - Chủ động xóa bớt các khung ảnh, ô hoặc dòng bảng dư thừa (hoặc thu gọn từ bảng khung 2 cột về 1 khung ảnh khi chỉ có 1 ảnh). Tuyệt đối không giữ lại hình ảnh của hồ sơ cũ và không để khung ảnh thừa trống làm lệch bố cục trang.
4. **Trường hợp khách hàng yêu cầu không đưa hình ảnh**:
   - Ở BCKQ: xóa khung ảnh trống và ghi chú rõ sau đầu mục ảnh theo quy định.
   - Ở TBKQ: xóa hẳn đầu mục hình ảnh cùng bảng khung ảnh liên quan.

---

### PHẦN 5: QUY TẮC CHUYỂN TRANG ĐỐI VỚI BẢNG VÀ CÂU KẾT LUẬN (PAGE BREAK & KEEP WITH NEXT)
1. **Nguyên tắc câu kết luận nằm cùng trang với bảng**:
   - Đối với mọi bảng kết quả có **câu kết luận / đánh giá kèm theo bảng** (bao gồm bảng `Ghi chú:` / `Note:` ngay dưới bảng, dòng tiêu đề `Đánh giá:` / `Assessment:` và câu kết luận `"Kết quả đo ... phù hợp với mức quy định..."` / `"The measurement result of ... complied with..."`), **câu kết luận bắt buộc phải nằm cùng trang với bảng** (nằm cùng trang với toàn bộ bảng nếu bảng ngắn, hoặc tối thiểu nằm cùng trang với các dòng cuối của bảng nếu bảng dài qua nhiều trang).
   - **Tuyệt đối không** để bảng kết quả nằm hết ở trang trước còn phần `Đánh giá:` + câu kết luận (hoặc `Ghi chú:` + `Đánh giá:` + câu kết luận) bị đẩy rớt lẻ loi sang đầu trang tiếp theo.
2. **Thiết lập kỹ thuật trong Word / OpenXML (`w:keepNext`, `w:keepLines`, `w:cantSplit`)**:
   - Gắn thuộc tính **Keep with next (`<w:keepNext/>`)** cho:
     + Đoạn tiêu đề mục ngay phía trên bảng kết quả (để tiêu đề không bị tách rời khỏi bảng ở cuối trang).
     + Các đoạn văn (`<w:p>`) nằm trong **1–2 dòng cuối cùng của bảng kết quả** (dòng vị trí cuối và dòng `Giới hạn tối đa cho phép` / `Mức quy định` / `Mức yêu cầu`).
     + Tất cả các đoạn văn trong bảng **`Ghi chú:` (`Note:`)** đi kèm ngay dưới bảng kết quả.
     + Đoạn tiêu đề **`Đánh giá:` (`Assessment:`)** để luôn gắn liền với câu kết luận phía dưới.
   - Gắn thuộc tính **Keep lines together (`<w:keepLines/>`)** cho đoạn văn câu kết luận để câu kết luận không bị ngắt dòng sang trang mới, và gắn **`<w:cantSplit/>`** cho các dòng của bảng (`<w:trPr>`) để từng dòng bảng không bị cắt đôi giữa hai trang.
   - Nếu bảng ngắn nhưng bị tràn vài dòng cuối kèm câu kết luận sang trang sau gây mất cân đối, ưu tiên ngắt trang (`Page Break`) trước tiêu đề mục của bảng hoặc thu gọn khoảng trống thừa để toàn bộ cụm **Tiêu đề + Bảng + Ghi chú + Câu kết luận** nằm gọn trên cùng một trang.
3. **Định dạng in đậm và nghiêng cho cụm từ kết luận**:
   - Trong câu kết luận/đánh giá đi kèm bảng, các cụm từ kết luận bắt buộc phải được định dạng **in đậm và in nghiêng** (`bold=True, italic=True`):
     + Tiếng Việt: **_phù hợp_**, **_không phù hợp_**
     + Tiếng Anh: **_comply_**, **_complied_**, **_did not comply_**
   - Các từ ngữ xung quanh trong câu giữ nguyên định dạng in thường.

---

### PHẦN 7: QUY TẮC CHUẨN HÓA ĐỊA CHỈ & VIẾT HOA - VIẾT THƯỜNG
1. **Mở rộng các từ viết tắt thành nguyên chữ**:
   - `KCN` -> `Khu công nghiệp` (ví dụ: `KCN Tân Đông Hiệp B` -> `Khu công nghiệp Tân Đông Hiệp B`, `KCN Vĩnh Lộc` -> `Khu công nghiệp Vĩnh Lộc`)
   - `KCX` -> `Khu chế xuất` (ví dụ: `KCX Tân Thuận` -> `Khu chế xuất Tân Thuận`, `KCX Linh Trung` -> `Khu chế xuất Linh Trung`)
   - `KDC` -> `Khu dân cư` (ví dụ: `KDC Thới An` -> `Khu dân cư Thới An`)
   - `CCN` -> `Cụm công nghiệp` (ví dụ: `CCN Tân Mỹ` -> `Cụm công nghiệp Tân Mỹ`)
   - `KĐT` -> `Khu đô thị` (ví dụ: `KĐT Sala` -> `Khu đô thị Sala`, `KĐT Phú Mỹ Hưng` -> `Khu đô thị Phú Mỹ Hưng`)
   - `KKT` -> `Khu kinh tế` (ví dụ: `KKT Dung Quất` -> `Khu kinh tế Dung Quất`)
   - `KCNC` -> `Khu công nghệ cao`
   - `Tp.HCM` / `TP.HCM` / `TP. HCM` / `Tp. HCM` -> `Thành phố Hồ Chí Minh`
   - `QL.` / `QL` -> `Quốc lộ` (ví dụ: `QL. 1A` -> `Quốc lộ 1A`, `QL 13` -> `Quốc lộ 13`)
   - `TL.` / `TL` -> `Tỉnh lộ` (ví dụ: `TL. 8` -> `Tỉnh lộ 8`)
   - `Đ.` / `Đg.` -> `đường`
   - `P.` -> `phường` (nếu tên chữ: `P. Thới An` -> `phường Thới An`) / `Phường` (nếu tên số: `P. 1` -> `Phường 1`)
   - `Q.` -> `quận` (nếu tên chữ: `Q. Gò Vấp` -> `quận Gò Vấp`) / `Quận` (nếu tên số: `Q. 1` -> `Quận 1`)
   - `H.` -> `huyện` (ví dụ: `H. Hóc Môn` -> `huyện Hóc Môn`)
   - `TX.` -> `thị xã` (ví dụ: `TX. Bến Cát` -> `thị xã Bến Cát`)
   - `TT.` -> `thị trấn` (ví dụ: `TT. Hóc Môn` -> `thị trấn Hóc Môn`)
   - `X.` -> `xã` (ví dụ: `X. Tân Thới Hiệp` -> `xã Tân Thới Hiệp`)
   - `KP.` -> `Khu phố` (ví dụ: `KP. 1` -> `Khu phố 1`)
   - `TDP.` -> `Tổ dân phố` (ví dụ: `TDP. 5` -> `Tổ dân phố 5`)
2. **Ví dụ chuẩn hóa**:
   - Gốc: `"A87, KDC Thới An, Đường Lê Thị Riêng, Khu Phố 1, Phường Thới An, TP.HCM, Việt Nam"`
   - Sau chuẩn hóa: `"A87, Khu dân cư Thới An, đường Lê Thị Riêng, Khu phố 1, phường Thới An, Thành phố Hồ Chí Minh"`
3. **Loại bỏ chữ "Việt Nam" ở cuối**:
   - Tuyệt đối không ghi `Việt Nam`, `Vietnam`, `VN` ở phần cuối chuỗi địa chỉ.
4. **Thành phố trực thuộc Trung ương**:
   - Bắt buộc viết hoa chữ `Thành phố`:
     + `Thành phố Hồ Chí Minh` (thay thế cho `TP.HCM`, `Tp.HCM`, `TP. HCM`, `thành phố Hồ Chí Minh`...)
     + `Thành phố Hà Nội` (thay thế cho `TP. Hà Nội`, `Tp. Hà Nội`, `Hà Nội`...)
     + `Thành phố Đà Nẵng` (thay thế cho `TP. Đà Nẵng`, `Đà Nẵng`...)
     + `Thành phố Hải Phòng` (thay thế cho `TP. Hải Phòng`, `Hải Phòng`...)
     + `Thành phố Cần Thơ` (thay thế cho `TP. Cần Thơ`, `Cần Thơ`...)
     + `Thành phố Huế` (nếu có)
5. **Tỉnh và Thành phố thuộc tỉnh**:
   - Viết thường chữ `tỉnh` và `thành phố`:
     + Tỉnh: `tỉnh Bình Dương`, `tỉnh Đồng Nai`, `tỉnh Long An`, `tỉnh Bà Rịa - Vũng Tàu`...
     + Thành phố thuộc tỉnh: `thành phố Thủ Dầu Một`, `thành phố Dĩ An`, `thành phố Thuận An`, `thành phố Biên Hòa`, `thành phố Thủ Đức`...
6. **Quận / Huyện / Thị xã**:
   - Tên riêng là chữ -> viết thường: `quận Bình Thạnh`, `quận Tân Bình`, `quận Gò Vấp`, `huyện Hóc Môn`, `huyện Củ Chi`, `thị xã Bến Cát`...
   - Tên riêng là số -> viết hoa từ chỉ đơn vị: `Quận 1`, `Quận 12`...
7. **Phường / Xã / Thị trấn**:
   - Tên riêng là chữ -> viết thường: `phường Thới An`, `phường An Khánh`, `phường Bến Nghé`, `xã Tân Thới Hiệp`, `thị trấn Hóc Môn`...
   - Tên riêng là số -> viết hoa từ chỉ đơn vị: `Phường 1`, `Phường 14`...
8. **Khu phố / Thôn / Ấp / Tổ dân phố**:
   - `Khu phố`: viết hoa chữ `Khu`, viết thường chữ `phố` (`Khu phố 1`, `Khu phố 2`, `Khu phố Đông Chiêu`).
   - `Ấp`: viết hoa chữ `Ấp` (`Ấp 1`, `Ấp Bình Mỹ`...).
   - `Thôn`: viết hoa chữ `Thôn` (`Thôn 1`, `Thôn Phú Mỹ`...).
   - `Tổ dân phố`: viết hoa chữ `Tổ`, viết thường `dân phố` (`Tổ dân phố 1`, `Tổ dân phố 5`...).
9. **Đường / Phố / Đại lộ / Hẻm / Ngõ / Ngách**:
   - Viết thường danh từ chỉ loại đường giao thông: `đường Lê Thị Riêng`, `đường số 3`, `phố Tràng Tiền`, `đại lộ Mai Chí Thọ`, `hẻm 45`, `ngõ 12`...
10. **Khu công nghiệp / Khu chế xuất / Khu dân cư / Cụm công nghiệp**:
   - Mở rộng viết đầy đủ và viết hoa chữ cái đầu: `Khu công nghiệp Tân Bình`, `Khu chế xuất Tân Thuận`, `Khu dân cư Thới An`, `Cụm công nghiệp Tân Mỹ`... (viết hoa chữ `Khu`, `Cụm`, viết thường các từ sau `công nghiệp`, `chế xuất`, `dân cư`).
11. **Tra cứu Google / Web Search để xác minh địa danh & cập nhật đơn vị hành chính mới**:
   - **Được phép và khuyến khích tra cứu Google:** Khi thông tin địa chỉ hoặc địa điểm thực hiện trong yêu cầu / hợp đồng / báo giá viết tắt, ghép liền cụm danh từ hoặc thiếu từ chỉ cấp hành chính (ví dụ: `"KCN Vĩnh Lộc Hóc Môn, Tp.HCM"`), người soạn thảo/AI được phép và chủ động tra cứu Google / Web Search để xác định chính xác các cấp đơn vị hành chính và phân tách rõ ràng bằng dấu phẩy.
   - **Ví dụ chuẩn hóa:** `"KCN Vĩnh Lộc Hóc Môn, Tp.HCM"` -> `"Khu công nghiệp Vĩnh Lộc, xã Hóc Môn, Thành phố Hồ Chí Minh"`.
   - **Lưu ý đơn vị hành chính mới cho tất cả các địa chỉ:** Luôn cập nhật theo đơn vị hành chính mới nhất (các phường/xã mới sau khi sắp xếp, sáp nhập tại TP.HCM, Hà Nội; các thành phố trực thuộc tỉnh mới thành lập như `thành phố Bến Cát`, `thành phố Tân Uyên`, `thành phố Thủ Đức`...). Tuyệt đối không giữ tên đơn vị hành chính cũ đã bị bãi bỏ hoặc sáp nhập.
