# Công cụ & workflow

Làm bằng cách nào: script, bẫy, công thức. Quy trình và quy ước: `QUY-TRINH-XU-LY-HO-SO.md`.
Cấu hình máy và đường dẫn: `CLAUDE.md`.

---

## 1. Các bẫy — phải tránh

### Script / PowerShell

**Bẫy 1 — ký tự ngoài ASCII trong `.ps1`.** PS 5.1 đọc `.ps1` không BOM theo ANSI → đường dẫn
tiếng Việt hỏng. Kể cả `…` cũng không được: đã gặp script tự sinh JSON chứa `………`, đọc theo ANSI
rồi ghi lại UTF-8 → ra `â€¦â€¦â€¦` trong docx. → **Script thuần ASCII**, lấy thư mục bằng wildcard
6 số, mọi chuỗi tiếng Việt / ký tự đặc biệt nằm sẵn trong **file JSON UTF-8**, đọc bằng
`[System.IO.File]::ReadAllText($p,[Text.Encoding]::UTF8)`.

**Bẫy 2 — `$` bị nuốt khi dùng `powershell -Command "…"`.** → Luôn ghi ra `.ps1`, chạy `-File`.

**Bẫy 7 — trùng tên biến trong PowerShell (không phân biệt hoa/thường).** `$rows = …` bên trong
script sẽ ghi đè tham số `-Rows`. Đặt tên biến cục bộ khác hẳn tên tham số.

**Bẫy 9 — đường dẫn Windows trong JSON sinh từ Bash.** `echo`/`printf`/`sed` trong Git Bash ăn mất
một lớp `\` → JSON ra `E:\HoSo\...` → `ConvertFrom-Json: Unrecognized escape sequence`; `sed` còn
làm mất `\` trong wildcard (`0.VIMCERTS078\*` thành `0.VIMCERTS078*`). → Trong JSON dùng **gạch
chéo xuôi** (`E:/HoSo/_claude/...`); sửa `.ps1` bằng Edit/Write, không bằng `sed`.

**Bẫy 19 — `"...\\$var"` trong Git Bash.** Chuỗi nháy kép `"E:\x\\$n"` bị hiểu thành `\$` → ra chữ `$n`
thật (thư mục tên `094661$n`, `...21.9.2026$f`). → Truyền đường dẫn cho `.ps1` bằng **gạch chéo xuôi**
(`"E:/HoSo/_claude/_claude_tmp/$n"`); PowerShell nhận được cả đường dẫn tiếng Việt truyền theo tham số.

**Bẫy 21 — `pdf-render.ps1` không nhận gạch chéo xuôi.** WinRT `StorageFile.GetFileFromPathAsync` báo
`AggregateException` với `E:/HoSo/...` (ngược với Bẫy 19). → Truyền `-Pdf`/`-OutDir` bằng **gạch chéo
ngược** (từ Git Bash: `"E:\\HoSo\\_claude\\..."`). Gặp ở Q326D111235.

**Bẫy 22 — file đặt tên `3 - TBKQ …` (gạch ngang thay dấu chấm).** `dump-all.ps1`/`rename-done.ps1` trước
đây chỉ lọc `^[123]\.` → bỏ sót hết TBKQ, map.txt chỉ có BCKQ. Đã sửa lọc thành `^[123](\.| - )`
(24/09/2026). Luôn đếm số file trong map.txt khớp số file `2.*`/`3.*` thực tế.

**Bẫy 20 — `set-cells.ps1` / `overwrite-images.ps1` có thể để backup `*_yyyymmdd_hhmmss.docx` ngay trong
thư mục hồ sơ.** Sau khi chạy: `find . -maxdepth 1 -name "*_2026*" -exec mv {} E:/HoSo/_claude/backup/ \;`.

**Bẫy 4 — file đang mở trong Word/Excel bị khoá.** Script báo `LOCKED` rồi chạy tiếp; dấu hiệu
file đang mở: có `~$<tên>` cùng thư mục. Nhắc người dùng đóng file rồi chạy lại.

**Bẫy 17 — ghi số vào ô Excel qua COM.** `$ws.Range('A1').Value2 = 6016` báo *Unable to cast
Double to String* (PowerShell truyền giá trị như tham số thứ 2 của `Range`). → Lấy range ra biến
trước rồi gán `.Formula` bằng chuỗi số InvariantCulture:
`$rg=$ws.Range('A1'); $rg.Formula=[string]::Format([Globalization.CultureInfo]::InvariantCulture,'{0}',$v)`.
Ô gộp (merged) phải thao tác qua `.MergeArea`; ảnh "Place in cell" (ô hiện chữ `Picture`) chép bằng
`Range.Copy(dest)` giữa hai file cùng bố cục.

**Bẫy 44 — tên hàm trùng alias, và `\\$1` trong hàm Git Bash.** Hàm PowerShell đặt tên `SC` bị alias `Set-Content`
chiếm: lệnh không chạy mà ghi file rác ở thư mục hiện hành → không đặt tên hàm trùng alias (`sc`, `gc`, `rd`, …).
Biến thể Bẫy 19 với tham số vị trí: `"$T\\gc\\$1"` trong hàm Git Bash ra chữ `$1` thật → dùng gạch chéo xuôi.
Gặp ở Q326D124796, Q326D124835.

### Docx — chỉ số run, header, ảnh

**Bẫy 3 — sửa docx phải sửa cả header/footer.** `word/document.xml` KHÔNG chứa header. Mã hồ sơ,
số chứng thư và **ngày ban hành** nằm ở `header1`…`header4` tuỳ mẫu — thường **cả** header trang
đầu (`header1`/`header2`) lẫn header phụ lục; có mẫu còn giấu số hiệu văn bản khác hoặc số hồ sơ
cũ hơn ở header không dùng tới. `dump-all.ps1` dump `document` + `header1..4` + `footer1..2`;
`final-check.ps1` quét **mọi** header/footer — luôn chạy trước khi báo CLEAN.

**Bẫy 5 — ghép số hồ sơ dính ký tự cứng.** Có mẫu viết sẵn nửa mã trong run chữ, vd run kết thúc
bằng `"...số Q326D0"` rồi run sau mới là 5 chữ số. Mã mới **không** bắt đầu bằng `0` mà set thẳng
6 chữ số vào run sau → **7 chữ số sai** (`Q326D0104559`). → Sau khi vá, kiểm tra bằng regex
`Q326D\d{6}(?!\d)` trên toàn bộ text đã ghép.

**Bẫy 8 — thứ tự vá.** `set-cells.ps1` có thể **chèn run mới** vào ô rỗng → lệch mọi chỉ số run phía
sau. Trong cùng một file: `apply-all.ps1` (theo chỉ số run) **trước**, `set-cells` **sau**.

**Bẫy 11 — `apply-all.ps1` chỉ đếm `<w:t>`, không đếm `<w:instrText>`.** Footer `Trang {PAGE}/3`:
tổng số trang ghi cứng nằm ở run `<w:t>` thứ 4 (index 3). Luôn đọc dòng `set=x/y` trong log;
`x<y` là vá trượt. Tổng trang ghi cứng phải khớp số trang thực tế đếm bằng Word COM
(`ComputeStatistics(2)`) **sau khi vá** — đã gặp mẫu ghi `/3` mà thực tế 4 trang (Q326D106222,
Q326D106227).

**Bẫy 6 — "Word found unreadable content" khi chèn ảnh mới.** Đuôi entry `word/media/<ten>.<duoi>`
phải khớp **ký tự-cho-ký tự** với một `<Default Extension="<duoi>" …/>` trong `[Content_Types].xml`
(entry `.jpg` mà chỉ khai báo `jpeg` → gói hỏng, dù `ZipFile`/`XmlDocument` vẫn đọc OK).
→ Sau khi chèn/đổi ảnh luôn chạy `test-open.ps1`.

**Bẫy 10 — chữ Việt trong mẫu có thể là dấu tổ hợp (NFD).** `Việt Nam` lưu dạng `e` + dấu rời:
tìm chuỗi chính xác (find/replace, `-Old` của final-check) **không khớp**. → Tìm theo phần không dấu
hoặc so sánh sau `normalize('NFC')`; rà sót dấu vết cũ bằng từ khoá ASCII.

**Bẫy 23 — `unhl` của `apply-all.ps1` không gỡ hết tô vàng.** Sót highlight ở dấu đoạn (`<w:pPr><w:rPr>`),
ở run chữ cố định tô sẵn từ mẫu (vd "US EPA 200.8", "Đánh giá:") và cả vài run vừa `set` (Q326D113750:
còn 7 `w:highlight` sau khi unhl đủ 28 run). → Sau khi điền xong **toàn bộ**, đếm `w:highlight` trong
mọi entry; không còn ô cố ý chờ thì gỡ sạch bằng Word COM `$doc.Content.HighlightColorIndex=0` (backup trước).

**Bẫy 24 — header BCKQ ghi cứng cả số trang hiện tại (`Trang 1/1`, không có field PAGE).** Sau khi
`dup-row` thêm mẫu, BCKQ tràn sang trang 2 → sửa mỗi tổng trang thành `1/2` thì trang 2 vẫn hiện `Trang 1/2`.
→ Kiểm tra `instrText` trong header; không có PAGE thì thay chữ số trang hiện tại bằng field PAGE qua Word COM
(`$doc.Fields.Add($range,33)`, 33 = wdFieldPage), rồi đặt tổng trang = `ComputeStatistics(2)`. Word lưu lại có thể
đổi tên `header1`→`header2`. Gặp ở Q326D112883.

**Bẫy 25 — `map.txt` in text XML thô (`&amp;`, `&gt;`), còn `apply-all.ps1` tự escape khi `set`.** Lấy chuỗi
từ map.txt sửa rồi đưa lại vào `plan.json` (vd `US EPA 3450C &amp;` → `3510C`) sẽ ra `&amp;amp;` — Word hiện
chữ `&amp;`. → Trong `set` luôn ghi ký tự thật (`&`, `<`, `>`); grep `amp` trong plan.json trước khi chạy.
Gặp ở Q326D107520.

**Bẫy 26 — Word COM: `$doc.Range(s,e)` luôn là thân văn bản, kể cả khi `s` lấy từ range header.** Khi sửa Bẫy 24,
lấy `Start` từ `Headers.Item(1).Range` rồi gọi `$doc.Range(...)` → ghi đè ký tự **trong thân** BCKQ, `Fields.Add`
báo "This command is not available". → Tạo range con từ chính story header: `$t=$h.Range.Duplicate; $t.SetRange(s,e)`,
`$h.Range.Fields.Add($t,33)`; backup trước và so `$doc.Content.Text` trước/sau phải giống hệt. Gặp ở Q326D106685.

**Bẫy 27 — footer phụ lục TBKQ khí nén `Trang {PAGE}/2` đánh số riêng phần phụ lục.** Word đếm cả tờ đầu
(`ComputeStatistics(2)`=3) nhưng trang cuối hiện `Trang 2/3` nếu sửa tổng thành 3 → **sai**. Tổng trang ở footer
phụ lục = số trang phụ lục (tổng trang − 1), không phải tổng cả file. Xuất PDF xem trang cuối trước khi sửa Bẫy 11.
Gặp ở Q326D106978. TBKQ nước thải (Việt Sơn, Q326D106227) cũng vậy: tờ đầu không đánh số, footer phụ lục
`Trang {PAGE}/n` với n = số trang phụ lục — lần làm khung ghi `/4` theo tổng file là sai.

**Bẫy 31 — `pdf-render.ps1` báo `AggregateException` với PDF nằm trong thư mục hồ sơ trên `Z:` có tên tiếng Việt**
(kể cả truyền gạch chéo ngược). → Copy PDF sang `_claude_tmp\<6 số>\` (tên ASCII) rồi render. Gặp ở Q326D106227.

**Bẫy 38 — `pdf-render.ps1` gọi từ Git Bash có lúc lỗi `AggregateException` hoặc im lặng không ra PNG** (PDF vừa xuất bằng
Word `SaveAs2(...,17)` trong `_claude_tmp`, kể cả gạch chéo ngược) dù cùng cú pháp trước đó chạy được. → Gọi lại qua công cụ
PowerShell (`& powershell -File ...pdf-render.ps1 -Pdf "E:\..."`) thì chạy ngay; luôn kiểm tra thư mục ra có `pNN.png`. Gặp ở Q326D115373.

**Bẫy 28 — Git Bash `cd "<dir>"/Multi*`: thư mục KQTN tải về (`Multi_Documents_…`) luôn đi kèm file `.zip` cùng tên
→ glob ra 2 đối số, `cd` lỗi, lệnh `cd ..` sau đó nhảy lên thư mục tổng 2026.** → Gõ đủ tên thư mục, hoặc `Multi*/`.
Gặp ở Q326D106978.

**Bẫy 29 — Excel COM: file đang mở ở Excel của người dùng → `Workbooks.Open` tự mở read-only, ghi ô và
`Save()` KHÔNG báo lỗi nhưng không lưu gì.** Trước khi ghi: kiểm tra `~$<tên>.xlsx` và `$wb.ReadOnly`; true → dừng,
nhờ người dùng đóng file. Máy in mặc định là máy in nhãn `HPRT HT330` → `PageSetup.LeftHeader=` có thể không ăn
và `ExportAsFixedFormat` ra PDF trang trắng; kiểm tra lại header sau khi lưu, xem PDF trước khi đè bản cũ.
Gặp ở Q326D106978.

**Bẫy 30 — sửa header BCKQ (Bẫy 24) bằng offset chuỗi hoặc Find "/2" đều trượt.** `$h.Range.Text` không khớp 1-1
với vị trí ký tự (IndexOf ra range sai → hỏng header); `Find "/2"` trong header khớp nhầm ngày HL `01/8/2025` → `3025`.
→ Tìm `Trang [!/]@/` (MatchWildcards) để đặt field PAGE; tổng trang thì dò từng ký tự sau `$fld.Result.End`
(+0..+4) tới ký tự `2` rồi thay. Backup trước, in lại text header để kiểm. Gặp ở Q326D110364.

**Bẫy 32 — Word COM `Save()` gộp lại các run.** Sau khi mở/lưu docx bằng Word COM (chèn ảnh, sửa header),
chỉ số run trong `map.txt` cũ không còn đúng: `apply-all` vẫn báo `set=x/x` nhưng ghi đè nhầm run khác
(Q326D096779: định sửa phút "09:10" lại ghi đè chữ "Máy đo"). → Sau mỗi lần lưu qua Word COM phải `dump-all` lại trước khi vá theo chỉ số run.

**Bẫy 33 — hồ sơ nhiều đợt nằm ở thư mục con (`<mã>\Thực hiện\Đợt n\`).** `dump-all`/`apply-all`/`final-check`/`test-open` chỉ lấy thư mục gốc của mã → không thấy file `2.*`/`3.*`. Dùng `set-cells -Docx` (nhận đường dẫn) hoặc script riêng trong `_claude_tmp\<6 số>\`. Khi dựng lại thân docx từ các phần tử XML, **đoạn chứa `<w:sectPr>`** (ngắt section, vd trước trang sơ đồ ngang) phải giữ lại — mất nó Word gộp section và bỏ header/footer phụ lục khi lưu. Gặp ở Q326D120423.

**Bẫy 34 — nhân dòng bảng ảnh bằng chép XML làm trùng mã ảnh.** Chép nguyên `<w:tr>` chứa ảnh (để đủ khung ảnh
cho nhiều mẫu) thì `<wp:docPr id="…">` (và `<pic:cNvPr id>`) bị lặp → Word COM chế độ tự động không mở được
(`test-open` lỗi/treo, sinh tiến trình WINWORD treo + file khoá `~$`). → Sau khi nhân dòng, **đánh lại `id` duy
nhất** cho mọi `wp:docPr`/`pic:cNvPr` trong `document.xml` rồi mới chèn/thay ảnh:
`E:\HoSo\_claude\maps\fix-docpr-ids.ps1 -Docx <file> [-Start 900100] [-BackupDir E:\HoSo\_claude\backup]`
(sửa XML trực tiếp, tự backup, không đụng header; đóng file trong Word trước khi chạy). Hoặc ưu tiên cách 4c (Word COM tự cấp mã). Nếu
phải tắt Word treo, chỉ tắt tiến trình do chính phiên mình mở — hồ sơ khác chạy song song có thể đang dùng Word.
Gặp ở Q326D117748.

**Bẫy 35 — Word COM `AddPicture` + `Save()` tự nén/resample ảnh** (md5 media trong docx khác file gốc `0.*`). → Sau khi chèn
khung ảnh bằng Word COM (tỉ lệ đã đúng theo ảnh gốc), ghi đè lại byte gốc bằng `overwrite-images.ps1` (copy ảnh gốc sang
`_claude_tmp\<6 số>\` tên ASCII trước), rồi so md5 entry `word/media/*` với ảnh gốc (vd 829 KB → 195 KB nếu bỏ bước này). Gặp ở Q326D110632, Q326D116729.

**Bẫy 36 — `dump-all.ps1` chỉ dump `footer1..2`; TBKQ EN phòng sạch để `Page {PAGE}/n` ở `footer4.xml`** (header phụ lục ở
`header4`, chỉ số run `Q326D`/số hồ sơ lệch 1 so với `header2`). → Liệt kê mọi `word/header*.xml`/`footer*.xml` trong zip,
đọc `<w:t>` từng entry trước khi vá; sau khi vá xuất PDF xem dòng số hồ sơ và "Page x/n" ở trang cuối. Gặp ở Q326D116326
(đã ra `116326101664/2/E` và `Page 176` trước khi sửa).

**Bẫy 37 — dựng lại dòng bảng từ dòng mẫu: ô nhiều đoạn có `<w:tab/>`/`<w:br/>` trong run.** Ghi chữ mới vào `<w:t>` mà
giữ tab của mẫu → tab đẩy xuống dòng, sinh dòng trống giữa các đoạn (TBKQ EN phòng sạch cột "Parameter and number…").
Khi ghi đè ô phải xoá `w:tab`/`w:br`/`w:cr` trong run của đoạn đó. Script mẫu dựng lại nhiều dòng (giữ vMerge, ô nhiều
đoạn, tô vàng): `E:\HoSo\_claude\maps\tblfill.ps1 -Docx … -Spec spec.json`. Lưu ý PowerShell: mảng một phần tử `@(0)` là **false**
trong `if(...)` → dùng `if($null -ne $x)`; không đặt tên hàm `rd` (alias của Remove-Item).

**Bẫy 39 — Word COM `Documents.Open` lỗi COMException với đường dẫn `Z:/...` (gạch chéo xuôi, thư mục tên tiếng Việt) truyền từ Git Bash.** → Copy docx sang `_claude_tmp<6 số>` (tên ASCII), truyền đường dẫn gạch chéo ngược, xử lý xong copy đè lại. Gặp ở Q326D122166.

**Bẫy 40 — Word COM `SaveAs2(...,17)` (xuất PDF) trong vòng lặp có thể treo vô hạn** (không lỗi, WINWORD ẩn giữ file
khoá `~$`). Gặp ở Q326D109860 khi chèn ảnh 11 TBKQ: `Save()` xong, treo ở xuất PDF. → Tách bước xuất PDF ra script riêng,
chạy có `timeout`; treo thì chỉ `Stop-Process` đúng PID WINWORD do phiên mình mở (so `StartTime`), xoá `~$` rồi chạy lại.

**Bẫy 41 — mẫu khí nén mới `1. TinhToanKhiNen_KNx.xlsx` (mỗi vị trí một file) KHÁC bố cục mục 7a.** N* ở `C31` (nhiệt độ)/`C32`
(Pdd=1), ô `C31` ghi cứng số cũ → đặt `=DiemSuong!E4` (scan bụi dùng đúng nhiệt độ TB điểm sương); dầu `C38:D41`, KQ `C44`; bụi:
tổng hợp `D4:D7` trỏ vùng **Lasair** (dòng 49/53/54), không trỏ vùng Aerotrak (dòng 58–63) → số AeroTrak nhập theo quy ước 7a vào
`C49:J49` (`C:F`=0, `G`=0,3, `H`=0,5, `I`=1+2+3, `J`=5) và chép thô vào `C58:H58`; vi sinh `C66:G72`. `DiemSuong`: `C3:D5`, `G3`,
giờ `C7`/`C18`, ảnh `B6:B15`/`B17:B26` (copy theo **cả merge area** từ `B8:B17`/`B19:B28` file KN.xlsx, Copy ô đơn báo "can't do that
to a merged cell"), ảnh CS là shape đầu tiên của file KN (tên không cố định — dùng `Shapes.Item(1)`). Chưa có KQTN mà để ô dầu
trống/chữ → công thức ra **`< 0,01` – Class 1 giả**: phải chặn bằng nhánh `"………"` ở `C40/C42/C43/C44/F9` và vi sinh dòng 71–72.
Vị trí không đo bụi → `-` ở vùng bụi + `D4:D7`,`K4:K7`,`F3`. Script mẫu: `_claude_tmp\115269\xfill.ps1` + `mk.js`. Gặp ở Q326D115269.

**Bẫy 42 — TBKQ bùn/chất thải (Vimcerts): section phụ lục nằm ngang đặt `RestartNumberingAtSection=True`** → phụ lục 4 trang
hiện `Trang 1/3, 1/3, 2/3, 3/3` (trang dọc đầu phụ lục và trang ngang đầu tiên cùng là 1). → Word COM:
`$d.Sections.Item(3).Footers.Item(1).PageNumbers.RestartNumberingAtSection=$false`, tổng = `ComputeStatistics(2)-1` (Bẫy 27)
thay bằng `Find.Execute('/3',...,'/'+$n,1)` trên footer section 2; xuất PDF xem từng trang. Script: `_claude_tmp\112121\pg.ps1`.
Gặp ở Q326D112121.

**Bẫy 45 — xuất PDF bằng Word COM treo ngay cả với một file nhỏ** (`SaveAs2(…,17)` lẫn `ExportAsFixedFormat`, không chỉ trong
vòng lặp như Bẫy 40). → Ghi PID WINWORD ra file trong `_claude_tmp\<6 số>\` ngay sau `New-Object` để tắt đúng tiến trình của
mình; không suy PID từ giờ file copy vì `Copy-Item` giữ nguyên mtime. Xuất PDF không được thì phải báo rõ là chưa kiểm tra
bố cục/ảnh bằng mắt. Gặp ở Q326D124796.

**Bẫy 46 — nhiều khung ảnh cùng trỏ một media (`rId`), hoặc khung trỏ entry `.png`.** BCKQ sản phẩm có 10 khung chung một
`rId` → `overwrite-images.ps1` ra 10 ảnh giống nhau; entry `.png` không ghi đè được bằng byte JPEG. → Thay bằng Word COM rồi
ghi lại byte gốc (Bẫy 35). Gặp ở Q326D124796.

**Bẫy 47 — vá lệch 1 chỉ số run khi nhiều run ngắn liền nhau** (`ngày`, `mẫu Blank:`, `: `). `apply-all` vẫn báo `set=x/x`;
chỉ thấy khi xuất PDF xem. → Với cụm run ngắn, đối chiếu lại text run trước/sau trong map.txt trước khi `set`, và xem PDF
sau khi vá. Gặp ở Q326D124835 (BCKQ-1).

**Bẫy 50 — hồ sơ nhiều đợt (Q326D120423, đợt 2): mẫu là bản đợt trước.** (a) BCKQGD mang **comment Word** của đợt trước (NP1…) →
gỡ `commentRangeStart/End` + run chứa `commentReference` trong `document.xml`, làm rỗng `comments*.xml`. (b) Ảnh gốc khung 4:3 ghi đè
byte trực tiếp được; **đừng để Word `Save()` sau đó** (25 MB → 1 MB, Bẫy 35) — tổng trang đặt thẳng trong XML (`footer4.xml` TBKQ =
tổng − 1, `header1.xml` BCKQGD), Word chỉ mở để đếm trang/xuất PDF. (c) Đoạn sơ đồ TBKQ có **nét vẽ nổi** (`Straight Connector`) của
sơ đồ cũ → xoá. (d) Ô diện tích TBKQ là 3 run `32 ` + `m` + `2` (chỉ số trên) → chỉ thay run đầu. (e) Sửa file `.js` bằng `sed`/`node -e`
trong Git Bash mất `\` (Bẫy 9) → dùng Write/Edit. Script dựng lại: `_claude_tmp\120423d2\js\` (`data.js`, `buildT.js`, `buildB.js`, `post.js`) + `pack.ps1`.

**Bẫy 51 — BCKQ nước uống/sinh hoạt (mẫu KPack) không có khung ảnh nào, ghi "Hình ảnh vị trí lấy mẫu: không có".** Không dùng được 4a/4c
(không có `InlineShapes.Item(1)` làm chuẩn). → Chèn thẳng XML, không qua Word (khỏi Bẫy 35): bảng không viền 3 ảnh/dòng + chú thích ký hiệu
ngay sau đoạn tiêu đề, thêm `Relationship` + entry `word/media/*.jpeg` (byte gốc, md5 khớp `0.*`), đổi `Trang 1/` ghi cứng ở `header1.xml`
thành field PAGE (Bẫy 24) rồi đặt tổng trang theo `test-open`. Chữ "không có" là 4 run (`' '`,`không`,`' '`,`có`) — đọc sót run dấu cách trong
map.txt là vá lệch, còn sót chữ "có". Script: `_claude_tmp\118300\bimg.js` + `pack.ps1`. Gặp ở Q326D118300.

**Bẫy 52 — hồ sơ vi sinh khí nén/khí nitơ (mẫu Iguacu + Olam, 2 bộ BCKQTĐ/TBKQ `- 1` khí nén, `- 2` khí nitơ).** (a) `apply-all` `hl` lên run mẫu
đã tô vàng sẵn → sinh **2 `<w:highlight>` trong một `rPr`** (script không kiểm tra) → gộp lại bằng regex sau khi vá. (b) Mỗi file 2 khung ảnh
chung một `rId` (Bẫy 46) nhưng khung 4:3 đúng tỉ lệ ảnh `0. vi tri` → không cần Word: đổi `r:embed` thứ hai sang `rId901`, thêm `Relationship`
+ entry `word/media/image901.jpeg`, ghi byte gốc (md5 khớp). (c) Dòng 2 bảng vị trí TBKQ `- 1` rỗng cả tên lẫn ký hiệu, bảng kết quả ghi `KN`
thiếu số. Phụ lục vừa khít 1 trang: thêm 1 dòng chữ là đoạn cuối tràn sang trang 3 (footer vẫn `Trang x/1`) → `test-open` sau mỗi lần thêm chữ.
(d) `1. KQ tong dau + vs.xlsx`: sheet vi sinh tên `<6 số>- 4` (2 vị trí/sheet, N* ở `C5`/`E5`) có công thức `#REF!` ở dòng 14–15, 19–20 (trỏ
số lỗ `C9`, thể tích `C10`) và số phiếu/ngày nằm ở **header trang in** (`PageSetup.LeftHeader/RightHeader`), không nằm trong ô; ô nhập `< 1` cho
"Không có chỉ thị" nên làm khung phải thêm nhánh `"………"`/`-`. Script: `_claude_tmp\124242\` (`mkplan.js`, `imgpatch.ps1`, `xfill.ps1`, `xhdr.ps1`). Gặp ở Q326D124242.

**Bẫy 53 — bộ mẫu nước uống + nước DI (BCKQ Sharp, TBKQ `- 1`/`- 4` Jones Lang LaSalle).** (a) BCKQ: đoạn "Kết quả đo tại hiên trường:"
mang `<w15:collapsed/>` → bảng Clo dư/TDS ngay dưới **bị ẩn khi in/xuất PDF** (ô TDS của mẫu còn chứa chữ rác "Không có mùi lạ"); có số đo hiện
trường thì gỡ `w15:collapsed`, bảng hiện ra làm BCKQ tràn trang 3 → rút gọn số phiếu (`.01 ÷ .04`) + bỏ đoạn trống trước "Ngày báo cáo".
(b) Header `- 4` ghi `3.1` trong **một** run, header `- 1` tách `4` + `.1` → chỉ số run ngày lệch 1 giữa hai bộ (đã ra `Ngày 1809/10`,
`final-check` không bắt được) → sau khi vá in lại text mọi `header*.xml`. (c) Dòng Pseudomonas TBKQ VN: `< 1` là run ngay **trước**
"Không phát hiện" — dễ vá nhầm sang cột giới hạn; ô kết quả này in đậm sẵn từ mẫu (đạt thì `unbold`). (d) Bảng kết quả `- 4`: VN là table 3,
EN là table 4 (cùng chỉ số dòng); bản EN lẫn dấu phẩy thập phân ở dòng 17–37 → `set-cells` lại cả cột kết quả; ô kết quả trống của mẫu
(Mo, phenol, xyano clorit, 1,4-DCB, clorit) là ô rỗng thật, không phải ô gộp. Chỉ số trên `(2)` cho giá trị `< LOQ` và ảnh: script
`_claude_tmp\120711\` (`gen.js`, `xml.js`, `repack.ps1`). Gặp ở Q326D120711.

**Bẫy 54 — ghi PID Excel/Word bằng cách so danh sách tiến trình trước/sau `New-Object` khi nhiều agent chạy song song.** Agent khác mở
Excel cùng lúc → hiệu tập hợp ra **2 PID**, `Stop-Process` cả hai là tắt nhầm tiến trình của hồ sơ khác. → Chỉ lấy PID qua
`$x.Hwnd` + `GetWindowThreadProcessId` (Excel) hoặc khi hiệu tập hợp đúng 1 PID; ra nhiều hơn thì **không** tắt, chỉ `Quit()`.
Gặp ở Q326D115269.

**Bẫy 55 — chạy nhiều hồ sơ song song: tắt Excel/Word theo "tiến trình mới nhất" hoặc so danh sách trước/sau là tắt nhầm của agent khác.**
Đợt 115269/115373/124835 (15/10/2026) đã tắt nhầm như vậy (may không mất dữ liệu). → Lấy PID **của chính đối tượng COM** ngay sau
`New-Object` (`GetWindowThreadProcessId` qua `$x.Hwnd`, Word: `$w.ActiveWindow.Hwnd` hoặc so tiêu đề cửa sổ), ghi ra file trong
`_claude_tmp\<6 số>\`, chỉ `Stop-Process` đúng PID đó; script chỉ đọc cũng phải theo quy tắc này. Xong việc mở lại file kiểm tra đã lưu.

**Bẫy 56 — tô vàng dạng `<w:highlight w:val="yellow" />` (có dấu cách trước `/>`) không bị gỡ.** `set-cells` (không `hl`) và
`apply-all` `unhl` đều báo OK nhưng ô vẫn vàng (regex chỉ khớp `"/>`). → Sau khi điền, liệt kê text các run còn `w:highlight`, gỡ bằng
regex `<w:highlight w:val="[^"]*"\s*/>` trên đúng run/bảng vừa điền (`_claude_tmp\116326\hlx.ps1`, `unhl.ps1`). Gặp ở Q326D116326.

**Bẫy 57 — Mất đường kết thúc văn bản (Ending Line) hoặc để sót đường nối mồ côi (Orphaned Connector).**
Trong văn bản TBKQ của QUATEST 3, ở cuối cùng của tài liệu (ngay dưới câu đánh giá/kết luận cuối cùng) luôn có đường kẻ đôi (`thinThick` compound line, 4.5pt strokeweight, màu đen) căn giữa trang.
Khi gán lại nội dung đoạn văn kết luận bằng code (`p.text = ""` hoặc xóa các run), nếu không cẩn thận sẽ vô tình xóa mất thẻ `<mc:AlternateContent>` chứa đối tượng shape vector của đường kết thúc này, hoặc khi sao chép/tái cấu trúc tài liệu sẽ để sót các nét vẽ nổi mồ côi ở các đoạn văn giữa văn bản.
→ **Giải pháp:**
1. Dọn sạch mọi connector mồ côi ở các đoạn văn giữa văn bản và neo đúng run `<w:r>` chứa đường kết thúc `thinThick` vào đoạn kết luận cuối cùng với `<wp:align>center</wp:align>` và offset `wp:posOffset` phù hợp (~580000 EMU cho 2 dòng, ~750000 EMU cho 3 dòng).
2. Chạy `qa_check.py final-check` để tự động xác nhận có đúng 1 đường kết thúc `thinThick` căn giữa ở cuối tài liệu và 0 connector mồ côi.

### Kiểm tra / backup

**Bẫy 13 — `*.orig.bak` của `apply-all.ps1` đặt theo tên file, không theo mã hồ sơ.** Tên mẫu chung
(`3. TBKQ.docx`) đã có `.orig.bak` từ hồ sơ khác thì **không** ghi đè → đó không phải bản gốc của hồ
sơ đang làm. Cần bản gốc → dùng backup có timestamp, hoặc tự copy vào `_claude_tmp\<6 số>\` trước.

**Bẫy 14 — `final-check.ps1 -Old`:** tách theo dấu phẩy (không đưa số thập phân kiểu `25,4` — sẽ
thành `25` và `4`, báo LEFTOVER giả); so khớp **không** phân biệt hoa/thường; từ khoá ngắn trùng
địa danh (`Minh` khớp "Hồ Chí Minh") → dùng cụm dài hơn (`Huy + Minh`). Nghi ngờ thì xác nhận bằng
Word COM / `grep` trên XML đã giải nén.

**Bẫy 16 — phiếu KQTN nhiều trang bị đọc thiếu.** `pdf-render.ps1` từng mặc định chỉ render 20
trang → sót mẫu 11 (blank) của Q326D106280 (PDF 22 trang). Đã nâng mặc định lên 999. Luôn đối chiếu
**số mẫu trên KQTN** với ký hiệu mẫu trên BBLM/Ghi chép (kể cả blank) trước khi kết luận thiếu.

**Bẫy 12 — (đã sửa 18/09/2026)** `test-open.ps1` nay tự tìm cả trong `0.VIMCERTS078`.
Hồ sơ khách quen có thể nằm ở thư mục con riêng (vd `0.1.Tân Hiệp\Q326D104560…`): `apply-all`/`dump-all`/
`final-check` có tìm đệ quy, nhưng `test-open.ps1` **không** → mở thử bằng Word COM trực tiếp.

**Bẫy 48 — mã cũ viết thiếu số 0, mẫu mang nhiều mã cũ, từ khoá `-Old` quá ngắn.** Header TBKQ EN (mẫu KPack / L.S Pack) ghi
`Q326D` + `98698` (5 số) → `-Old` 6 số không bắt được, phải thêm dạng 5 số. Một bộ mẫu có thể mang 3 mã cũ khác nhau
(081961, 031777, 098698) → đưa hết vào `-Old`. `TSI` khớp nhầm trong câu tiếng Anh → dùng `TSI 9350`. Lưu ý thêm:
`rename-done.ps1` đổi tên cả PDF scan `1.*` (vd `1. kq do bui.pdf`), không chỉ docx/xlsx. Gặp ở Q326D124796, Q326D124835.

### File tính toán Excel

**Bẫy 15 — `#VALUE!` ở mục vi sinh file khí nén** — xem mục 7a.

**Bẫy 18 — file tính toán chép từ hồ sơ cũ.** `0. Tinh Toan VSKK.xlsx`, `1. tinh toan tong dau.xlsx`
thường mang số phiếu, ngày, số đo của hồ sơ khác (mtime trước ngày đo là dấu hiệu; đã gặp
Q326D100311, Q326D106280). → Thay **hết** vùng nhập liệu, không chỉ phần kết quả đang cần.

**Bẫy 43 — `0. Tinh Toan VSKK.xlsx` bản hiện dùng KHÁC bảng mục 7b.** Ngày tính toán ở `O1` (không phải `Q1`); kết quả
CFU/m³ ở `N` (tổng số) / `O` (men) / `P` (mốc), `Q:S` là E.coli/Coliform/S.aureus. Ô nhập ghi `-` (chỉ tiêu không làm) đưa vào
`VLOOKUP` → `#N/A` ở `H:S`. → Sửa công thức dòng dữ liệu: `H:M` = `IF(B8="-","-",<công thức cũ>)`, `N:S` = `IF(H8="-","-",…)`.
`#N/A` ở `Z382:AA550` là đuôi bảng Feller (>380 lỗ) có sẵn của mẫu, không phải lỗi nhập. Gặp ở Q326D110631.

**Bẫy 49 — file VSKK vẫn có bản đúng bố cục mục 7b** (`Q1`, kết quả `P/Q/R`), không phải lúc nào cũng là bố cục Bẫy 43 → xem ô
ngày và tiêu đề cột trước khi nhập. Làm khung chưa có KQTN: ô nhập để `………` thì `VLOOKUP` ra `#N/A`, để trống thì ra `< 10` giả
→ thêm nhánh `"………"` (cùng nhánh `-`) vào công thức `I:S` các dòng dữ liệu. Gặp ở Q326D124835.
Cột `T:V` (CFU/m³ của `M:O`) cũng có công thức, không có nhánh `-` → `#VALUE!` khi `F:H`=`-`; quét lỗi phải tới cột `Y`, thêm `IF(M8="-","-",…)`.

---

## 2. Bộ script tại `E:\HoSo\_claude\`

```
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "E:\HoSo\_claude\<script>.ps1" -Code Q326D108033
```

Chỉ các script dưới đây là công cụ dùng chung. Script một lần của từng hồ sơ đã dời vào
`_claude_tmp\_misc\scripts-legacy\` — không gọi lại, viết mới vào `_claude_tmp\<6 số>\` nếu cần.

### Tìm hiểu hồ sơ / đọc dữ liệu

| Script | Việc |
|---|---|
| `startup-check.ps1` | Kiểm tra công cụ + dọn thư mục tạm quá 30 ngày — chạy đầu phiên |
| `dump-all.ps1 -Code … -OutFile …` | Dump chỉ số run `document` + `header1..4` + `footer1..2` của mọi file `1./2./3.` (mặc định ra `E:\HoSo\_claude\map.txt` — nên trỏ `-OutFile` vào `_claude_tmp\<6 số>\`) |
| `set-cells.ps1 -Docx … -Dump` | In bản đồ **bảng/dòng/cột** của docx — dễ hơn chỉ số run khi làm bảng kết quả |
| `pdf-img.js` (Node) | Trích ảnh JPEG nhúng trong PDF scan → `pNN.jpg` (mục 5) |
| `pdf-render.ps1 -Pdf … -OutDir …` | Render trang PDF ra PNG (WinRT) — PDF **khoá quyền copy** (phiếu KQTN, PYC) hoặc PDF vector |
| `img-crop.ps1` | Cắt / xoay / phóng to một vùng ảnh (mục 5) |

### Vá nội dung

| Script | Việc |
|---|---|
| `apply-all.ps1 -PlanFile plan.json` | Vá theo **chỉ số run** (mọi file, mọi entry), `hl` tô vàng, `unhl` gỡ tô vàng; tự backup |
| `set-cells.ps1 -Docx … -DataJson …` | Vá theo **(table,row,col)** 0-based — cả ô RỖNG (tự chèn run). `hl`, `b` (đậm), `unbold` |
| `dup-row.ps1 -Table n -Row m -Count k` | Nhân bản dòng bảng k lần (giữ định dạng) |
| `del-row.ps1 -Table n -Rows "2,3"` | Xoá dòng bảng theo chỉ số |
| `del-tables.ps1 -Tables "3,4" [-WithHeading]` | Xoá cả bảng, kèm đoạn tiêu đề ngay trước |
| `delete-rows.ps1 -StartMarker … -EndMarker …` | Xoá dòng theo chuỗi mốc (mốc thuần ASCII — Bẫy 1) |
| `rename-done.ps1 -Code …` | Thêm đuôi mã hồ sơ vào mọi file `1./2./3.` |
| `kn-fill-excel.ps1` | Mẫu nhập liệu Excel qua COM từ JSON (mục 7a) — sửa mã hồ sơ + đường dẫn JSON trước khi chạy |

### Ảnh

| Script | Việc |
|---|---|
| `overwrite-images.ps1 -Docx … -MapJson …` | Ghi đè byte ảnh trong `word/media/` (mục 4) |
| `resize-img.ps1 -In <file\|folder> -OutDir … [-MaxW 1400]` | Thu nhỏ ảnh điện thoại 3–4 MB → ~250 KB |

### Kiểm tra

| Script | Việc |
|---|---|
| `final-check.ps1 -Code … -Old "063668,22/06/2026,…"` | Quét `document.xml` **và** mọi header/footer tìm sót dấu vết hồ sơ cũ |
| `test-open.ps1 -Num <6 số>` | Mở mọi docx bằng Word COM, in số trang — bắt "unreadable content" (Bẫy 6), đối chiếu tổng trang (Bẫy 11) |
| `ps-calc.ps1` | Tính nồng độ hạt từ CSV thô — chỉ khi không có bản scan tính toán |

`apply-all.ps1`, `final-check.ps1`, `test-open.ps1` tự tìm cả trong `0.VIMCERTS078`.

### `plan.json` (cho `apply-all.ps1`)

```json
{ "code": "Q326D108033",
  "jobs": [
    { "file": "2. BCKQ - 2 - khi nen.docx", "entry": "word/document.xml",
      "set": { "1": "Q326D", "2": "108033/2" }, "hl": [132, 137], "unhl": [140] },
    { "file": "3. TBKQ - vn - 2 - khi nen.docx", "entry": "word/header3.xml",
      "set": { "1": "Q326D", "2": "108033/2" }, "hl": [4] }
  ] }
```

`hl`: tô vàng ô còn thiếu (`………`). `unhl`: gỡ tô vàng ô vừa điền đủ (vd khi có KQTN).

### `DataJson` (cho `set-cells.ps1`)

```json
{ "cells": [
    { "table": 3, "row": 2, "col": 4, "text": "0,05", "b": true },
    { "table": 3, "row": 3, "col": 4, "text": "………", "hl": true },
    { "table": 3, "row": 4, "col": 4, "text": "0,64", "unbold": true }
  ] }
```

`set-cells` ghi text vào `<w:t>` đầu tiên của ô và **xoá** các `<w:t>` còn lại → không dùng cho ô
nhiều đoạn (vd danh sách hoạt chất BVTV) — dùng `apply-all` theo từng run.

**Vì sao vá theo chỉ số run:** Word cắt một câu thành nhiều `<w:r>/<w:t>` (vd `Q326D0` + `63668/1`),
find-replace theo chuỗi hay trượt. Vá theo chỉ số giữ nguyên 100% style, font, bảng. Backup tự
động vào `E:\HoSo\_claude\backup\`.

**Tái sử dụng bản đồ chỉ số run:** file mẫu thường copy từ hồ sơ cũ **cùng khách hàng** → chỉ số gần
như y hệt. Lưu `plan.json` theo tên khách vào `_claude\maps\<khach-hang>.json`. Vẫn `dump-all`
xác minh vài chỉ số mốc trước khi vá.

---

## 3. Công thức (chỉ khi phải tự tính)

Nồng độ hạt tích lũy (hạt/m³) từ `database_all.csv` (phân cách `;`, thập phân dấu chấm — parse
bằng `InvariantCulture`; luôn lọc theo `Location` / `Data ID`):

```
d>=0,5 = (Counts6 + Counts7 + Counts8) / SampleVolume
d>=1,0 = (Counts7 + Counts8) / SampleVolume
d>=5,0 =  Counts8 / SampleVolume
```

Kênh `Counts1..8` = 0,1 / 0,15 / 0,2 / 0,25 / 0,3 / 0,5 / 1,0 / 5,0 µm.

Giới hạn ISO 14644-1:2015 cấp **ISO 8**: d>=0,5 → 3 520 000 ; d>=1 → 832 000 ; d>=5 → 29 300 hạt/m³.

**Mẫu hỏng:** dòng có `Sample Volume` lệch hẳn (vd 0,0108 m³ so với 0,0283 m³ chuẩn) là mẫu lỗi —
bản in đã loại thì phải loại theo.

Khí nén: quy đổi `N* = (1/Pdd) x (273+Tdd)/(273+20)`, phân loại ISO 8573-1 — xem mục 7a.

---

## 4. Ảnh trong docx

### 4a. Thay ảnh có sẵn (cách an toàn nhất — không đụng XML)

1. Liệt kê `word/media/*` trong docx.
2. Ghép thứ tự ảnh với thứ tự `r:embed` trong `document.xml` → `word/_rels/document.xml.rels` —
   **đừng đoán theo tên file media**.
3. `overwrite-images.ps1`: ghi đè byte entry `word/media/imageN.<ext>`, giữ nguyên tên entry →
   không cần sửa `.rels`/`document.xml`, khung ảnh giữ nguyên.
4. Không có ảnh mới → ghi đè bằng **ảnh trắng**, không giữ ảnh cũ.
   **Khung có `<a:srcRect …/>` (mẫu cũ đã crop ảnh trong Word)** → ghi đè byte vẫn bị crop theo tỉ lệ cũ,
   ảnh mới mất một phần (gặp ở Q326D112853). Kiểm tra `srcRect` trước; có thì bỏ `srcRect` hoặc dựng lại khung.
   **Quy ước: ảnh vị trí luôn dùng NGUYÊN ảnh gốc `0.*` — giữ nguyên kích thước pixel (không resample,
   không nén lại, không thu nhỏ file trước khi nhúng), không crop, không độn nền trắng, không méo; chỉ
   thu gọn kích thước hiển thị (`Width`/`Height`, khoá tỉ lệ) cho vừa khung.** Cách làm (Word COM): ghi `Width` của khung cũ, `Delete()` ảnh cũ, `AddPicture`
   ảnh gốc vào đúng range, `LockAspectRatio=-1`; ảnh ngang đặt `Width`=khung cũ, ảnh dọc đặt `Height`=khung cũ.
   Kiểm tra crop = 0, số trang không đổi, `Content.Text` trước/sau giống hệt, rồi xuất PDF xem. Lưu ý Bẫy 7:
   biến `$W` (độ rộng) trùng `$w` (Word.Application) — đặt tên khác. Gặp ở Q326D112853.
   Word COM tự nén ảnh khi lưu → sau khi chèn phải ghi đè lại byte gốc: xem **Bẫy 35**.
5. Kiểm tra nhanh ảnh đang nhúng có phải của hồ sơ này chưa: giải nén `word/media/imageN` ra
   `_claude_tmp` rồi so bằng mắt với ảnh `0.*` (Q326D100311 đã đúng sẵn, không cần thay).

### 4b. Chèn ảnh mới (khi mẫu chưa có khung ảnh)

Thêm entry `word/media/<ten>.jpeg`, thêm `Relationship` vào `document.xml.rels`, chèn khối
`<w:drawing>` vào `document.xml` (mẫu cũ: `_claude_tmp\_misc\scripts-legacy\insert-images-*.ps1`).

- **Trước khi chèn phải xoá hết ảnh cũ** (entry + `Relationship` + `<w:drawing>`/`<w:pic>`).
- **Đếm `<w:t>` KHÔNG đủ để kết luận ô rỗng:** ô chứa `<w:drawing>` không sinh `<w:t>` → đã gặp
  chèn thêm ảnh vào ô vốn có ảnh, chứng thư hiện ảnh **lặp đôi**. Đếm riêng `<w:drawing>`; đã có
  ảnh → dùng 4a.
- Đuôi entry khớp `[Content_Types].xml` (Bẫy 6), rồi `test-open.ps1`.

---

### 4c. Tăng hoặc giảm số lượng khung ảnh mapping vị trí lấy mẫu & hình ảnh thực hiện (BCKQ và TBKQ)

Số lượng hình ảnh mapping / sơ đồ vị trí lấy mẫu (quan trắc) và hình ảnh quá trình thực hiện có thể **tăng hoặc giảm** theo từng hồ sơ:

- **Khi số ảnh tăng (nhiều hơn mẫu):**
  - *Ảnh nằm ngoài bảng (InlineShapes nối tiếp):* Dùng Word COM (Word tự khai báo media/rels/content types — tránh Bẫy 6): lấy `$pic=$doc.InlineShapes.Item(1)` làm chuẩn, chèn ngay sau nó `$doc.InlineShapes.AddPicture($img,$false,$true,$range)` cho từng mẫu, đặt `LockAspectRatio=-1` (hoặc kích thước khung chuẩn), 2 ảnh/dòng cách nhau 2 dấu cách, xuống dòng bằng `[char]11`. Khung cũ có chú thích thì nhân bản chú thích theo.
  - *Ảnh / sơ đồ mapping nằm trong bảng không viền (1x2, 2x3…):* Nhân bản dòng bảng (`dup-row.ps1` + chạy `fix-docpr-ids.ps1` theo Bẫy 34) hoặc thêm dòng/ô mới cùng kích thước, sau đó thay từng ảnh và dòng chú thích tương ứng.
- **Khi số ảnh giảm (ít hơn mẫu):**
  - Xoá bớt dòng bảng dư (`del-row.ps1`) hoặc xoá khung ảnh/chú thích ở các ô thừa (nếu bảng 1x2 chỉ còn 1 ảnh thì xoá cột/ô thừa hoặc dùng khung 1x1 căn giữa), tuyệt đối không để sót ảnh hồ sơ cũ hoặc để ô trống thừa.
- Sau khi thêm/bớt ảnh: ghi đè lại byte ảnh gốc nếu đi qua Word COM (Bẫy 35), xuất PDF (`SaveAs2(...,17)`) + `pdf-render` xem thứ tự ảnh và bố cục trang, `ComputeStatistics(2)` đối chiếu tổng trang header/footer (Bẫy 11, Bẫy 24), `test-open`. Mẫu script: `E:\HoSo\_claude\maps\add-images-bckq-106685.ps1`.

### 4d. Quy tắc chuyển trang: Giữ câu kết luận nằm cùng trang & lặp lại tiêu đề bảng

Khi bảng kết quả có câu kết luận đi kèm (`Đánh giá:` / `Assessment:` + câu kết luận, kèm bảng `Ghi chú:` / `Note:` ngay dưới bảng):
- **Bắt buộc câu kết luận nằm cùng trang với bảng** (không để bảng kết thúc ở cuối trang trước rồi đẩy riêng `Đánh giá:` / câu kết luận sang đầu trang tiếp theo).
- **Lặp lại dòng tiêu đề khi sang trang mới (`Repeat Header Rows`):** Mọi bảng dữ liệu/kết quả khi ngắt trang bắt buộc phải có thuộc tính lặp lại hàng tiêu đề (`<w:tblHeader/>`) để dễ theo dõi các cột.
- **Cách xử lý trong XML / Word:**
  - Đặt `<w:keepNext/>` trong `<w:pPr>` cho: (1) tiêu đề mục ngay trên bảng, (2) các đoạn `<w:p>` ở 1–2 dòng cuối của bảng kết quả, (3) mọi đoạn `<w:p>` trong bảng `Ghi chú:` / `Note:` ngay dưới bảng, và (4) đoạn tiêu đề `Đánh giá:` / `Assessment:`.
  - Đặt `<w:keepLines/>` cho đoạn văn câu kết luận và `<w:cantSplit/>` cho các dòng `<w:trPr>` của bảng.
  - Đặt `<w:tblHeader/>` trong thẻ `<w:trPr>` của (các) dòng tiêu đề cột của bảng.
  - Luôn kiểm tra lại khi xuất PDF / đếm trang (`test-open.ps1`).

### 4e. Định dạng in đậm và nghiêng cho cụm từ kết luận đánh giá

Khi điền hoặc tạo mới câu đánh giá/kết luận sau bảng kết quả (`Đánh giá:` / `Assessment:`):
- Cụm từ đánh giá kết quả bắt buộc phải được **in đậm và in nghiêng** (`bold=True, italic=True`):
  + Tiếng Việt: **_phù hợp_**, **_không phù hợp_**
  + Tiếng Anh: **_comply_**, **_complied_**, **_did not comply_**
- Các từ ngữ xung quanh trong câu giữ nguyên định dạng in thường (Times New Roman, 13pt).

### 4f. Quy tắc Số chứng thư trên Header vs Số hồ sơ trong thân văn bản (Table 0)

- **Đồng bộ Số chứng thư trên mọi Header:** Mọi header trong tài liệu (Section 0 và Section 1) bắt buộc có cùng một Số chứng thư (vd: `Số: Q326D122266/2 - Ngày 20/10/2026` hoặc `No: Q326D122266/2/E – Date: 20/10/2026`).
- **Số hồ sơ trong thân văn bản (Table 0 & Body):** Bắt buộc ghi **mã trần (mã gốc)**, ví dụ `Q326D122266`. Tuyệt đối **KHÔNG** ghi đuôi `/1`, `/2`, `/3`, `/E` vào thân văn bản.
- Giữ nguyên nhãn ô của Table 0 (ô `Giấy yêu cầu và nhận giám định số:` / `Inspection application no. :`), không ghi đè làm mất nhãn.

### 4g. Quy tắc đồng bộ Tên vị trí, Thiết bị và Số liệu (Tiếng Việt & Tiếng Anh)

- **Tên vị trí thực hiện:** Bắt buộc phải đồng nhất số lượng, thứ tự vị trí và mô tả vị trí giữa bản tiếng Việt và bản tiếng Anh theo bản dịch chuẩn (ví dụ: VT1, VT2... hoặc tên vị trí cụ thể). Tuyệt đối không để lệch vị trí hay đảo lộn thứ tự.
- **Danh mục thiết bị:** Danh mục thiết bị sử dụng giữa các bảng tiếng Việt và tiếng Anh **phải hoàn toàn giống nhau** (cùng chủng loại thiết bị, số lượng và quy cách tương ứng).
- **Số liệu đo đạc:** Toàn bộ số liệu điền vào giữa hai bản **phải hoàn toàn trùng khớp**.
- **Biểu diễn số thập phân:** Bản tiếng Việt bắt buộc dùng **dấu phẩy (`,`)** (vd: `23,5`), bản tiếng Anh bắt buộc dùng **dấu chấm (`.`)** (vd: `23.5`).

## 5. Đọc scan, chữ viết tay, phiếu KQTN

Máy không có `pdftoppm`/Python nên Read không mở trực tiếp PDF. Ba đường, theo thứ tự thử:

1. **PDF scan có ảnh JPEG nhúng** (BBLM, Ghi chép, bản scan tính toán) → `node pdf-img.js "<pdf>" "<ra>"`.
2. **PDF text thường** → mở bằng Word COM lấy chữ.
3. **PDF khoá quyền copy / vector** (phiếu KQTN, PYC — `pdf-img.js` báo `anh trich duoc: 0`) →
   `pdf-render.ps1`. Kiểm tra số trang: `grep -a "/Count"` trên file PDF.

Phóng to / xoay vùng chữ tay (BBLM hay bị xoay 90°):

```
powershell -NoProfile -ExecutionPolicy Bypass -File "E:\HoSo\_claude\img-crop.ps1" `
  -In "...\p01.jpg" -Out "...\ten.png" -X 1950 -Y 250 -W 350 -H 480 -Rotate 270 -Scale 1.5
```

`-Rotate` 0/90/180/270 theo chiều kim đồng hồ (BBLM quét dọc thường cần **270**; ra ngược thì đổi
90). Toạ độ tính theo **ảnh gốc** — Read hiển thị ảnh thu nhỏ kèm hệ số, nhân vào.

**Đọc phiếu KQTN nhiều mẫu tiết kiệm token:** mỗi mẫu 2 trang (trang lẻ: tên mẫu, trang chẵn: bảng
kết quả). Render xong, `img-crop` riêng dải đầu trang (`-X 100 -Y 270 -W 1450 -H 420 -Scale 0.5`)
rồi mới Read — đủ đọc số phiếu `.NN`, tên mẫu, bảng kết quả.

**Hàng chục phiếu cùng bố cục (vd 32 máy lọc LEGO Q326D100120):** render hết (`-Width 1200`), rồi ghép
cột "Kết quả" (x≈770–950) của 16 mẫu cạnh nhau + cột tên chỉ tiêu của mẫu đầu thành một ảnh bằng
System.Drawing (`E:\HoSo\_claude\montage.ps1 -Root <thư mục render> -Dirs "a,b" -Page p02.png -X 770 -Y 390 -W 180 -H 1010 -Out <png> -LabelW 350 -LabelDir a`) → một lần Read đọc 16 mẫu. Ghép dải tên mẫu
trang 1 để đối chiếu `.NN` ↔ ký hiệu. KQTN khoá quyền copy nên Word COM không mở được PDF.

**Nguyên tắc:** chữ không chắc chắn → không đoán, để `………` + tô vàng, báo "đọc được là X nhưng chưa
chắc". Đối chiếu chéo giờ BBLM với dấu thời gian máy đo, số vị trí với phiếu tính toán; lệch → báo.
Ảnh trích ra để trong `_claude_tmp\<6 số>\`, không để trong thư mục hồ sơ.

---

## 6. TBKQ dạng Excel (tủ BSC và các mẫu .xlsx khác)

Kiểm chứng trên `Q326D108306` (Y khoa Thủ Đức).

- **File tính toán đã liên kết công thức sang TBKQ** → vá thông tin khách hàng là chưa đủ: sheet
  nguồn (thường "Bảng tính") vẫn mang **số liệu đo thô của hồ sơ cũ**. Đối chiếu từng số với scan
  tính toán của hồ sơ đang làm và vá hết — kể cả số không còn công thức nào tham chiếu.
- Trước khi kết luận "đã liên kết", mở bằng Excel COM kiểm tra công thức ô kết quả; không có liên
  kết → vá tay kết quả như thường.
- **Ngưỡng sát biên:** kết quả đúng bằng ngưỡng mà công thức dùng `>` ra "Không đạt" sai → sửa
  **phép so sánh**, không chỉnh số đo.
- **Không sửa zip .xlsx tại chỗ** bằng `ZipArchiveMode.Update` + `Delete()`/`CreateEntry()` — hỏng
  file. Ghi lại toàn bộ zip mới theo đúng thứ tự entry, hoặc sửa qua Excel COM. Xác nhận bằng Excel COM.

---

## 7. File tính toán Excel theo loại hồ sơ

Chung cho mọi loại:
- Nhập qua Excel COM (Bẫy 17), backup trước vào `backup\<tên> - <mã>_<timestamp>.xlsx`.
- Sau khi nhập: quét ô lỗi (`#VALUE!`, `#N/A`…) toàn workbook, đọc kết quả, **đối chiếu với bản
  scan / phiếu KQTN**; lệch → báo người dùng, chứng thư theo bản scan.

### 7a. Khí nén — `1. tinh toan tong dau.xlsx`

Mọi tính toán trong sheet **`Tinh_Toan_Khi_Co_Ap_V6`**. Trang 1 (B1:K17) là bảng tổng hợp + phân
loại ISO 8573-1 — **chỉ công thức, không gõ tay**. Trang 2 (từ dòng 33) là vùng nhập liệu.

| Nhóm (trang 2) | Ô nhập | Nguồn |
|---|---|---|
| Quy đổi N* | `C34` = `DiemSuong!E5` (tự lấy), `C35` áp suất đo = 1 (đo sau bộ giảm áp) | Scan điểm sương |
| Điểm sương | Sheet `DiemSuong`: `C2` tên máy, `C4:D4` T1 đầu/cuối, `C5:D5` nhiệt độ khí, `C6:D6` áp suất, `G4` điểm sương quy đổi 8573 (phần mềm CS), `C9`/`C20` giờ đầu/cuối | Scan điểm sương |
| 1. Tổng dầu | `C41:D42` lưu lượng/thời gian dầu hạt – dầu hơi; `C44` KQTN dầu hạt, `D44` dầu hơi (µg/mẫu, KPH ghi `KPH`); `J41`/`J42` số phiếu | KQTN HDTN (dầu hạt), MTTN (dầu hơi) |
| 2. Nồng độ hạt | `C52:J52` theo nhóm ISO, `C53` hệ số pha loãng | Scan tính toán bụi |
| 3. Vi sinh | `C60` số lỗ, `C61` thể tích; `C64:G64` Tổng số/Men/Mốc/Coliform (`< 1` không phát hiện, `-` không làm); `I60` số phiếu | KQTN VSTN |

- **Nồng độ hạt AeroTrak 9350** (chế độ Δ, kênh 0,3/0,5/1/2/3/5 µm): `C52:F52` = 0, `G52` = kênh
  0,3; `H52` = kênh 0,5; `I52` = tổng kênh 1+2+3; `J52` = kênh 5.
- Có hạt d > 5,0 µm → **"Không phân loại"** (EN: *Not classified*).
- Cả dầu hạt và dầu hơi KPH → công thức ra `< 0,01` – Class 1 (lưu ý: giới hạn phát hiện dầu hơi
  5 µg/240 L ≈ 0,021 mg/m³ — đã báo người dùng ở Q326D100311).
- **Bẫy 15:** khách không yêu cầu vi sinh → dòng 64 ghi `-`, công thức Feller dòng 65 đưa `-` vào
  `LN()` → `#VALUE!`. Lỗi công thức mẫu, **không phải file hỏng** → sửa dòng 65 coi `-`/`< 1` là 0,
  dòng 66 trả `-` khi ô nhập là `-`. Không điền số giả.
- **Ảnh trong sheet** (ảnh "Place in cell"): `DiemSuong!B8`/`B19` ảnh màn hình máy đo đầu/cuối,
  ảnh phần mềm CS đặt tại `D9`; trang 1 `B17:D28` ảnh vị trí. Nguồn: file
  `thuc hien\kq diem suong\<mã>- KN.xlsx` (cùng bố cục — `Range.Copy` B8:B17, B19:B28; ảnh CS là
  shape `Picture 2` → Copy/Paste khi `$x.Visible=$true`) và `0. kn.jpg` (`Shapes.AddPicture`, giữ tỉ lệ).
  Ảnh chụp thư mục `D2609xx\BMxxxxx.jpg` chụp ~1,5 phút/ảnh, **không** đúng giờ in trên scan → lấy từ file KN.xlsx.
- Excel có thể lệch ±1 so với scan do scan làm tròn N* (vd 1,048) — chứng thư theo scan.
- Sơ đồ hệ thống khí nén: giữ nguyên từ hồ sơ cũ cùng khách hàng nếu không đổi.
- Mẫu JSON: `E:\HoSo\_claude\maps\kn-fill-Q326D100311.json`.

### 7b. Vi sinh không khí — `0. Tinh Toan VSKK.xlsx`

Sheet `Sheet1`, thiết bị Buck Bio-Culture B30120 (bảng hiệu chỉnh Feller 380 lỗ ở `AB:AD`).

| Ô | Nhập |
|---|---|
| `A1` | `Số: <mã hồ sơ>` |
| `Q1` | Ngày tính toán (`=DATE(yyyy,m,d)`) — theo ngày KQTN |
| `D3` / `D4` | Lưu lượng (L/phút) / thời gian lấy mẫu (phút) — từ phiếu Ghi chép → `D5` thể tích m³ tự tính |
| `A8:A…` | Ký hiệu vị trí (KK1…) |
| `B` / `C` / `D` | Tổng số / Men / Mốc (CFU/đĩa từ KQTN; không phát hiện ghi `< 1`) |
| `E:H` | Chỉ tiêu không làm ghi `-` |

- Kết quả CFU/m³ ở `P` (tổng số), `Q` (men), `R` (mốc) → đưa vào TBKQ đúng thứ tự cột của mẫu
  (TBKQ Sanyo: Hiếu khí / Men / Mốc). Không phát hiện → `< 10` (với 0,1 m³).
- Mẫu chỉ có công thức sẵn tới một số dòng — dòng thiếu thì `Range('I16:V16').Copy(I17:V17)`.
- Mẫu blank (vd số 11 trên Ghi chép) có phiếu `.11` riêng trong KQTN → ghi vào dòng "KQTN của mẫu
  Blank" ở BCKQTĐ, không đưa vào bảng TBKQ.


---

## 8. Kiến trúc Module Tái Sử Dụng (QUATEST 3 Core Framework)

Dự án đã được tối ưu và đóng gói thành hệ thống module dùng chung tại thư mục `Core/`, loại bỏ hoàn toàn việc phải viết lại mã nguồn xử lý Word (OpenXML / Styling / Table / PDF) cho mỗi hồ sơ:

### Cấu trúc Module (`Core/`):
- `Core.Models`: Các lớp dữ liệu chuẩn hóa (`Dossier`, `ClientInfo`, `Equipment`, `PersonnelInfo`, `MonitoringLocation`). Tự động xử lý người ký và người báo cáo nội bộ theo quy tắc QUATEST 3 (Section 2 & 7 trong `AGENTS.md`).
- `Core.Styling`: Các module xử lý OpenXML chuyên sâu:
  + `styles.py`: Font Times New Roman 13pt/10pt, margins, padding, vAlign.
  + `tables.py`: `set_table_autofit_window` (bắt buộc cho bảng ảnh và layout), `set_table_row_height` (0,7 - 0,8 cm), `add_tbl_header` (lặp header khi ngắt trang), `add_cant_split`.
  + `paragraphs.py`: `set_paragraph_text`, `set_assessment_paragraph` (tự động format ***phù hợp***, ***không phù hợp***, ***comply***, ***complied*** in đậm + in nghiêng).
  + `elements.py`: `add_document_ending_line` (đường kẻ đôi `thinThick` 4.5pt căn giữa, chuẩn xác 100% không đè chữ), `strip_all_highlights`, `update_headers`, `update_table0` (mã trần không đuôi, bold=False).
- `Core.Builders`:
  + `bckq_builder.py`: Sinh Báo cáo kết quả nội bộ BM3030.03 đúng 2 trang, không có đường Ending Line.
  + `tbkq_builder.py`: Điều phối sinh TBKQ Tiếng Việt và Tiếng Anh.
  + `handlers/`: Handlers cho từng loại hình quan trắc (`laf_bsc.py`, `cleanroom.py`, `compressed_air.py`). Tự động chuẩn hóa dấu thập phân (VN dùng dấu phẩy `,`, EN dùng dấu chấm `.`).
- `Core.Exporters`:
  + `pdf_exporter.py`: Chuyển đổi DOCX sang PDF bằng Microsoft Word COM chuẩn xác, tự động diệt process treo mồ côi, lưu và đồng bộ trực tiếp vào `Input\<tên hồ sơ>\TBKQ\`, `Output\` và ổ mạng Z.
- `Core.Services`:
  + `dossier_runner.py`: Orchestrator chạy toàn trình 4 bước: Sinh BCKQ -> Sinh TBKQ VN/EN -> Export PDF & Sync -> Chạy QA verification (`qa_check.py`).

### Hướng dẫn sử dụng cho hồ sơ mới:
```python
from Core import Dossier, ClientInfo, Equipment, PersonnelInfo, MonitoringLocation, DossierRunner

dossier = Dossier(
    code="Q326D126588",
    cert_date="21/10/2026",
    sample_date="16/10/2026",
    dossier_type="laf", # 'laf', 'cleanroom', 'compressed_air'
    client=ClientInfo(name_vn="...", address_vn="..."),
    personnel=PersonnelInfo(executor="Phát"),
    equipments=[...],
    locations=[...],
    photos=[...],
    layouts=[...],
    results_data={...}
)

runner = DossierRunner(dossier)
runner.run_all()
```

### Chạy qua CLI:
- LAF Profile: `python Core/run.py --sample-laf`
- Khí nén Profile: `python Core/run.py --sample-khinen`
- File JSON: `python Core/run.py --profile profile.json`
