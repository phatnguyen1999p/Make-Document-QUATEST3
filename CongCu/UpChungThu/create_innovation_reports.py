# -*- coding: utf-8 -*-
"""
Script to generate:
1. BM23.01.01 Báo cáo sáng kiến - Phần mềm UpChungThu - Nguyễn Tiến Phát.docx
2. BM23.01.02 Bảng tổng hợp sáng kiến - UpChungThu.xls
3. Mau 03 - Phu luc I - UpChungThu.xls
"""

import os
import sys
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import win32com.client

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
REF_DIR = os.path.join(BASE_DIR, "Báo cáo sáng kiến", "Reference")
OUT_DIR = os.path.join(BASE_DIR, "Báo cáo sáng kiến")
IMG_DIR = os.path.join(OUT_DIR, "Images")

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def add_styled_paragraph(doc, text="", style='Normal', space_before=0, space_after=4, line_spacing=1.2, align=WD_ALIGN_PARAGRAPH.JUSTIFY, bold=False, italic=False, font_size=13, font_name="Times New Roman"):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    
    if text:
        run = p.add_run(text)
        run.bold = bold
        run.italic = italic
        run.font.size = Pt(font_size)
        run.font.name = font_name
        rPr = run._r.get_or_add_rPr()
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:ascii'), font_name)
        rFonts.set(qn('w:hAnsi'), font_name)
        rFonts.set(qn('w:cs'), font_name)
        rPr.append(rFonts)
    return p

def add_run(paragraph, text, bold=False, italic=False, font_size=13, font_name="Times New Roman", color=None):
    run = paragraph.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(font_size)
    run.font.name = font_name
    if color:
        run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:cs'), font_name)
    rPr.append(rFonts)
    return run

def set_table_borders(table, color="D3D3D3", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'  <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'  <w:left w:val="none"/>\n'
        f'  <w:right w:val="none"/>\n'
        f'  <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'  <w:insideV w:val="none"/>\n'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

def add_figure(doc, img_filename, caption_prefix, caption_desc, width_inches=6.2):
    """Insert a centered illustrative screenshot with a keep-with-next caption underneath."""
    img_path = os.path.join(IMG_DIR, img_filename)
    if not os.path.exists(img_path):
        print(f"[Cảnh báo] Không tìm thấy ảnh minh họa: {img_path}")
        return
    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(6)
    p_img.paragraph_format.space_after = Pt(2)
    p_img.paragraph_format.keep_with_next = True
    run_img = p_img.add_run()
    run_img.add_picture(img_path, width=Inches(width_inches))

    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.space_before = Pt(2)
    p_cap.paragraph_format.space_after = Pt(8)
    p_cap.paragraph_format.line_spacing = 1.15
    add_run(p_cap, caption_prefix, bold=True, italic=True, font_size=11, color=RGBColor(0, 51, 102))
    add_run(p_cap, caption_desc, bold=False, italic=True, font_size=11, color=RGBColor(40, 40, 40))

def build_docx_report():
    print("--- Đang xây dựng báo cáo Word BM23.01.01 ---")
    doc = docx.Document()
    
    # Page setup: A4, margin top 1.5cm, bottom 1.5cm, left 2.0cm, right 1.5cm
    section = doc.sections[0]
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(0.79)
    section.bottom_margin = Inches(0.79)
    section.left_margin = Inches(0.98)
    section.right_margin = Inches(0.79)
    
    # Header Table (Quốc hiệu & Đơn vị)
    t_header = doc.add_table(rows=2, cols=2)
    t_header.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_header.autofit = False
    
    col_widths = [Inches(3.3), Inches(3.2)]
    for row in t_header.rows:
        for idx, width in enumerate(col_widths):
            row.cells[idx].width = width
            
    # Cell (0, 0): Đơn vị
    c00 = t_header.cell(0, 0)
    p00 = c00.paragraphs[0]
    p00.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p00.paragraph_format.line_spacing = 1.15
    p00.paragraph_format.space_after = Pt(2)
    add_run(p00, "ỦY BAN TIÊU CHUẨN\nĐO LƯỜNG CHẤT LƯỢNG QUỐC GIA\n", bold=False, font_size=11)
    add_run(p00, "TRUNG TÂM KỸ THUẬT 3", bold=True, font_size=11)
    
    # Cell (0, 1): Quốc hiệu
    c01 = t_header.cell(0, 1)
    p01 = c01.paragraphs[0]
    p01.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p01.paragraph_format.line_spacing = 1.15
    p01.paragraph_format.space_after = Pt(2)
    add_run(p01, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n", bold=True, font_size=11)
    add_run(p01, "Độc lập - Tự do - Hạnh phúc", bold=True, font_size=11)
    
    # Cell (1, 1): Địa danh ngày tháng
    c11 = t_header.cell(1, 1)
    p11 = c11.paragraphs[0]
    p11.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p11.paragraph_format.space_before = Pt(4)
    p11.paragraph_format.space_after = Pt(6)
    add_run(p11, "TP. Hồ Chí Minh, ngày 15 tháng 10 năm 2026", italic=True, font_size=11)
    
    # Tiêu đề báo cáo
    p_title = add_styled_paragraph(doc, space_before=14, space_after=14, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_run(p_title, "BÁO CÁO KẾT QUẢ SÁNG KIẾN", bold=True, font_size=14)
    
    # 1. Thông tin cá nhân
    p_info_title = add_styled_paragraph(doc, space_before=6, space_after=4, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_run(p_info_title, "Thông tin cá nhân tham gia sáng kiến:", bold=True, font_size=13)
    
    p_author = add_styled_paragraph(doc, space_before=2, space_after=2, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_run(p_author, "• Họ và tên: ", bold=False, font_size=13)
    add_run(p_author, "Nguyễn Tiến Phát", bold=True, font_size=13)
    add_run(p_author, "\t- Mã NV: ", bold=False, font_size=13)
    add_run(p_author, "953N224", bold=True, font_size=13)
    add_run(p_author, "\t- Phòng: ", bold=False, font_size=13)
    add_run(p_author, "Phòng Nghiệp vụ 2", bold=True, font_size=13)
    add_run(p_author, "\t- Chức vụ: ", bold=False, font_size=13)
    add_run(p_author, "Giám định viên", bold=True, font_size=13)
    
    p_role = add_styled_paragraph(doc, space_before=2, space_after=4, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_run(p_role, "• Vai trò tham gia: ", bold=False, font_size=13)
    add_run(p_role, "Tác giả ý tưởng, chủ trì nghiên cứu và xây dựng giải pháp (Tỷ lệ đóng góp: 100%)", bold=True, italic=True, font_size=13)
    
    p_note_role = add_styled_paragraph(doc, space_before=0, space_after=8, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_run(p_note_role, "(Ghi chú vai trò: Đối với sáng kiến là giải pháp: ý tưởng/chủ trì/tham gia; Đối với sáng kiến là đề tài/đề án/công trình KH&CN: chủ trì/thành viên)", italic=True, font_size=11, color=RGBColor(80, 80, 80))
    
    # 2. Tên sáng kiến
    p_name_title = add_styled_paragraph(doc, space_before=6, space_after=4, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_run(p_name_title, "Tên sáng kiến: ", bold=True, font_size=13)
    add_run(p_name_title, "“XÂY DỰNG VÀ ỨNG DỤNG PHẦN MỀM TỰ ĐỘNG HÓA XỬ LÝ VĂN BẢN, XUẤT CHỨNG THƯ ĐIỆN TỬ VÀ KIỂM SOÁT TÍNH TOÀN VẸN HỒ SƠ GIÁM ĐỊNH (UPCHUNGTHU)”", bold=True, font_size=13, color=RGBColor(0, 51, 102))
    
    # 3. Nội dung sáng kiến
    p_content_title = add_styled_paragraph(doc, space_before=10, space_after=4, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_run(p_content_title, "Nội dung sáng kiến:", bold=True, font_size=13)
    
    p_sub1 = add_styled_paragraph(doc, space_before=4, space_after=4, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_run(p_sub1, "3.1. Hiện trạng trước khi áp dụng sáng kiến:", bold=True, italic=True, font_size=13)
    
    p = add_styled_paragraph(doc)
    add_run(p, "Tại Phòng Nghiệp vụ 2 (đặc biệt là Nhóm Quan trắc Môi trường) thuộc Trung tâm Kỹ thuật Tiêu chuẩn Đo lường Chất lượng 3 (QUATEST 3), khối lượng tài liệu kỹ thuật, biên bản kết quả thử nghiệm và chứng thư giám định phát hành hàng ngày là rất lớn. Mỗi hồ sơ giám định thường bao gồm nhiều tài liệu thành phần như: thông báo kết quả (TBKQ), các kết quả phân nhánh đánh số theo đợt (/1, /2...), các bản song ngữ tiếng Anh (- E, - EN), và các bảng tính, biểu mẫu phụ lục đính kèm. Trước khi sáng kiến được nghiên cứu và đưa vào ứng dụng, quy trình xử lý và hoàn thiện hồ sơ chứng thư gặp phải các bất cập, hạn chế nghiêm trọng sau:")
    
    p = add_styled_paragraph(doc)
    add_run(p, "a) Thao tác thủ công, lãng phí thời gian và năng suất lao động rất thấp: ", bold=True)
    add_run(p, "Quy trình chuyển đổi tài liệu Word và Excel sang định dạng PDF để ban hành điện tử hoàn toàn phụ thuộc vào việc thao tác bằng tay. Giám định viên (GDV) phải mở từng tệp tin Word, kiểm tra thủ công nội dung, bấm chọn File -> Save As/Export -> chọn định dạng PDF -> đặt tên file -> duyệt đến đúng thư mục lưu trữ -> lưu file. Đo lường thực tế cho thấy, mỗi thao tác xuất PDF thủ công như vậy tiêu tốn trung bình từ ")
    add_run(p, "30 đến 45 giây cho mỗi file", bold=True, color=RGBColor(180, 0, 0))
    add_run(p, ". Thông thường, mỗi đợt trả kết quả bao gồm khoảng ")
    add_run(p, "10 đến 20 file TBKQ", bold=True, color=RGBColor(0, 51, 102))
    add_run(p, " (Thông báo kết quả/chứng thư kèm các phân nhánh và bản tiếng Anh). Để xuất xong số file này, GDV phải mất liên tục từ ")
    add_run(p, "10 đến 15 phút", bold=True)
    add_run(p, " chỉ riêng cho việc ngồi nhấp chuột mở file, canh chỉnh và lưu PDF lặp đi lặp lại một cách cơ học. Khi cộng dồn nhiều đợt trả kết quả trong ngày và trong tuần, thời gian thao tác vô ích này lên đến hàng giờ mỗi ngày, trở thành “nút thắt cổ chai” kìm hãm tiến độ công việc và gây mệt mỏi lớn cho GDV.")
    
    p = add_styled_paragraph(doc)
    add_run(p, "b) Nguy cơ cao xuất sót file và gánh nặng công việc lớn cho Giám định viên: ", bold=True)
    add_run(p, "Trong mỗi đợt trả kết quả gồm khoảng ")
    add_run(p, "10 đến 20 file TBKQ", bold=True)
    add_run(p, " với các tiền tố, hậu tố phân nhánh tương tự nhau (ví dụ: ")
    add_run(p, "70454 - TBKQ - vn.docx, 70454 - TBKQ - en.docx, 63998 - TBKQ - 1 - vn.docx...", italic=True)
    add_run(p, "), việc nhấp chuột xuất thủ công từng file tạo áp lực căng thẳng và mỏi mắt rất lớn cho GDV. Thực tế đã ghi nhận không ít trường hợp GDV bị xuất sót 1 – 2 file (đặc biệt là sót các chứng thư phân nhánh /1, /2 hoặc sót bản tiếng Anh /E). Hậu quả là hồ sơ gửi cho khách hàng bị thiếu sót, phòng phải tốn thời gian xác minh, tiếp nhận phản hồi và xuất bổ sung, gây ảnh hưởng tiêu cực đến hình ảnh và uy tín chuyên nghiệp của Trung tâm 3.")
    
    p = add_styled_paragraph(doc)
    add_run(p, "c) Khó khăn trong việc chuyển đổi số và chèn phôi Giấy tiêu đề điện tử: ", bold=True)
    add_run(p, "Trước đây, kết quả giám định dạng bảng cứng được in trực tiếp từ máy in lên nền phôi giấy tiêu đề in sẵn của Trung tâm (pre-printed letterhead paper). Khi chuyển dịch sang phát hành chứng thư điện tử phục vụ khách hàng qua email, cổng thông tin trực tuyến và ký số văn bản, yêu cầu bắt buộc là trang đầu tiên của file PDF chứng thư phải được chèn hình nền giấy tiêu đề chuẩn (")
    add_run(p, "GIAY TIEU DE.pdf", bold=True)
    add_run(p, " cho bản tiếng Việt và ")
    add_run(p, "GIAY TIEU DE - E.pdf", bold=True)
    add_run(p, " cho bản tiếng Anh). Nếu thực hiện thủ công, GDV phải dùng các phần mềm chỉnh sửa PDF bên thứ ba bên ngoài rất phức tạp, dễ làm xô lệch hoặc giảm độ phân giải văn bản, tiềm ẩn rủi ro lộ lọt thông tin dữ liệu khách hàng, và rất dễ xảy ra nhầm lẫn tai hại giữa mẫu tiêu đề tiếng Việt và mẫu tiêu đề tiếng Anh.")
    
    p = add_styled_paragraph(doc)
    add_run(p, "d) Bất cập trong việc tìm kiếm hồ sơ và trỏ thủ công vào thư mục lưu trữ: ", bold=True)
    add_run(p, "Các hồ sơ kỹ thuật, quan trắc môi trường được lưu trữ theo cấu trúc thư mục nhiều tầng, phân cấp phức tạp trên máy tính cá nhân hoặc các ổ đĩa mạng chia sẻ (chia theo năm, theo loại hình dịch vụ, theo mã khách hàng...). Mỗi khi cần xuất kết quả hoặc xử lý chứng thư, Giám định viên phải mất từ ")
    add_run(p, "1 đến 2 phút", bold=True)
    add_run(p, " chỉ để duyệt tìm thủ công qua từng nhánh thư mục con để tìm đúng thư mục chứa tệp kết quả (")
    add_run(p, "tbkq", italic=True)
    add_run(p, "). Thao tác này vừa chậm chạp, tốn công sức, vừa dễ xảy ra nguy cơ chọn nhầm thư mục của hồ sơ khác có mã số gần giống nhau.")

    p = add_styled_paragraph(doc)
    add_run(p, "e) Vấn đề sai lệch ngày ban hành và không đồng bộ thời gian với hệ thống Odoo: ", bold=True)
    add_run(p, "Theo quy trình vận hành và kiểm soát chất lượng tại QUATEST 3, toàn bộ chứng thư giám định/thử nghiệm điện tử sau khi hoàn thiện sẽ được đăng tải (upload) lên hệ thống quản lý tác nghiệp ")
    add_run(p, "Odoo", bold=True, color=RGBColor(128, 0, 128))
    add_run(p, " để lưu trữ số và phát hành chính thức đến khách hàng. Yêu cầu bắt buộc là ")
    add_run(p, "ngày ban hành ghi trên chứng thư (Header) phải trùng khớp chính xác với ngày giờ thực tế mà phòng up hồ sơ lên hệ thống Odoo", bold=True)
    add_run(p, ". Trước đây khi cập nhật thủ công, GDV thường chỉ sửa được Header ở trang 1 mà quên cập nhật ở các trang tiếp theo; hoặc nếu dùng lệnh thay thế toàn văn bản (Ctrl+H) thì dễ bị sửa nhầm vào các trường ngày kỹ thuật trong thân bài (ngày lấy mẫu, ngày thử nghiệm). Đáng lo ngại hơn, do không kiểm soát chặt chẽ mốc giờ làm việc, hồ sơ hoàn thiện sau 16h00 chiều hoặc vào cuối tuần thường bị để nguyên ngày hiện tại thay vì cộng ngày làm việc tiếp theo để đồng bộ với thời điểm tải lên Odoo. Sự sai lệch này gây khó khăn lớn cho công tác đối soát, tra cứu của khách hàng, vi phạm quy trình kiểm soát chứng thư điện tử và ảnh hưởng trực tiếp đến kết quả đánh giá chất lượng nội bộ.")
    
    p = add_styled_paragraph(doc)
    add_run(p, "f) Rủi ro sai lệch các thông tin pháp lý trọng yếu khác (Số hồ sơ, Người ký, Số trang):", bold=True)
    add_run(p, "\n• ")
    add_run(p, "Sai lệch số hồ sơ (Document ID): ", bold=True)
    add_run(p, "Quy định số hiệu hồ sơ tại QUATEST 3 (định dạng chuẩn ")
    add_run(p, "Q3\\d{2}[A-D]\\d{6}", bold=True)
    add_run(p, ") yêu cầu quy chuẩn kiểm soát cực kỳ nghiêm ngặt: Mã số gốc phải có ở thân trang 1; đối với bản tiếng Anh, Header tất cả các trang bắt buộc phải có hậu tố ")
    add_run(p, "/E", bold=True)
    add_run(p, " hoặc ")
    add_run(p, "/E.R", bold=True)
    add_run(p, " và phải hoàn toàn đồng nhất giữa các trang; đối với bản tiếng Việt, Header tuyệt đối không được lẫn đuôi /E; các bản ban hành lại sau chỉnh sửa phải có hậu tố ")
    add_run(p, ".R", bold=True)
    add_run(p, ". Kiểm soát thủ công bằng mắt trên hàng trăm trang tài liệu thường xuyên bỏ sót các lỗi này.")
    add_run(p, "\n• ")
    add_run(p, "Sai lệch người ký (Giám định viên & Lãnh đạo phê duyệt): ", bold=True)
    add_run(p, "Trong cùng một bộ hồ sơ, thông tin người lập và người duyệt ký phải đảm bảo tính đồng nhất tuyệt đối. Đối với chứng thư tiếng Anh, tên người ký bắt buộc phải chuyển đổi sang định dạng không dấu theo quy chuẩn quốc tế. Việc kiểm tra thủ công rất khó phát hiện các trường hợp copy nhầm tên người ký từ hồ sơ cũ hoặc gõ sai lỗi chính tả.")
    add_run(p, "\n• ")
    add_run(p, "Sai lệch đánh số trang: ", bold=True)
    add_run(p, "Quy cách đánh số trang liên tục từ trang 2 trở đi (dạng ")
    add_run(p, "i/(total - 1)", italic=True)
    add_run(p, ") dễ bị lỗi khi thêm bớt bảng biểu hoặc chỉnh sửa layout văn bản Word.")
    
    p_sub2 = add_styled_paragraph(doc, space_before=8, space_after=4, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_run(p_sub2, "3.2. Ý tưởng và Nội dung giải pháp sáng kiến (Phần mềm UpChungThu):", bold=True, italic=True, font_size=13)
    
    p = add_styled_paragraph(doc)
    add_run(p, "Từ những bất cập thực tiễn nêu trên, tác giả đã chủ động nghiên cứu, thiết kế và phát triển phần mềm desktop chuyên dụng ")
    add_run(p, "UpChungThu", bold=True, color=RGBColor(0, 51, 102))
    add_run(p, " (ngôn ngữ Python, giao diện đồ họa Tkinter hiện đại hỗ trợ Dark/Light mode, ứng dụng công nghệ COM Automation tương tác sâu với Microsoft Office). Phần mềm giải quyết triệt để bài toán thông qua các trụ cột kỹ thuật đột phá sau:")

    add_figure(
        doc,
        "fig1_main_menu.png",
        "Hình 1: ",
        "Giao diện điều hướng chính của phần mềm UpChungThu và cấu hình đồng bộ dữ liệu mạng nội bộ (LAN Mode).",
        width_inches=6.0
    )

    p = add_styled_paragraph(doc)
    add_run(p, "1) Tính năng Mapping thư mục thông minh & Tìm kiếm nhanh số hồ sơ (Fast Record Autocomplete Mapping):", bold=True, color=RGBColor(0, 51, 102))
    add_run(p, "\nNhằm mục đích ")
    add_run(p, "nhập nhanh số hồ sơ và loại bỏ hoàn toàn việc phải tìm kiếm, trỏ vào hồ sơ thủ công", bold=True)
    add_run(p, ", phần mềm tích hợp module ")
    add_run(p, "MappingService", italic=True)
    add_run(p, ". Hệ thống tự động quét ngầm các cây thư mục gốc lưu trữ hồ sơ lớn (hàng nghìn thư mục hồ sơ) trên ổ cứng hoặc ổ đĩa mạng LAN, lập chỉ mục (indexing) toàn bộ các thư mục mang mã hồ sơ 6 chữ số (")
    add_run(p, "\\d{6}", italic=True)
    add_run(p, ") chỉ trong chưa đầy 1 giây. Khi làm việc, Giám định viên chỉ cần nhập nhanh số hồ sơ (ví dụ: ")
    add_run(p, "26A01234", bold=True)
    add_run(p, " hoặc 6 số) vào thanh tìm kiếm tự động hoàn thành (Autocomplete Search Bar), hệ thống sẽ lập tức gợi ý và khi chọn, phần mềm sẽ ")
    add_run(p, "tự động định vị và trỏ thẳng vào đúng thư mục con chứa kết quả (tbkq) trong thời gian chưa đầy 0.5 giây", bold=True, color=RGBColor(0, 128, 0))
    add_run(p, ". Tính năng này giúp tiết kiệm hoàn toàn từ 1 đến 2 phút thao tác tìm kiếm cho mỗi hồ sơ, ngăn ngừa triệt để nguy cơ trỏ nhầm thư mục hồ sơ của khách hàng khác.")

    add_figure(
        doc,
        "fig2_mapping_autocomplete.png",
        "Hình 2: ",
        "Tính năng Mapping tự động 768 hồ sơ trên ổ mạng Z: trong ~460 ms và tìm kiếm nhanh Autocomplete trỏ thẳng vào thư mục TBKQ.",
        width_inches=6.1
    )

    p = add_styled_paragraph(doc)
    add_run(p, "2) Tự động cập nhật chuẩn xác ngày ban hành trong Header – Đồng bộ chính xác với thời gian up hồ sơ lên hệ thống Odoo:", bold=True, color=RGBColor(128, 0, 128))
    add_run(p, "\nNhằm mục đích ")
    add_run(p, "cập nhật đúng thời gian thực tế phòng up hồ sơ lên hệ thống Odoo", bold=True)
    add_run(p, ", phần mềm áp dụng thuật toán ")
    add_run(p, "WordEngine.update_word_date()", italic=True)
    add_run(p, " giới hạn phạm vi tìm kiếm và thay thế nghiêm ngặt ")
    add_run(p, "chỉ trong Section Headers", bold=True)
    add_run(p, " (bao gồm cả các Shape, TextFrame bên trong Header). Nội dung thân tài liệu (body text, ngày lấy mẫu, ngày thử nghiệm) được bảo vệ nguyên vẹn tuyệt đối. Hệ thống tích hợp ")
    add_run(p, "Quy tắc cộng ngày thông minh theo mốc giờ làm việc tương ứng với chu kỳ up hồ sơ lên Odoo", bold=True)
    add_run(p, ":")
    add_run(p, "\n• ")
    add_run(p, "Xử lý trước 09:00 sáng: ", bold=True)
    add_run(p, "Cộng 0 ngày (+0 ngày, lấy ngày hiện tại) do hồ sơ được ký duyệt và tải lên hệ thống Odoo ngay trong buổi sáng.")
    add_run(p, "\n• ")
    add_run(p, "Xử lý từ 09:00 đến 16:00 chiều: ", bold=True)
    add_run(p, "Tự động cộng thêm 1 ngày (+1 ngày) để dự trù thời gian kiểm tra, ký số và up lên Odoo vào buổi sáng ngày làm việc tiếp theo.")
    add_run(p, "\n• ")
    add_run(p, "Xử lý sau 16:00 chiều: ", bold=True)
    add_run(p, "Tự động cộng thêm 2 ngày (+2 ngày) do hồ sơ sẽ chuyển sang chu kỳ làm việc và up Odoo của ngày kế tiếp.")
    add_run(p, "\n• ")
    add_run(p, "Tự động loại trừ Thứ 7 và Chủ Nhật: ", bold=True)
    add_run(p, "Nếu ngày tính toán rơi vào ngày nghỉ cuối tuần, hệ thống tự động nhảy sang ngày Thứ 2 đầu tuần làm việc tiếp theo.")
    add_run(p, "\nNhờ cơ chế này, ngày ban hành ghi trên chứng thư luôn trùng khớp chính xác 100% với thời điểm phòng hoàn tất thủ tục và tải chứng thư lên hệ thống Odoo, đảm bảo tính pháp lý, tính nhất quán dữ liệu số và loại bỏ hoàn toàn việc tính nhẩm lịch bằng tay.")

    add_figure(
        doc,
        "fig3_time_rules_odoo.png",
        "Hình 3: ",
        "Cấu hình Quy tắc cộng ngày tự động theo mốc giờ làm việc và Lịch chọn ngày nhằm đồng bộ chính xác với thời điểm tải hồ sơ lên hệ thống Odoo.",
        width_inches=6.1
    )

    p = add_styled_paragraph(doc)
    add_run(p, "3) Tự động hóa chuyển đổi hàng loạt (Batch Processing) siêu tốc từ Word/Excel sang PDF:", bold=True)
    add_run(p, "\nThay vì mở và xuất từng file bằng tay, phần mềm khởi tạo tiến trình ngầm ")
    add_run(p, "Word.Application", italic=True)
    add_run(p, " và ")
    add_run(p, "Excel.Application", italic=True)
    add_run(p, " thông qua giao thức COM (Visible=False, DisplayAlerts=False). Chỉ với một cú nhấp chuột, toàn bộ các file tài liệu trong thư mục được tự động duyệt, xử lý và xuất ra định dạng PDF chuẩn xác. Thời gian xuất trung bình giảm từ 30 - 45 giây/file xuống chỉ còn ")
    add_run(p, "2 đến 3 giây/file (tốc độ tăng gấp 10 - 15 lần)", bold=True, color=RGBColor(0, 128, 0))
    add_run(p, ". Mỗi đợt trả kết quả khoảng 10 – 20 file TBKQ được hoàn thành toàn bộ chỉ trong ")
    add_run(p, "chưa đầy 1 phút", bold=True)
    add_run(p, " (thay vì 10 – 15 phút làm tay). Cơ chế duyệt cây thư mục toàn diện giúp quét sạch 100% tệp tin, ")
    add_run(p, "loại bỏ hoàn toàn nguy cơ xuất sót file hồ sơ", bold=True)
    add_run(p, ", giải phóng hoàn toàn gánh nặng công việc cơ học cho Giám định viên.")

    add_figure(
        doc,
        "fig4_word_batch_success.png",
        "Hình 4: ",
        "Kết quả xử lý hàng loạt (Batch Processing) bộ 12 file chứng thư Word song ngữ Việt – Anh của hồ sơ Q326D012000 hoàn tất tự động và đạt chuẩn 100% (OK).",
        width_inches=6.1
    )

    p = add_styled_paragraph(doc)
    add_run(p, "4) Tự động nhận diện ngôn ngữ và chèn phôi Giấy tiêu đề điện tử (Smart Letterhead Overlay):", bold=True)
    add_run(p, "\nPhần mềm tích hợp module thuật toán ")
    add_run(p, "detect_language()", italic=True)
    add_run(p, " tự động quét tên file và tên sheet (phát hiện các tiền tố/hậu tố '- E', '- EN', '- ENG') để phân loại chính xác tài liệu tiếng Việt hay tiếng Anh. Ngay sau khi xuất PDF thô, module ")
    add_run(p, "PdfService", italic=True)
    add_run(p, " lập tức nạp phôi tiêu đề tương ứng (")
    add_run(p, "GIAY TIEU DE.pdf", bold=True)
    add_run(p, " hoặc ")
    add_run(p, "GIAY TIEU DE - E.pdf", bold=True)
    add_run(p, ") và thực hiện kỹ thuật ghép lớp vector (overlay merge) chính xác lên trang 1 của tài liệu. Quá trình ghép nền diễn ra trong vài mili-giây, bảo toàn 100% định dạng, phông chữ và bố cục tài liệu, tối ưu hóa triệt để khâu hoàn thiện chứng thư số thay cho việc in phôi giấy cứng truyền thống.")

    add_figure(
        doc,
        "fig5_letterhead_overlay.png",
        "Hình 5: ",
        "Đối chứng trực quan trang 1 chứng thư PDF (hồ sơ Q326D095083) trước và sau khi tự động nhận diện ngôn ngữ (VN/EN) và ghép phôi Giấy tiêu đề điện tử.",
        width_inches=6.4
    )

    p = add_styled_paragraph(doc)
    add_run(p, "5) Hệ thống Kiểm soát toàn vẹn hồ sơ đa tầng (Post-Export Multi-Check Pipeline):", bold=True)
    add_run(p, "\nSau khi xuất file PDF và chèn giấy tiêu đề, phần mềm tự động kích hoạt bộ quét trích xuất dữ liệu đa luồng chạy ngầm kiểm tra đồng thời 4 tiêu chí pháp lý khắt khe:")
    add_run(p, "\n• ")
    add_run(p, "Kiểm soát số hồ sơ (Document ID): ", bold=True)
    add_run(p, "Quét toàn bộ các trang PDF theo Regex ")
    add_run(p, "Q3\\d{2}[A-D]\\d{6}(?:/\\d+)?(?:/E)?(?:\\.R)?", italic=True)
    add_run(p, "; kiểm tra tính thống nhất của mã số gốc; đối soát nghiêm ngặt bản tiếng Anh (header mọi trang phải có đuôi /E, thân trang 1 phải là mã gốc không có /E) và bản tiếng Việt (header không có /E); kiểm soát mã sửa đổi .R. Báo cáo chi tiết từng vị trí vi phạm (ví dụ: ")
    add_run(p, "P1(thiếu /E header), P1(thiếu số gốc body)...", italic=True)
    add_run(p, ").")
    add_run(p, "\n• ")
    add_run(p, "Kiểm soát ngày ban hành: ", bold=True)
    add_run(p, "Kiểm tra số lần xuất hiện của ngày ban hành trên từng trang; áp dụng quy tắc chuẩn: mỗi trang phải có đúng 1 ngày. Cảnh báo chính xác trang bị thừa hoặc thiếu ngày.")
    add_run(p, "\n• ")
    add_run(p, "Kiểm soát người ký (Thư viện Người ký GDV & Boss): ", bold=True)
    add_run(p, "Quản lý danh sách GDV và Lãnh đạo; tích hợp thuật toán chuẩn hóa chuỗi và hàm bỏ dấu ")
    add_run(p, "remove_accents()", italic=True)
    add_run(p, " để nhận diện chính xác cả tên có dấu tiếng Việt và tên không dấu tiếng Anh. Đặc biệt, cơ chế ")
    add_run(p, "Consistency Check", bold=True)
    add_run(p, " tự động lấy bộ người ký của file đầu tiên làm mẫu chuẩn cho toàn bộ thư mục, phát hiện ngay lập tức các file có người ký khác mẫu (cảnh báo ")
    add_run(p, "KHÁC MẪU", bold=True, color=RGBColor(180, 0, 0))
    add_run(p, ").")
    add_run(p, "\n• ")
    add_run(p, "Kiểm soát đánh số trang: ", bold=True)
    add_run(p, "Tự động kiểm tra chuỗi số trang liên tục ")
    add_run(p, "i/(total - 1)", italic=True)
    add_run(p, " từ trang 2 trở đi, thông báo rõ ràng trang nào bị lỗi định dạng hoặc nhảy số trang.")

    add_figure(
        doc,
        "fig6_qa_qc_multicheck.png",
        "Hình 6: ",
        "Hệ thống Kiểm soát toàn vẹn 4 tầng tự động bôi đỏ cảnh báo chi tiết các lỗi vi phạm pháp lý (Số hồ sơ Q3, Ngày, Chữ ký KHÁC MẪU, Trang) và cửa sổ Quản lý Thư viện Người ký.",
        width_inches=6.4
    )

    p = add_styled_paragraph(doc)
    add_run(p, "6) Các tiện ích mở rộng phục vụ vận hành hàng ngày và đồng bộ dữ liệu mạng LAN:", bold=True)
    add_run(p, "\n• ")
    add_run(p, "Phân hệ Excel chuyên sâu: ", bold=True)
    add_run(p, "Tự động quét cấu trúc sheet của toàn bộ file Excel trong thư mục; cơ chế chọn nhanh (Bulk Select) và đánh dấu gộp sheet ")
    add_run(p, "[ + ]", bold=True)
    add_run(p, " để ghép nối nhiều sheet thành 1 tệp PDF duy nhất; hỗ trợ 2 chế độ: Xuất tính toán (nội bộ) và Xuất chứng thư (chèn nền và kiểm tra lỗi).")
    add_run(p, "\n• ")
    add_run(p, "Đồng bộ cấu hình qua mạng nội bộ LAN: ", bold=True)
    add_run(p, "Cho phép lưu trữ file cấu hình dùng chung trên ổ mạng chia sẻ (LAN Mode), giúp toàn bộ các máy tính của GDV trong phòng tự động đồng bộ danh sách người ký, mốc giờ và thư mục mapping mà không cần cài đặt lặp lại.")

    add_figure(
        doc,
        "fig7_excel_processor.png",
        "Hình 7: ",
        "Giao diện Phân hệ Xử lý File Excel tự động quét cấu trúc Sheet (hồ sơ Q326D008933), hộp thoại Chọn nhanh & Gộp Sheet (+) và xuất chứng thư đạt chuẩn.",
        width_inches=6.4
    )
    
    p_sub3 = add_styled_paragraph(doc, space_before=8, space_after=4, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_run(p_sub3, "3.3. Tên văn bản ban hành hoặc tên đề tài, đề án, sản phẩm KH&CN:", bold=True, italic=True, font_size=13)
    p = add_styled_paragraph(doc)
    add_run(p, "• Tên sản phẩm phần mềm: ")
    add_run(p, "Phần mềm tự động hóa xử lý văn bản, xuất chứng thư và kiểm tra tính toàn vẹn hồ sơ giám định (UpChungThu - Phiên bản v6.0).", bold=True)
    add_run(p, "\n• Bộ tài liệu kỹ thuật và quản lý chất lượng ban hành kèm theo:")
    add_run(p, "\n  + Sổ tay hướng dẫn sử dụng chi tiết (USER_GUIDE.md - Hệ thống xử lý văn bản QUATEST 3).")
    add_run(p, "\n  + Tài liệu Kiến trúc Hệ thống và Quy chuẩn kỹ thuật phần mềm (ARCHITECTURE.md).")
    
    # 4. Thời gian thực hiện
    p_time_title = add_styled_paragraph(doc, space_before=10, space_after=4, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_run(p_time_title, "Thời gian thực hiện:", bold=True, font_size=13)
    p = add_styled_paragraph(doc)
    add_run(p, "• Thời gian bắt đầu nghiên cứu, khảo sát và xây dựng kiến trúc: ")
    add_run(p, "Tháng 01/2026.", bold=True)
    add_run(p, "\n• Thời gian lập trình thử nghiệm, hoàn thiện các phân hệ Word, Excel, PDF, Mapping và Odoo Sync: ")
    add_run(p, "Tháng 02/2026 – Tháng 03/2026.", bold=True)
    add_run(p, "\n• Thời gian đưa vào ứng dụng chính thức tại Nhóm Quan trắc Môi trường – Phòng Nghiệp vụ 2: ")
    add_run(p, "Từ tháng 04/2026 đến nay.", bold=True, color=RGBColor(0, 102, 204))
    
    # 5. Kết quả
    p_res_title = add_styled_paragraph(doc, space_before=10, space_after=4, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_run(p_res_title, "Kết quả đạt được:", bold=True, font_size=13)
    p = add_styled_paragraph(doc)
    add_run(p, "a) Sản phẩm cụ thể: ", bold=True)
    add_run(p, "Đã hoàn thành xây dựng và đóng gói bộ phần mềm ")
    add_run(p, "UpChungThu v6.0", bold=True)
    add_run(p, " với đầy đủ các phân hệ xử lý Word, Excel, PDF, tính năng Mapping thư mục tự động, cơ chế đồng bộ ngày theo chu kỳ up Odoo, đồng bộ mạng LAN và thư viện người ký. Phần mềm được đóng gói thành tệp thực thi (.exe) độc lập, chạy trực tiếp trên hệ điều hành Windows mà không đòi hỏi người dùng phải cài đặt môi trường lập trình Python hay các thư viện bổ trợ.")
    
    p = add_styled_paragraph(doc)
    add_run(p, "b) Bảng số liệu định lượng đối chứng hiệu quả trước và sau khi áp dụng sáng kiến:", bold=True)
    
    # Table comparison (8 rows)
    table_comp = doc.add_table(rows=9, cols=4)
    table_comp.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_comp.autofit = False
    set_table_borders(table_comp)
    
    widths = [Inches(1.8), Inches(1.8), Inches(1.8), Inches(1.3)]
    headers = ["Tiêu chí đánh giá", "Trước khi áp dụng (Thủ công)", "Sau khi áp dụng (Phần mềm UpChungThu)", "Hiệu quả cải thiện"]
    
    for c_idx, h_text in enumerate(headers):
        cell = table_comp.cell(0, c_idx)
        cell.width = widths[c_idx]
        set_cell_margins(cell, 120, 120, 100, 100)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_run(p, h_text, bold=True, font_size=10.5)
        
    data_rows = [
        ("Tìm kiếm & trỏ vào thư mục hồ sơ", "Duyệt thủ công qua nhiều tầng thư mục mất 1 – 2 phút, dễ chọn nhầm", "Gõ nhanh số hồ sơ qua Mapping Autocomplete, tự động trỏ vào 'tbkq' < 0.5s", "Nhanh tức thì, chính xác 100%, không nhầm lẫn"),
        ("Thời gian xuất PDF cho 1 file", "30 – 45 giây / file (mở, kiểm tra, chọn Save As, đặt tên)", "2 – 3 giây / file (xử lý ngầm tự động hoàn toàn)", "Nhanh hơn gấp 10 – 15 lần"),
        ("Thời gian xuất 1 đợt kết quả (10 – 20 file TBKQ)", "10 – 15 phút thao tác liên tục bằng tay", "Khoảng 30 giây – 1 phút (chỉ với 1 cú nhấp chuột)", "Tiết kiệm ~90 – 95% thời gian"),
        ("Rủi ro xuất sót file hồ sơ", "3% – 5% (do mỏi mắt, thao tác nhầm với tên file tương tự)", "0% (thuật toán quét và xử lý duyệt toàn bộ thư mục)", "Triệt tiêu hoàn toàn rủi ro sót file"),
        ("Chèn phôi giấy tiêu đề điện tử", "Thủ công ngoài, mất 1-2 phút/file, dễ nhầm mẫu VN/EN", "Tự động nhận diện ngôn ngữ và ghép nền trong vài ms", "Chính xác 100%, tức thì, không xô lệch"),
        ("Cập nhật ngày ban hành & đồng bộ hệ thống Odoo", "Sửa tay từng trang, dễ sửa nhầm thân bài; lệch ngày với thời điểm upload Odoo", "Tự động cập nhật Header theo mốc giờ, loại trừ T7/CN; đồng bộ 100% với giờ up Odoo", "Chính xác tuyệt đối, chuẩn hóa chu kỳ phát hành Odoo"),
        ("Kiểm soát sai lệch (ID, Ký, Trang)", "Soát bằng mắt thường, dễ sót lỗi trên hồ sơ nhiều trang", "Hệ thống kiểm soát 4 tầng tự động bôi đỏ cảnh báo lỗi", "Kiểm soát toàn vẹn 100% hồ sơ"),
        ("Gánh nặng cho Giám định viên", "Căng thẳng, áp lực thao tác cơ học lặp đi lặp lại", "Thư thái, an tâm, tập trung vào chuyên môn kỹ thuật", "Giải phóng hoàn toàn sức lao động")
    ]
    
    for r_idx, row_data in enumerate(data_rows, 1):
        for c_idx, cell_value in enumerate(row_data):
            cell = table_comp.cell(r_idx, c_idx)
            cell.width = widths[c_idx]
            set_cell_margins(cell, 90, 90, 90, 90)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx == 3 else (WD_ALIGN_PARAGRAPH.LEFT if c_idx == 0 else WD_ALIGN_PARAGRAPH.JUSTIFY)
            bold_flag = True if c_idx in [0, 3] else False
            color = RGBColor(0, 128, 0) if c_idx == 3 else (RGBColor(180, 0, 0) if c_idx == 1 and r_idx in [1, 2, 3, 4] else None)
            add_run(p, cell_value, bold=bold_flag, font_size=9.5, color=color)
            
    # 6. Hiệu quả áp dụng
    p_eff_title = add_styled_paragraph(doc, space_before=12, space_after=4, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_run(p_eff_title, "Hiệu quả áp dụng:", bold=True, font_size=13)
    
    p = add_styled_paragraph(doc)
    add_run(p, "1) Hiệu quả kinh tế & Năng suất lao động: ", bold=True)
    add_run(p, "Nhờ tích hợp tính năng Mapping thư mục tìm kiếm tức thì và tự động hóa toàn diện quá trình xuất chứng thư, thời gian xử lý mỗi đợt hồ sơ được rút ngắn tối đa (từ 10 – 15 phút thao tác tay xuống chỉ còn 30 giây – 1 phút). Ước tính với lưu lượng phát hành từ ")
    add_run(p, "2.000 đến 3.000 chứng thư/năm", bold=True)
    add_run(p, " tại nhóm, phần mềm giúp tiết kiệm từ ")
    add_run(p, "250 đến 350 giờ làm việc có chuyên môn", bold=True, color=RGBColor(0, 102, 204))
    add_run(p, " của các Giám định viên mỗi năm. Rút ngắn đáng kể thời gian trả kết quả cho khách hàng, nâng cao năng suất xử lý công việc hàng ngày của nhóm.")
    
    p = add_styled_paragraph(doc)
    add_run(p, "2) Hiệu quả quản lý chất lượng & Đồng bộ quy trình hệ thống Odoo: ", bold=True)
    add_run(p, "Chuẩn hóa 100% hình thức chứng thư điện tử phát hành, loại bỏ hoàn toàn các lỗi sai sót hành chính (sai ngày, lệch ngày up Odoo, sai mã Q3, sai người ký, sót file...). Đảm bảo tính nhất quán tuyệt đối giữa dữ liệu bản mềm PDF và dữ liệu tác nghiệp trên hệ thống ")
    add_run(p, "Odoo", bold=True, color=RGBColor(128, 0, 128))
    add_run(p, ", đáp ứng nghiêm ngặt các tiêu chuẩn quản lý chất lượng phòng thử nghiệm và tổ chức giám định theo ")
    add_run(p, "ISO/IEC 17025", bold=True)
    add_run(p, " và ")
    add_run(p, "ISO 9001", bold=True)
    add_run(p, ", nâng cao sự hài lòng và tin cậy của khách hàng trong và ngoài nước đối với thương hiệu QUATEST 3.")
    
    p = add_styled_paragraph(doc)
    add_run(p, "3) Cải thiện điều kiện làm việc và giải phóng sức lao động: ", bold=True)
    add_run(p, "Giải phóng Giám định viên khỏi các chuỗi thao tác chuột lặp đi lặp lại vô cùng mệt mỏi, giảm thiểu stress và chứng căng thẳng thị giác do nhìn màn hình liên tục. Giúp đội ngũ cán bộ kỹ thuật an tâm, tập trung trí tuệ vào công tác chuyên môn sâu như phân tích dữ liệu quan trắc, đánh giá kỹ thuật và nghiên cứu phương pháp thử nghiệm mới.")
    
    # 7. Phạm vi ảnh hưởng
    p_scope_title = add_styled_paragraph(doc, space_before=10, space_after=4, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_run(p_scope_title, "Phạm vi ảnh hưởng (cấp cơ sở/ cấp Bộ,…):", bold=True, font_size=13)
    p = add_styled_paragraph(doc)
    add_run(p, "• Cấp cơ sở: ", bold=True)
    add_run(p, "Sáng kiến được triển khai và áp dụng thực tế phục vụ trực tiếp công tác xử lý hồ sơ, phát hành chứng thư hàng ngày tại ")
    add_run(p, "Nhóm Quan trắc Môi trường – Phòng Nghiệp vụ 2", bold=True)
    add_run(p, " thuộc Trung tâm Kỹ thuật Tiêu chuẩn Đo lường Chất lượng 3 (QUATEST 3).")
    
    # Signatures Table
    doc.add_paragraph().paragraph_format.space_after = Pt(8)
    t_sig = doc.add_table(rows=2, cols=2)
    t_sig.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_sig.autofit = False
    for r in t_sig.rows:
        for idx, w in enumerate([Inches(3.3), Inches(3.2)]):
            r.cells[idx].width = w
            
    c00 = t_sig.cell(0, 0)
    p00 = c00.paragraphs[0]
    p00.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p00.paragraph_format.line_spacing = 1.15
    add_run(p00, "XÁC NHẬN CỦA GIÁM ĐỐC\n", bold=True, font_size=12)
    add_run(p00, "(Ký, ghi rõ họ tên, đóng dấu)", italic=True, font_size=11)
    
    c01 = t_sig.cell(0, 1)
    p01 = c01.paragraphs[0]
    p01.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p01.paragraph_format.line_spacing = 1.15
    add_run(p01, "NGƯỜI BÁO CÁO\n", bold=True, font_size=12)
    add_run(p01, "(Ký, ghi rõ họ tên)", italic=True, font_size=11)
    
    c10 = t_sig.cell(1, 0)
    p10 = c10.paragraphs[0]
    p10.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p10.paragraph_format.space_before = Pt(60)
    add_run(p10, "", bold=True, font_size=12)
    
    c11 = t_sig.cell(1, 1)
    p11 = c11.paragraphs[0]
    p11.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p11.paragraph_format.space_before = Pt(60)
    add_run(p11, "Nguyễn Tiến Phát", bold=True, font_size=12)
    
    out_docx = os.path.join(OUT_DIR, "BM23.01.01 Báo cáo sáng kiến - Phần mềm UpChungThu - Nguyễn Tiến Phát.docx")
    doc.save(out_docx)
    print("Xây dựng thành công:", out_docx)
    return out_docx

def build_excel_bm23_01_02():
    print("--- Đang cập nhật tệp Excel BM23.01.02 ---")
    ref_file = os.path.join(REF_DIR, "BM23.01.02.xls")
    out_file = os.path.join(OUT_DIR, "BM23.01.02 Bảng tổng hợp sáng kiến - UpChungThu.xls")
    
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(os.path.abspath(ref_file))
        ws = wb.Sheets("Bang tom tat sang kien")
        
        # Cập nhật tiêu đề năm 2026
        ws.Cells(5, 1).Value = "BẢNG TỔNG HỢP SÁNG KIẾN CỦA CÁC CÁ NHÂN THUỘC TRUNG TÂM KỸ THUẬT 3 NĂM 2026"
        
        # Cập nhật thông tin hàng 9 (Row 9 in Excel, tương ứng TT số 1)
        ws.Cells(9, 1).Value = 1
        ws.Cells(9, 2).Value = "953N224"
        ws.Cells(9, 3).Value = "Nguyễn Tiến Phát"
        ws.Cells(9, 4).Value = "Phòng NV2"
        ws.Cells(9, 5).Value = "Giám định viên"
        ws.Cells(9, 6).Value = "“Xây dựng và ứng dụng phần mềm tự động hóa xử lý văn bản, xuất chứng thư điện tử và kiểm soát tính toàn vẹn hồ sơ giám định (UpChungThu)”"
        
        hien_trang = (
            "Hiện trạng: Thao tác xuất PDF và chuyển đổi chứng thư thực hiện thủ công từng file, tốn trung bình 30 - 45 giây/file. "
            "Mỗi đợt trả kết quả gồm khoảng 10 - 20 file TBKQ phân nhánh, việc xuất thủ công từng file tốn 10 - 15 phút, gây mỏi mắt và nguy cơ xuất sót file lớn. "
            "GDV phải duyệt tìm thủ công qua các cây thư mục nhiều tầng mất 1 - 2 phút/hồ sơ, dễ chọn nhầm thư mục. "
            "Việc chuyển dịch từ in phôi giấy tiêu đề cứng sang chèn giấy tiêu đề điện tử phải làm thủ công ngoài, dễ nhầm lẫn giữa bản VN và EN. "
            "Thường xuyên xảy ra sai lệch ngày ban hành so với thời điểm phòng up hồ sơ lên hệ thống Odoo (do sửa tay, nhầm mốc giờ, sót trang Header); "
            "đồng thời tiềm ẩn nguy cơ sai lệch mã số hồ sơ Q3, sai lệch người ký (GDV/Boss) và lỗi đánh số trang."
        )
        ws.Cells(9, 7).Value = hien_trang
        
        y_tuong = (
            "Ý tưởng đối với giải pháp: Xây dựng phần mềm desktop chuyên dụng UpChungThu (Python/Tkinter, COM Automation) tự động hóa toàn diện: "
            "(1) Tính năng Mapping thư mục thông minh, nhập nhanh số hồ sơ tự động trỏ vào thư mục 'tbkq' < 0.5s, không cần tìm thủ công; "
            "(2) Tự động cập nhật ngày ban hành trong Header theo mốc giờ làm việc, loại trừ T7/CN, đồng bộ chuẩn xác 100% với thời gian phòng up hồ sơ lên hệ thống Odoo; "
            "(3) Xuất hàng loạt Word/Excel sang PDF siêu tốc (2-3s/file), mỗi đợt 10 - 20 file TBKQ chỉ mất 30s - 1 phút thay vì 10 - 15 phút, triệt tiêu rủi ro sót file; "
            "(4) Tự động nhận diện ngôn ngữ và ghép phôi giấy tiêu đề điện tử lên trang 1; "
            "(5) Hệ thống kiểm soát tính toàn vẹn 4 tầng (Mã số Q3, Ngày ban hành, Chữ ký GDV & Boss hỗ trợ tên không dấu tiếng Anh, Đánh số trang). "
            "Vai trò: Tác giả ý tưởng, chủ trì nghiên cứu và xây dựng giải pháp (100%)."
        )
        ws.Cells(9, 8).Value = y_tuong
        
        ket_qua = (
            "Kết quả: Hoàn thành phần mềm UpChungThu v6.0 đóng gói chạy độc lập trên Windows, kèm bộ Sổ tay HDSD (USER_GUIDE.md) và Kiến trúc hệ thống (ARCHITECTURE.md). "
            "Hiệu quả: Rút ngắn thời gian định vị thư mục và xuất file gấp 10 - 15 lần (mỗi đợt 10 - 20 file chỉ tốn 30s - 1 phút thay vì 10 - 15 phút); với lưu lượng 2.000 - 3.000 chứng thư/năm tại nhóm, giúp tiết kiệm 250 - 350 giờ làm việc/năm cho GDV; "
            "loại bỏ 100% rủi ro sót file; đồng bộ chuẩn xác ngày ban hành với hệ thống Odoo; kiểm soát toàn diện không còn sai lệch ngày, mã Q3, người ký, số trang; chuẩn hóa chứng thư số theo ISO/IEC 17025."
        )
        ws.Cells(9, 9).Value = ket_qua
        
        ws.Cells(9, 10).Value = "Cấp cơ sở (Áp dụng trực tiếp tại Nhóm Quan trắc Môi trường - Phòng Nghiệp vụ 2, Trung tâm Kỹ thuật 3)."
        
        wb.SaveAs(os.path.abspath(out_file), FileFormat=56)
        wb.Close(False)
        print("Xây dựng thành công:", out_file)
    finally:
        excel.Quit()
    return out_file

def build_excel_mau_03():
    print("--- Đang cập nhật tệp Excel Mau 03-Phu luc I ---")
    ref_file = os.path.join(REF_DIR, "Mau 03-Phu luc I (tat ca SK).xls")
    out_file = os.path.join(OUT_DIR, "Mau 03 - Phu luc I - UpChungThu.xls")
    
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(os.path.abspath(ref_file))
        ws = wb.Sheets("Bang tom tat sang kien")
        
        # Cập nhật tiêu đề năm 2026
        ws.Cells(3, 1).Value = "DANH MỤC SÁNG KIẾN ĐỀ NGHỊ CÔNG NHẬN HIỆU QUẢ ÁP DỤNG VÀ PHẠM VI ẢNH HƯỞNG CẤP CƠ SỞ NĂM 2026"
        
        # Cập nhật hàng 7 (Row 7 in Excel, STT số 1)
        ws.Cells(7, 1).Value = 1
        ws.Cells(7, 2).Value = "953N224"
        ws.Cells(7, 3).Value = "Nguyễn Tiến Phát"
        ws.Cells(7, 4).Value = "Phòng NV2"
        ws.Cells(7, 5).Value = "“Xây dựng và ứng dụng phần mềm tự động hóa xử lý văn bản, xuất chứng thư điện tử và kiểm soát tính toàn vẹn hồ sơ giám định (UpChungThu)”"
        
        noi_dung_kq_hq = (
            "Nội dung: Xây dựng phần mềm desktop chuyên dụng UpChungThu ứng dụng COM Automation điều khiển MS Word và Excel ngầm định để tự động hóa toàn bộ quy trình xuất chứng thư điện tử. "
            "Điểm mới cốt lõi bao gồm: "
            "(1) Tính năng Mapping thư mục thông minh, cho phép nhập nhanh số hồ sơ qua Autocomplete và tự động định vị, trỏ thẳng vào thư mục kết quả 'tbkq' trong < 0.5 giây, xóa bỏ hoàn toàn thao tác tìm kiếm duyệt thư mục thủ công; "
            "(2) Tự động cập nhật chuẩn xác ngày ban hành trong Header theo mốc giờ làm việc và loại trừ T7, CN, đảm bảo khớp đúng thời gian thực tế phòng up hồ sơ lên hệ thống Odoo mà không ảnh hưởng nội dung thân bài; "
            "(3) Chuyển đổi hàng loạt Word/Excel sang PDF với tốc độ siêu tốc (2-3 giây/file thay vì 30-45 giây/file thủ công), mỗi đợt khoảng 10 - 20 file TBKQ chỉ mất 30s - 1 phút thay vì 10 - 15 phút, triệt tiêu hoàn toàn nguy cơ xuất sót file hồ sơ; "
            "(4) Tự động nhận diện ngôn ngữ (VN/EN) và phủ ghép phôi giấy tiêu đề điện tử lên trang 1 của chứng thư, giải quyết bài toán chuyển dịch từ phôi in giấy cứng sang chứng thư số; "
            "(5) Tích hợp hệ thống kiểm soát toàn vẹn đa tầng tự động phát hiện mọi sai lệch về mã số hồ sơ Q3 (đặc biệt phân biệt chặt chẽ bản VN và bản EN), sai lệch ngày ban hành, sai lệch người ký (GDV và Boss, hỗ trợ tên không dấu tiếng Anh) và sai lệch đánh số trang.\n\n"
            "Kết quả: Hoàn thành xây dựng và đóng gói phần mềm UpChungThu phiên bản v6.0 chạy độc lập trên Windows (không cần cài đặt môi trường lập trình); "
            "ban hành bộ Sổ tay Hướng dẫn sử dụng chi tiết (USER_GUIDE.md) và Tài liệu Kiến trúc Hệ thống (ARCHITECTURE.md). Phần mềm đã được đưa vào vận hành thực tế tại nhóm từ tháng 04/2026 đến nay.\n\n"
            "Hiệu quả: Nâng cao năng suất lao động vượt trội, giảm thời gian xử lý hồ sơ gấp 10-15 lần (mỗi đợt 10 - 20 file TBKQ xuất xong trong chưa đầy 1 phút); với lưu lượng 2.000 - 3.000 chứng thư/năm tại nhóm, tiết kiệm ước tính 250 - 350 giờ làm việc mỗi năm cho Giám định viên; "
            "loại bỏ triệt để rủi ro xuất sót file, sai lệch ngày với hệ thống Odoo và các sai lệch thông tin pháp lý trên chứng thư; giải phóng sức lao động cơ học, giảm áp lực mệt mỏi cho Giám định viên; "
            "chuẩn hóa 100% hình thức chứng thư điện tử phục vụ chuyển đổi số và tuân thủ tiêu chuẩn ISO/IEC 17025."
        )
        ws.Cells(7, 6).Value = noi_dung_kq_hq
        
        pham_vi = (
            "Đơn vị áp dụng: Nhóm Quan trắc Môi trường – Phòng Nghiệp vụ 2 thuộc Trung tâm Kỹ thuật Tiêu chuẩn Đo lường Chất lượng 3 (QUATEST 3).\n\n"
            "Lĩnh vực áp dụng: Xử lý văn bản, tự động hóa xuất chứng thư điện tử và kiểm soát chất lượng hồ sơ giám định/thử nghiệm của nhóm."
        )
        ws.Cells(7, 7).Value = pham_vi
        
        wb.SaveAs(os.path.abspath(out_file), FileFormat=56)
        wb.Close(False)
        print("Xây dựng thành công:", out_file)
    finally:
        excel.Quit()
    return out_file

if __name__ == "__main__":
    print("Bắt đầu khởi tạo các báo cáo sáng kiến năm 2026...")
    f1 = build_docx_report()
    f2 = build_excel_bm23_01_02()
    f3 = build_excel_mau_03()
    print("HOÀN TẤT TOÀN BỘ CÁC TỆP BÁO CÁO!")
