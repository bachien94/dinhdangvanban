import streamlit as st
from docx import Document
from docx.shared import Cm, RGBColor, Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io

# Cấu hình trang
st.set_page_config(
    page_title="Công Cụ Xử Lý Tệp Word",
    page_icon="📄",
    layout="wide"
)

st.title("📄 Công Cụ Xử Lý Tệp Word (.docx)")
st.markdown("Tải tệp lên ở bên trái và tùy chỉnh nhanh ở bên phải.")

# Chia 2 cột ngang
col_left, col_right = st.columns(2, gap="large")

with col_left:
    st.subheader("📂 1. Tải tệp lên")
    uploaded_file = st.file_uploader("Chọn tệp Word (.docx)", type=["docx"])
    
    doc = None
    if uploaded_file is not None:
        try:
            doc = Document(uploaded_file)
            st.success(f"Đã tải lên thành công: **{uploaded_file.name}**")
        except Exception as e:
            st.error(f"Lỗi đọc tệp: {e}")
            
    st.markdown("---")
    st.subheader("📥 3. Tải xuống kết quả")
    download_placeholder = st.empty()

with col_right:
    st.subheader("🛠️ 2. Tùy chỉnh xử lý")
    
    is_enabled = (uploaded_file is not None)
    
    # 1. Tùy chỉnh lề trang
    custom_margin = st.checkbox("Tùy chỉnh lề trang (cm)", disabled=not is_enabled)
    top_margin, bottom_margin, left_margin, right_margin = 2.0, 2.0, 3.0, 2.0
    if custom_margin:
        m_col1, m_col2 = st.columns(2)
        with m_col1:
            top_margin = st.number_input("Lề trên", min_value=0.5, max_value=10.0, value=2.0, step=0.1)
            bottom_margin = st.number_input("Lề dưới", min_value=0.5, max_value=10.0, value=2.0, step=0.1)
        with m_col2:
            left_margin = st.number_input("Lề trái", min_value=0.5, max_value=10.0, value=3.0, step=0.1)
            right_margin = st.number_input("Lề phải", min_value=0.5, max_value=10.0, value=2.0, step=0.1)
    
    # 2. Loại bỏ màu nền và đưa chữ về đen
    remove_background = st.checkbox("Loại bỏ màu nền và đưa chữ về màu đen", disabled=not is_enabled)

    # 3. Định dạng Văn bản & Đoạn văn
    standardize_typography = st.checkbox("Định dạng Văn bản & Đoạn văn (Font, Giãn dòng)", disabled=not is_enabled)
    
    font_name, font_size = "Times New Roman", 13
    line_spacing, space_after = 1.15, 6.0
    
    if standardize_typography:
        t_col1, t_col2 = st.columns(2)
        with t_col1:
            font_name = st.selectbox("Font chữ", ["Times New Roman", "Arial", "Calibri", "Cambria"], index=0)
            font_size = st.number_input("Cỡ chữ (pt)", min_value=8, max_value=72, value=13, step=1)
        with t_col2:
            line_spacing = st.selectbox("Giãn dòng", [1.0, 1.15, 1.5, 2.0], index=1)
            space_after = st.number_input("Khoảng cách sau đoạn (pt)", min_value=0.0, max_value=30.0, value=6.0, step=1.0)

    # 4. Tùy chọn Autofit chuẩn của Word
    autofit_tables = st.checkbox("Tự động co giãn bảng (Autofit chuẩn Word)", value=True, disabled=not is_enabled)

# Xử lý tệp khi người dùng đã upload
if uploaded_file is not None and doc is not None:
    if custom_margin:
        for section in doc.sections:
            section.top_margin = Cm(top_margin)
            section.bottom_margin = Cm(bottom_margin)
            section.left_margin = Cm(left_margin)
            section.right_margin = Cm(right_margin)
    
    if remove_background:
        for paragraph in doc.paragraphs:
            for run in paragraph.runs:
                run.font.color.rgb = RGBColor(0, 0, 0)
                rPr = run._r.get_or_add_rPr()
                highlight = rPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}highlight')
                if highlight is not None:
                    rPr.remove(highlight)
        
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    tcPr = cell._tc.get_or_add_tcPr()
                    shd = tcPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}shd')
                    if shd is not None:
                        tcPr.remove(shd)
                    
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.color.rgb = RGBColor(0, 0, 0)
                            rPr = run._r.get_or_add_rPr()
                            highlight = rPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}highlight')
                            if highlight is not None:
                                rPr.remove(highlight)

    if standardize_typography:
        def format_paragraph(p):
            p.paragraph_format.line_spacing = line_spacing
            p.paragraph_format.space_after = Pt(space_after)
            for run in p.runs:
                run.font.name = font_name
                run.font.size = Pt(font_size)

        for paragraph in doc.paragraphs:
            format_paragraph(paragraph)
            
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        format_paragraph(paragraph)

    # Áp dụng Autofit chuẩn của Word
    if autofit_tables:
        for table in doc.tables:
            table.autofit = True
            tblPr = table._tbl.tblPr
            
            # Kích hoạt chế độ autofit layout trong XML của Word
            tblLayout = tblPr.find(qn('w:tblLayout'))
            if tblLayout is None:
                tblLayout = OxmlElement('w:tblLayout')
                tblPr.append(tblLayout)
            tblLayout.set(qn('w:type'), 'autofit')
            
            # Đặt kiểu bảng về dạng tự động mở rộng theo nội dung
            tblW = tblPr.find(qn('w:tblW'))
            if tblW is not None:
                tblW.set(qn('w:type'), 'auto')
                tblW.set(qn('w:w'), '0')
                
            # Xóa bỏ độ rộng ép buộc ở từng ô để cơ chế Autofit của Word tự tính toán
            for row in table.rows:
                for cell in row.cells:
                    tcPr = cell._tc.get_or_add_tcPr()
                    tcW = tcPr.find(qn('w:tcW'))
                    if tcW is not None:
                        tcPr.remove(tcW)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    
    with download_placeholder:
        st.download_button(
            label="Tải xuống tệp kết quả (.docx)",
            data=bio,
            file_name=f"processed_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
else:
    with download_placeholder:
        st.info("Vui lòng tải tệp lên ở trên để kích hoạt nút tải xuống.")
