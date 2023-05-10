import streamlit as st
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_LINE_SPACING
import base64

st.set_page_config(page_title="Docx to Book Formatter", page_icon=":books:")

st.title("Docx to Book Formatter")

uploaded_file = st.file_uploader("Choose a .docx file", type="docx")

if uploaded_file is not None:
    document = Document(uploaded_file)

    # Allow the user to choose specifications
    paper_size = st.selectbox("Choose paper size", ["5 x 8 inches", "5.5 x 8.5 inches", "6 x 9 inches"])
    font = st.selectbox("Choose font", ["Times New Roman", "Georgia", "Book Antiqua", "Alegreya"])
    font_size = st.slider("Choose font size", min_value=8, max_value=72, value=12)
    top_margin = st.slider("Choose top margin (in inches)", min_value=0.5, max_value=2.0, value=1.0)
    bottom_margin = st.slider("Choose bottom margin (in inches)", min_value=0.5, max_value=2.0, value=1.0)
    right_margin = st.slider("Choose right margin (in inches)", min_value=0.5, max_value=2.0, value=1.0)
    left_margin = st.slider("Choose left margin (in inches)", min_value=0.5, max_value=2.0, value=1.0)
    line_spacing_options = {"Single": WD_LINE_SPACING.SINGLE, "1.5 lines": WD_LINE_SPACING.ONE_POINT_FIVE, "Double": WD_LINE_SPACING.DOUBLE}
    line_spacing = st.select_slider("Choose line spacing", options=list(line_spacing_options.keys()))

    # Apply the chosen specifications to the document
    section = document.sections[0]
    if paper_size == "5 x 8 inches":
        section.page_height = Inches(8)
        section.page_width = Inches(5)
    elif paper_size == "5.5 x 8.5 inches":
        section.page_height = Inches(8.5)
        section.page_width = Inches(5.5)
    elif paper_size == "6 x 9 inches":
        section.page_height = Inches(9)
        section.page_width = Inches(6)
    section.left_margin = Inches(left_margin)
    section.right_margin = Inches(right_margin)
    section.top_margin = Inches(top_margin)
    section.bottom_margin = Inches(bottom_margin)

    style = document.styles['Normal']
    font_obj = style.font
    font_obj.name = font
    font_obj.size = Pt(font_size)

    for paragraph in document.paragraphs:
        paragraph.line_spacing_rule = line_spacing_options[line_spacing]

    # Save the formatted document and allow the user to download it as a .docx file
    document.save('formatted_book.docx')

    with open('formatted_book.docx', 'rb') as f:
        data = f.read()
        b64_data = base64.b64encode(data).decode()
        href = f'<a href="data:application/octet-stream;base64,{b64_data}" download="formatted_book.docx">Download formatted book</a>'
        st.markdown(href, unsafe_allow_html=True)
