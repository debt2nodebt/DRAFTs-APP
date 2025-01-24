import streamlit as st
from docx import Document
from fpdf import FPDF
from io import BytesIO
import os

def load_templates():
    templates_folder = "templates"
    template_files = [f for f in os.listdir(templates_folder) if f.endswith(".docx")]
    return template_files

def generate_word_file(template_name, user_data):
    doc = Document(os.path.join("templates", template_name))
    for para in doc.paragraphs:
        for key, value in user_data.items():
            if key in para.text:
                para.text = para.text.replace(key, value)
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

def convert_word_to_pdf(word_file):
    doc = Document(word_file)
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    
    for para in doc.paragraphs:
        try:
            pdf.multi_cell(0, 10, para.text.encode('latin-1', 'replace').decode('latin-1'))
        except UnicodeEncodeError:
            pdf.multi_cell(0, 10, para.text)
    
    pdf_output = BytesIO()
    pdf.output(pdf_output, 'F')
    pdf_output.seek(0)
    return pdf_output

st.title("Bank Draft Generator")
templates = load_templates()

selected_template = st.selectbox("Choose a template", templates)

st.subheader("Enter Details")
user_data = {}
user_data["{{Name}}"] = st.text_input("Name")
user_data["{{Amount}}"] = st.text_input("Amount")
user_data["{{Date}}"] = st.text_input("Date")

generate_button = st.button("Generate Documents")

if generate_button:
    word_file = generate_word_file(selected_template, user_data)
    st.download_button(
        label="Download Word File",
        data=word_file,
        file_name="generated_draft.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    pdf_file = convert_word_to_pdf(word_file)
    st.download_button(
        label="Download PDF File",
        data=pdf_file,
        file_name="generated_draft.pdf",
        mime="application/pdf"
    )
