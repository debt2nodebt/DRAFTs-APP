import streamlit as st
from docx import Document
import pdfkit
from pathlib import Path

# Output directory
output_dir = Path("C:/Dell/user/output_files")
output_dir.mkdir(parents=True, exist_ok=True)  # Create output directory if it doesn't exist

# Templates folder (ensure these files exist in your environment)
Templates = {
    "bank_draft": "templates/Python_Bank_Draft_Template.docx",
    "cessation_draft": "templates/Python_Cessation_Template.docx",
    "settlement_draft": "templates/Python_settlement_draft_template.docx"
}

# Function to generate Word draft
def generate_word_draft(template_path, output_path, replacements):
    try:
        doc = Document(template_path)
        for paragraph in doc.paragraphs:
            for key, value in replacements.items():
                if key in paragraph.text:
                    paragraph.text = paragraph.text.replace(key, value)
        doc.save(output_path)
    except Exception as e:
        st.error(f"Error generating Word file: {e}")

# Function to convert Word to PDF using pdfkit
def convert_to_pdf(word_path, pdf_path):
    try:
        # Convert Word document to HTML format
        html_path = word_path.with_suffix(".html")
        html_content = Path(word_path).read_text(encoding='utf-8')
        pdfkit.from_string(html_content, pdf_path)
        st.success(f"PDF generated: {pdf_path}")
    except Exception as e:
        st.error(f"Failed to convert Word to PDF: {e}")

# Streamlit app setup
st.set_page_config(layout="wide")
st.title("Document Generator App")

# Bank Draft Section
st.subheader("1. Bank Draft")
with st.form("bank_draft_form"):
    col1, col2 = st.columns(2)
    with col1:
        bank_name = st.text_input("Bank Name")
        loan_type = st.text_input("Loan Type")
    with col2:
        loan_number = st.text_input("Loan Number")
        client_name = st.text_input("Client Name")
        mobile_number = st.text_input("Mobile Number")
    submitted_bank_draft = st.form_submit_button("Generate Bank Draft")
    
    if submitted_bank_draft:
        if client_name and bank_name:
            replacements = {
                "{BankName}": bank_name,
                "{LoanType}": loan_type,
                "{LoanNumber}": loan_number,
                "{ClientName}": client_name,
                "{MobileNumber}": mobile_number
            }
            word_path = output_dir / f"{client_name}_{bank_name}_BankDraft.docx"
            pdf_path = output_dir / f"{client_name}_{bank_name}_BankDraft.pdf"
            generate_word_draft(Templates["bank_draft"], word_path, replacements)
            convert_to_pdf(word_path, pdf_path)
            st.text(f"Files saved at: {output_dir}")
