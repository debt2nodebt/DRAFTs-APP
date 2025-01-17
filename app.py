import streamlit as st
import os
from docx import Document
import pypandoc
from pathlib import Path

# Automatically create the output directory if it doesn't exist
output_dir = Path("output_files")
output_dir.mkdir(exist_ok=True)

# Templates (Ensure these files exist in the templates folder)
Templates = {
    "bank_draft": "templates/Python_Bank_Draft_Template.docx",
    "cessation_draft": "templates/Python_Cessation_Template.docx",
    "settlement_draft": "templates/Python_Settlement_Draft_Template.docx"
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
        st.error(f"Error generating Word draft: {e}")

# Function to convert Word to PDF
def convert_to_pdf(word_path, pdf_path):
    try:
        pypandoc.convert_file(str(word_path), 'pdf', outputfile=str(pdf_path))
    except Exception as e:
        st.error(f"Error converting to PDF: {e}")

# Streamlit app
st.set_page_config(layout="wide")
st.title("Document Generator App")

# Bank Draft Section
st.subheader("1. Bank Draft")
with st.form("bank_draft_form"):
    col1, col2 = st.columns(2)
    with col1:
        bank_name = st.text_input("Bank Name", key="bank_draft_bank_name")
        loan_type = st.text_input("Loan Type", key="bank_draft_loan_type")
    with col2:
        loan_number = st.text_input("Loan Number", key="bank_draft_loan_number")
        client_name = st.text_input("Client Name", key="bank_draft_client_name")
        mobile_number = st.text_input("Mobile Number", key="bank_draft_mobile_number")
    submitted_bank_draft = st.form_submit_button("Generate Bank Draft")
    
    if submitted_bank_draft:
        if client_name and bank_name:
            try:
                # Replacements
                replacements = {
                    "{BankName}": bank_name,
                    "{LoanType}": loan_type,
                    "{LoanNumber}": loan_number,
                    "{ClientName}": client_name,
                    "{MobileNumber}": mobile_number
                }
                # Generate Word and PDF
                word_path = output_dir / f"{client_name}_{bank_name}_BankDraft.docx"
                pdf_path = output_dir / f"{client_name}_{bank_name}_BankDraft.pdf"
                generate_word_draft(Templates["bank_draft"], word_path, replacements)
                convert_to_pdf(word_path, pdf_path)
                st.success("Bank Draft Generated!")
                st.markdown(f"[Download Word File](./{word_path})", unsafe_allow_html=True)
                st.markdown(f"[Download PDF File](./{pdf_path})", unsafe_allow_html=True)
            except Exception as e:
                st.error(f"Error: {e}")
        else:
            st.error("Please fill in all required fields.")

# Settlement Draft Section
st.subheader("2. Settlement Draft")
with st.form("settlement_draft_form"):
    client_name_settlement = st.text_input("Client Name", key="settlement_client_name")
    settlement_amount = st.text_input("Settlement Amount", key="settlement_amount")
    settlement_bank = st.text_input("Bank Name", key="settlement_bank_name")
    submitted_settlement_draft = st.form_submit_button("Generate Settlement Draft")
    
    if submitted_settlement_draft:
        if client_name_settlement and settlement_bank:
            try:
                # Replacements
                replacements = {
                    "{ClientName}": client_name_settlement,
                    "{SettlementAmount}": settlement_amount,
                    "{BankName}": settlement_bank
                }
                # Generate Word and PDF
                word_path = output_dir / f"{client_name_settlement}_SettlementDraft.docx"
                pdf_path = output_dir / f"{client_name_settlement}_SettlementDraft.pdf"
                generate_word_draft(Templates["settlement_draft"], word_path, replacements)
                convert_to_pdf(word_path, pdf_path)
                st.success("Settlement Draft Generated!")
                st.markdown(f"[Download Word File](./{word_path})", unsafe_allow_html=True)
                st.markdown(f"[Download PDF File](./{pdf_path})", unsafe_allow_html=True)
            except Exception as e:
                st.error(f"Error: {e}")
        else:
            st.error("Please fill in all required fields.")

# Cessation Draft Section
st.subheader("3. Cessation Draft")
with st.form("cessation_draft_form"):
    employee_name = st.text_input("Employee Name", key="cessation_employee_name")
    cessation_reason = st.text_input("Reason for Cessation", key="cessation_reason")
    submitted_cessation_draft = st.form_submit_button("Generate Cessation Draft")
    
    if submitted_cessation_draft:
        if employee_name and cessation_reason:
            try:
                # Replacements
                replacements = {
                    "{EmployeeName}": employee_name,
                    "{CessationReason}": cessation_reason
                }
                # Generate Word and PDF
                word_path = output_dir / f"{employee_name}_CessationDraft.docx"
                pdf_path = output_dir / f"{employee_name}_CessationDraft.pdf"
                generate_word_draft(Templates["cessation_draft"], word_path, replacements)
                convert_to_pdf(word_path, pdf_path)
                st.success("Cessation Draft Generated!")
                st.markdown(f"[Download Word File](./{word_path})", unsafe_allow_html=True)
                st.markdown(f"[Download PDF File](./{pdf_path})", unsafe_allow_html=True)
            except Exception as e:
                st.error(f"Error: {e}")
        else:
            st.error("Please fill in all required fields.")
