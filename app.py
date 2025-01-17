import streamlit as st
import os
from docx import Document
import pypandoc  # For Word to PDF conversion
from pathlib import Path

# Output directory
output_dir = Path("output_files")
output_dir.mkdir(exist_ok=True)  # Create output directory if it doesn't exist

# Templates folder (ensure these files are uploaded to Streamlit Cloud in the correct structure)
Templates = {
    "bank_draft": "templates/Python_Bank_Draft_Template.docx",
    "cessation_draft": "templates/Python_Cessation_Template.docx",
    "settlement_draft": "templates/Python_settlement_draft_template.docx"
}

# Function to generate Word draft
def generate_word_draft(template_path, output_path, replacements):
    doc = Document(template_path)
    for paragraph in doc.paragraphs:
        for key, value in replacements.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, value)
    doc.save(output_path)

# Function to convert Word to PDF using pypandoc
def convert_to_pdf(word_path, pdf_path):
    try:
        pypandoc.convert_file(word_path, "pdf", outputfile=pdf_path)
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
            st.success(f"Bank Draft Generated: {word_path} and {pdf_path}")
            st.text_area("Generated Files", value=f"{word_path}\n{pdf_path}", height=100)
        else:
            st.error("Please fill in all required fields.")

# Settlement Draft Section
st.subheader("2. Settlement Draft")
with st.form("settlement_draft_form"):
    col1, col2, col3 = st.columns(3)
    with col1:
        bank_name = st.text_input("Bank Name")
        loan_type = st.text_input("Loan Type")
    with col2:
        loan_number = st.text_input("Loan Number")
        loan_amount = st.text_input("Loan Amount")
    with col3:
        one_time_payment = st.text_input("One Time Payment")
        client_name = st.text_input("Client Name")
        our_mobile_number = st.text_input("Mobile Number")
    submitted_settlement_draft = st.form_submit_button("Generate Settlement Draft")
    
    if submitted_settlement_draft:
        if client_name and bank_name:
            replacements = {
                "{BankName}": bank_name,
                "{LoanType}": loan_type,
                "{LoanNumber}": loan_number,
                "{LoanAmount}": loan_amount,
                "{OneTimePayment}": one_time_payment,
                "{ClientName}": client_name,
                "{OurMobileNumber}": our_mobile_number
            }
            word_path = output_dir / f"{client_name}_{bank_name}_SettlementDraft.docx"
            pdf_path = output_dir / f"{client_name}_{bank_name}_SettlementDraft.pdf"
            generate_word_draft(Templates["settlement_draft"], word_path, replacements)
            convert_to_pdf(word_path, pdf_path)
            st.success(f"Settlement Draft Generated: {word_path} and {pdf_path}")
            st.text_area("Generated Files", value=f"{word_path}\n{pdf_path}", height=100)
        else:
            st.error("Please fill in all required fields.")

# Cessation Draft Section
st.subheader("3. Cessation Draft")
with st.form("cessation_draft_form"):
    col1, col2 = st.columns(2)
    with col1:
        bank_name = st.text_input("Bank Name")
        client_name = st.text_input("Client Name")
    with col2:
        loan_type = st.text_input("Loan Type")
        loan_number = st.text_input("Loan Number")
        mobile_number = st.text_input("Mobile Number")
    submitted_cessation_draft = st.form_submit_button("Generate Cessation Draft")
    
    if submitted_cessation_draft:
        if client_name and bank_name:
            replacements = {
                "{BankName}": bank_name,
                "{ClientName}": client_name,
                "{LoanType}": loan_type,
                "{LoanNumber}": loan_number,
                "{MobileNumber}": mobile_number
            }
            word_path = output_dir / f"{client_name}_{bank_name}_CessationDraft.docx"
            pdf_path = output_dir / f"{client_name}_{bank_name}_CessationDraft.pdf"
            generate_word_draft(Templates["cessation_draft"], word_path, replacements)
            convert_to_pdf(word_path, pdf_path)
            st.success(f"Cessation Draft Generated: {word_path} and {pdf_path}")
            st.text_area("Generated Files", value=f"{word_path}\n{pdf_path}", height=100)
        else:
            st.error("Please fill in all required fields.")
