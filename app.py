import pythoncom
import win32com.client
import streamlit as st
from docx import Document
from io import BytesIO
from docx2pdf import convert

pythoncom.CoInitialize()

# Updated template paths based on your GitHub structure
Templates = {
    "bank_draft": "./templates/Python Bank Draft Template.docx",
    "cessation_draft": "./templates/Python Cessation Template.docx",
    "settlement_draft": "./templates/Python settlement draft template.docx"
}

# Function to generate Word draft and return as bytes
def generate_word_draft(template_path, replacements):
    doc = Document(template_path)
    for paragraph in doc.paragraphs:
        for key, value in replacements.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, value)
    
    # Save document to bytes
    doc_io = BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# Streamlit app
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

if submitted_bank_draft and client_name and bank_name:
    replacements = {
        "{BankName}": bank_name,
        "{LoanType}": loan_type,
        "{LoanNumber}": loan_number,
        "{ClientName}": client_name,
        "{MobileNumber}": mobile_number
    }
    
    word_file = generate_word_draft(Templates["bank_draft"], replacements)
    
    # Download button for Word file
    st.download_button(
        label="Download Bank Draft (Word)",
        data=word_file,
        file_name=f"{client_name}_{bank_name}_BankDraft.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

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

if submitted_settlement_draft and client_name and bank_name:
    replacements = {
        "{BankName}": bank_name,
        "{LoanType}": loan_type,
        "{LoanNumber}": loan_number,
        "{LoanAmount}": loan_amount,
        "{OneTimePayment}": one_time_payment,
        "{ClientName}": client_name,
        "{OurMobileNumber}": our_mobile_number
    }

    word_file = generate_word_draft(Templates["settlement_draft"], replacements)
    
    # Download button for Word file
    st.download_button(
        label="Download Settlement Draft (Word)",
        data=word_file,
        file_name=f"{client_name}_{bank_name}_SettlementDraft.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

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

if submitted_cessation_draft and client_name and bank_name:
    replacements = {
        "{BankName}": bank_name,
        "{ClientName}": client_name,
        "{LoanType}": loan_type,
        "{LoanNumber}": loan_number,
        "{MobileNumber}": mobile_number
    }

    word_file = generate_word_draft(Templates["cessation_draft"], replacements)
    
    # Download button for Word file
    st.download_button(
        label="Download Cessation Draft (Word)",
        data=word_file,
        file_name=f"{client_name}_{bank_name}_CessationDraft.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

pythoncom.CoUninitialize()
