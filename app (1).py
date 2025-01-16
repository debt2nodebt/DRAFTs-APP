import pythoncom
import win32com.client
import streamlit as st
import os
from docx import Document
from docx2pdf import convert

# Initialize COM object
pythoncom.CoInitialize()

# Output directory
output_dir = "."
os.makedirs(output_dir, exist_ok=True)

# Templates
Templates = {
    "bank_draft": "./DRAFTs-APP/Python Bank Draft Template.docx",
    "cessation_draft": "./DRAFTs-APP/Python Cessation Template.docx",
    "settlement_draft": "./DRAFTs-APP/Python settlement draft template.docx"
}

# Word application object
word = win32com.client.Dispatch("Word.Application")

# Function to generate Word draft
def generate_word_draft(template_path, output_path, replacements):
    doc = Document(template_path)
    for paragraph in doc.paragraphs:
        for key, value in replacements.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, value)
    doc.save(output_path)

# Streamlit app configuration
st.set_page_config(layout="wide")
st.title("Document Generator App")

# Initialize state variables for storing file paths
if "generated_files" not in st.session_state:
    st.session_state.generated_files = {"word_path": None, "pdf_path": None}

# Function to create download buttons
def create_download_buttons(word_output_path, pdf_output_path):
    if word_output_path and os.path.exists(word_output_path):
        with open(word_output_path, "rb") as file:
            st.download_button(
                label="Download Word File",
                data=file,
                file_name=os.path.basename(word_output_path),
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    if pdf_output_path and os.path.exists(pdf_output_path):
        with open(pdf_output_path, "rb") as file:
            st.download_button(
                label="Download PDF File",
                data=file,
                file_name=os.path.basename(pdf_output_path),
                mime="application/pdf"
            )

# 1. Bank Draft Section
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
            # Replacements
            replacements = {
                "{BankName}": bank_name,
                "{LoanType}": loan_type,
                "{LoanNumber}": loan_number,
                "{ClientName}": client_name,
                "{MobileNumber}": mobile_number
            }
            # Generate Word and PDF
            word_path = f"{output_dir}\\{client_name}_{bank_name}_BankDraft.docx"
            pdf_path = f"{output_dir}\\{client_name}_{bank_name}_BankDraft.pdf"
            generate_word_draft(Templates["bank_draft"], word_path, replacements)
            convert(word_path)  # Convert Word to PDF
            
            # Store file paths in session state
            st.session_state.generated_files["word_path"] = word_path
            st.session_state.generated_files["pdf_path"] = pdf_path
            
            st.success(f"Bank Draft Generated!")

# 2. Cessation Draft Section
st.subheader("2. Cessation Draft")
with st.form("cessation_draft_form"):
    col1, col2 = st.columns(2)
    with col1:
        employer_name = st.text_input("Employer Name")
        employee_name = st.text_input("Employee Name")
    with col2:
        cessation_reason = st.text_input("Cessation Reason")
        cessation_date = st.date_input("Cessation Date")
    submitted_cessation_draft = st.form_submit_button("Generate Cessation Draft")
    
    if submitted_cessation_draft:
        if employee_name and employer_name:
            # Replacements
            replacements = {
                "{EmployerName}": employer_name,
                "{EmployeeName}": employee_name,
                "{CessationReason}": cessation_reason,
                "{CessationDate}": cessation_date.strftime("%d-%m-%Y")
            }
            # Generate Word and PDF
            word_path = f"{output_dir}\\{employee_name}_CessationDraft.docx"
            pdf_path = f"{output_dir}\\{employee_name}_CessationDraft.pdf"
            generate_word_draft(Templates["cessation_draft"], word_path, replacements)
            convert(word_path)  # Convert Word to PDF
            
            # Store file paths in session state
            st.session_state.generated_files["word_path"] = word_path
            st.session_state.generated_files["pdf_path"] = pdf_path
            
            st.success(f"Cessation Draft Generated!")

# 3. Settlement Draft Section
st.subheader("3. Settlement Draft")
with st.form("settlement_draft_form"):
    col1, col2 = st.columns(2)
    with col1:
        creditor_name = st.text_input("Creditor Name")
        settlement_amount = st.text_input("Settlement Amount")
    with col2:
        due_date = st.date_input("Due Date")
        debtor_name = st.text_input("Debtor Name")
    submitted_settlement_draft = st.form_submit_button("Generate Settlement Draft")
    
    if submitted_settlement_draft:
        if creditor_name and debtor_name:
            # Replacements
            replacements = {
                "{CreditorName}": creditor_name,
                "{SettlementAmount}": settlement_amount,
                "{DueDate}": due_date.strftime("%d-%m-%Y"),
                "{DebtorName}": debtor_name
            }
            # Generate Word and PDF
            word_path = f"{output_dir}\\{debtor_name}_SettlementDraft.docx"
            pdf_path = f"{output_dir}\\{debtor_name}_SettlementDraft.pdf"
            generate_word_draft(Templates["settlement_draft"], word_path, replacements)
            convert(word_path)  # Convert Word to PDF
            
            # Store file paths in session state
            st.session_state.generated_files["word_path"] = word_path
            st.session_state.generated_files["pdf_path"] = pdf_path
            
            st.success(f"Settlement Draft Generated!")

# Display download buttons after any form submission
if st.session_state.generated_files["word_path"] and st.session_state.generated_files["pdf_path"]:
    create_download_buttons(
        st.session_state.generated_files["word_path"],
        st.session_state.generated_files["pdf_path"]
    )
