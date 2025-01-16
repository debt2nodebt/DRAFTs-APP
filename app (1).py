import streamlit as st
from docx import Document
from io import BytesIO
from datetime import date
import os

# Streamlit App Configuration
st.set_page_config(layout="wide", page_title="Document Generator App")
st.title("Document Generator App")

# Initialize State for Files
if "generated_files" not in st.session_state:
    st.session_state.generated_files = {"word_file": None, "pdf_file": None}

# Upload Templates
st.sidebar.title("Upload Templates")
uploaded_bank_template = st.sidebar.file_uploader("Upload Bank Draft Template", type=["docx"])
uploaded_cessation_template = st.sidebar.file_uploader("Upload Cessation Template", type=["docx"])
uploaded_settlement_template = st.sidebar.file_uploader("Upload Settlement Template", type=["docx"])

# Helper Function: Generate Word Draft
def generate_word_draft(template_file, replacements):
    doc = Document(template_file)
    for paragraph in doc.paragraphs:
        for key, value in replacements.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, value)

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# Helper Function: Create Download Buttons
def create_download_buttons(word_file, file_name):
    if word_file:
        st.download_button(
            label="Download Word File",
            data=word_file,
            file_name=f"{file_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
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
        if uploaded_bank_template:
            # Replacements
            replacements = {
                "{BankName}": bank_name,
                "{LoanType}": loan_type,
                "{LoanNumber}": loan_number,
                "{ClientName}": client_name,
                "{MobileNumber}": mobile_number,
            }
            # Generate Word File
            word_file = generate_word_draft(uploaded_bank_template, replacements)
            st.session_state.generated_files["word_file"] = word_file
            st.success(f"Bank Draft for {client_name} generated!")
        else:
            st.error("Please upload the Bank Draft template.")

# 2. Cessation Draft Section
st.subheader("2. Cessation Draft")
with st.form("cessation_draft_form"):
    col1, col2 = st.columns(2)
    with col1:
        employer_name = st.text_input("Employer Name")
        employee_name = st.text_input("Employee Name")
    with col2:
        cessation_reason = st.text_input("Cessation Reason")
        cessation_date = st.date_input("Cessation Date", value=date.today())
    submitted_cessation_draft = st.form_submit_button("Generate Cessation Draft")

    if submitted_cessation_draft:
        if uploaded_cessation_template:
            # Replacements
            replacements = {
                "{EmployerName}": employer_name,
                "{EmployeeName}": employee_name,
                "{CessationReason}": cessation_reason,
                "{CessationDate}": cessation_date.strftime("%d-%m-%Y"),
            }
            # Generate Word File
            word_file = generate_word_draft(uploaded_cessation_template, replacements)
            st.session_state.generated_files["word_file"] = word_file
            st.success(f"Cessation Draft for {employee_name} generated!")
        else:
            st.error("Please upload the Cessation Draft template.")

# 3. Settlement Draft Section
st.subheader("3. Settlement Draft")
with st.form("settlement_draft_form"):
    col1, col2 = st.columns(2)
    with col1:
        creditor_name = st.text_input("Creditor Name")
        settlement_amount = st.text_input("Settlement Amount")
    with col2:
        due_date = st.date_input("Due Date", value=date.today())
        debtor_name = st.text_input("Debtor Name")
    submitted_settlement_draft = st.form_submit_button("Generate Settlement Draft")

    if submitted_settlement_draft:
        if uploaded_settlement_template:
            # Replacements
            replacements = {
                "{CreditorName}": creditor_name,
                "{SettlementAmount}": settlement_amount,
                "{DueDate}": due_date.strftime("%d-%m-%Y"),
                "{DebtorName}": debtor_name,
            }
            # Generate Word File
            word_file = generate_word_draft(uploaded_settlement_template, replacements)
            st.session_state.generated_files["word_file"] = word_file
            st.success(f"Settlement Draft for {debtor_name} generated!")
        else:
            st.error("Please upload the Settlement Draft template.")

# Display Download Buttons
if st.session_state.generated_files["word_file"]:
    create_download_buttons(st.session_state.generated_files["word_file"], "Generated_Document")
