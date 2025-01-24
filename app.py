import streamlit as st
from docx import Document
from io import BytesIO
import os

# Get the absolute path to the templates directory
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
Templates = {
    "bank_draft": os.path.join(BASE_DIR, "templates", "Python Bank Draft Template.docx"),
    "cessation_draft": os.path.join(BASE_DIR, "templates", "Python Cessation Template.docx"),
    "settlement_draft": os.path.join(BASE_DIR, "templates", "Python settlement draft template.docx")
}

# Function to generate Word draft and return as bytes
def generate_word_draft(template_path, replacements):
    if not os.path.exists(template_path):
        st.error(f"Error: Template file not found at '{template_path}'")
        return None
    
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

# Streamlit app layout configuration
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

if submitted_bank_draft and client_name and bank_name:
    replacements = {
        "{BankName}": bank_name,
        "{LoanType}": loan_type,
        "{LoanNumber}": loan_number,
        "{ClientName}": client_name,
        "{MobileNumber}": mobile_number
    }
    
    word_file = generate_word_draft(Templates["bank_draft"], replacements)
    
    if word_file:
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
        bank_name = st.text_input("Bank Name", key="settlement_bank_name")
        loan_type = st.text_input("Loan Type", key="settlement_loan_type")
    with col2:
        loan_number = st.text_input("Loan Number", key="settlement_loan_number")
        loan_amount = st.text_input("Loan Amount", key="settlement_loan_amount")
    with col3:
        one_time_payment = st.text_input("One Time Payment", key="settlement_one_time_payment")
        client_name = st.text_input("Client Name", key="settlement_client_name")
        our_mobile_number = st.text_input("Mobile Number", key="settlement_mobile_number")
    
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
    
    if word_file:
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
        bank_name = st.text_input("Bank Name", key="cessation_bank_name")
        client_name = st.text_input("Client Name", key="cessation_client_name")
    with col2:
        loan_type = st.text_input("Loan Type", key="cessation_loan_type")
        loan_number = st.text_input("Loan Number", key="cessation_loan_number")
        mobile_number = st.text_input("Mobile Number", key="cessation_mobile_number")
    
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
    
    if word_file:
        st.download_button(
            label="Download Cessation Draft (Word)",
            data=word_file,
            file_name=f"{client_name}_{bank_name}_CessationDraft.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

# Option for users to upload their own templates
st.subheader("4. Upload Custom Template")
uploaded_file = st.file_uploader("Upload a Word template", type=["docx"])

if uploaded_file:
    st.success("Template uploaded successfully. Fill in the placeholders and download the document.")
    with st.form("custom_draft_form"):
        st.write("Enter values for placeholders:")
        placeholder_values = st.text_area("Provide replacements in JSON format (e.g. {\"{Placeholder1}\": \"Value1\", \"{Placeholder2}\": \"Value2\"})")
        submitted_custom_draft = st.form_submit_button("Generate Custom Document")
    
    if submitted_custom_draft and placeholder_values:
        try:
            replacements = eval(placeholder_values)
            word_file = generate_word_draft(uploaded_file, replacements)
            if word_file:
                st.download_button(
                    label="Download Custom Draft (Word)",
                    data=word_file,
                    file_name="CustomDraft.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        except Exception as e:
            st.error(f"Error processing replacements: {e}")
