import streamlit as st
import os
from docx import Document
import pypandoc  # Import for Word to PDF conversion

# Output directory
output_dir = "output_files"
os.makedirs(output_dir, exist_ok=True)

# Templates
Templates = {
    "bank_draft": "templates/Python Bank Draft Template.docx",
    "cessation_draft": "templates/Python Cessation Template.docx",
    "settlement_draft": "templates/Python Settlement Draft Template.docx"
}

# Function to generate Word draft
def generate_word_draft(template_path, output_path, replacements):
    doc = Document(template_path)
    for paragraph in doc.paragraphs:
        for key, value in replacements.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, value)
    doc.save(output_path)

# Function to convert Word to PDF
def convert_to_pdf(input_path, output_path):
    output = pypandoc.convert_file(input_path, "pdf", outputfile=output_path)
    if output:
        return output_path
    else:
        raise RuntimeError("Failed to convert to PDF.")

# Streamlit app
st.set_page_config(layout="wide")
st.title("Document Generator App")

# Draft Sections
def handle_draft(template_key, fields, output_name):
    if all(fields.values()):
        replacements = {f"{{{key}}}": value for key, value in fields.items()}
        word_path = os.path.join(output_dir, f"{output_name}.docx")
        pdf_path = os.path.join(output_dir, f"{output_name}.pdf")
        generate_word_draft(Templates[template_key], word_path, replacements)
        try:
            convert_to_pdf(word_path, pdf_path)
            st.success(f"Draft Generated: {word_path} and {pdf_path}")
            st.text_area("Generated Files", value=f"{word_path}\n{pdf_path}", height=100)
        except RuntimeError as e:
            st.error(f"Error: {e}")
    else:
        st.error("Please fill in all required fields.")

# Bank Draft Section
st.subheader("1. Bank Draft")
with st.form("bank_draft_form"):
    fields = {
        "BankName": st.text_input("Bank Name"),
        "LoanType": st.text_input("Loan Type"),
        "LoanNumber": st.text_input("Loan Number"),
        "ClientName": st.text_input("Client Name"),
        "MobileNumber": st.text_input("Mobile Number")
    }
    submitted = st.form_submit_button("Generate Bank Draft")
    if submitted:
        handle_draft("bank_draft", fields, f"{fields['ClientName']}_{fields['BankName']}_BankDraft")

# Settlement Draft Section
st.subheader("2. Settlement Draft")
with st.form("settlement_draft_form"):
    fields = {
        "BankName": st.text_input("Bank Name"),
        "LoanType": st.text_input("Loan Type"),
        "LoanNumber": st.text_input("Loan Number"),
        "LoanAmount": st.text_input("Loan Amount"),
        "OneTimePayment": st.text_input("One Time Payment"),
        "ClientName": st.text_input("Client Name"),
        "OurMobileNumber": st.text_input("Mobile Number")
    }
    submitted = st.form_submit_button("Generate Settlement Draft")
    if submitted:
        handle_draft("settlement_draft", fields, f"{fields['ClientName']}_{fields['BankName']}_SettlementDraft")

# Cessation Draft Section
st.subheader("3. Cessation Draft")
with st.form("cessation_draft_form"):
    fields = {
        "BankName": st.text_input("Bank Name"),
        "ClientName": st.text_input("Client Name"),
        "LoanType": st.text_input("Loan Type"),
        "LoanNumber": st.text_input("Loan Number"),
        "MobileNumber": st.text_input("Mobile Number")
    }
    submitted = st.form_submit_button("Generate Cessation Draft")
    if submitted:
        handle_draft("cessation_draft", fields, f"{fields['ClientName']}_{fields['BankName']}_CessationDraft")
