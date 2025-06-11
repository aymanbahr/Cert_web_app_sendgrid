import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import os
import re
import yagmail

st.set_page_config(page_title="Certificate Sender (SendGrid)", layout="centered")

st.title("Volaris Certificate Cleaner & Sender")

uploaded_excel = st.file_uploader("Upload Attendee Excel (.xlsx)", type=["xlsx"])
uploaded_pdf = st.file_uploader("Upload Certificate Template (.pdf)", type=["pdf"])

sendgrid_api_key = st.text_input("SendGrid API Key", type="password")
from_email = st.text_input("Sender Email (verified with SendGrid)")
email_subject = st.text_input("Email Subject", value="Your Attendance Certificate")
email_body = st.text_area("Email Body (use {first_name})", value="Dear {first_name},\n\nThank you for attending.\nYour certificate is attached.")

def is_valid_name(name):
    bad_words = ['correct', 'yes', 'no', 'test', 'none', 'n/a', '123', 'nil']
    if not isinstance(name, str): return False
    if name.strip().lower() in bad_words: return False
    if len(name.strip().split()) < 2: return False
    if any(char.isdigit() for char in name): return False
    return True

def is_valid_email(email):
    if not isinstance(email, str): return False
    regex = r"^[\w\.-]+@[\w\.-]+\.\w+$"
    return bool(re.match(regex, email.strip()))

def generate_certificate(name, template_bytes, output_path):
    doc = fitz.open(stream=template_bytes, filetype="pdf")
    for page in doc:
        for inst in page.search_for("<fullName>"):
            page.insert_text(inst[:2], name, fontsize=24, color=(0, 0, 0))
            page.add_redact_annot(inst)
        page.apply_redactions()
    doc.save(output_path)
    doc.close()

if uploaded_excel and uploaded_pdf:
    df = pd.read_excel(uploaded_excel)
    df["Valid Name"] = df["Name"].apply(is_valid_name)
    df["Valid Email"] = df["Email"].apply(is_valid_email)

    st.subheader("Data Validation")
    invalid_df = df[~(df["Valid Name"] & df["Valid Email"])]
    if not invalid_df.empty:
        st.warning("Some rows have invalid names or emails.")
        st.dataframe(invalid_df)
        st.download_button("Download Rows Needing Review", invalid_df.to_csv(index=False), "invalid_rows.csv")
    else:
        st.success("All data looks good!")

    if st.button("Start Sending Certificates") and sendgrid_api_key and from_email:
        yag = yagmail.SMTP(
            user="apikey",
            password=sendgrid_api_key,
            host="smtp.sendgrid.net",
            port=587,
            smtp_starttls=True
        )

        valid_df = df[df["Valid Name"] & df["Valid Email"]]
        os.makedirs("output", exist_ok=True)
        for _, row in valid_df.iterrows():
            name = row["Name"]
            email = row["Email"]
            first_name = name.split()[0]
            cert_path = f"output/{name}.pdf"
            generate_certificate(name, uploaded_pdf.read(), cert_path)
            yag.send(
                to=email,
                subject=email_subject.format(first_name=first_name),
                contents=email_body.format(first_name=first_name),
                attachments=cert_path,
                headers={"From": from_email}
            )
            st.write(f"Sent to {name} ({email})")
        st.success("All certificates sent!")
