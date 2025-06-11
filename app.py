import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import os
import re
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail, Attachment, FileContent, FileName, FileType, Disposition
import base64

st.set_page_config(page_title="Certificate Sender (SendGrid API)", layout="centered")
st.title("Certificate Sender & Cleaner with SendGrid API")

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
            page.add_redact_annot(inst, fill=(1, 1, 1))
            page.insert_textbox(inst, name, fontsize=24, color=(0, 0, 0), align=1)
    doc.save(output_path)
    doc.close()

def send_email_via_sendgrid(api_key, from_email, to_email, subject, html_content, file_path):
    message = Mail(
        from_email=from_email,
        to_emails=to_email,
        subject=subject,
        html_content=html_content
    )
    with open(file_path, 'rb') as f:
        data = f.read()
        encoded = base64.b64encode(data).decode()
        attachedFile = Attachment(
            FileContent(encoded),
            FileName(os.path.basename(file_path)),
            FileType('application/pdf'),
            Disposition('attachment')
        )
        message.attachment = attachedFile

    try:
        sg = SendGridAPIClient(api_key)
        response = sg.send(message)
        return response.status_code
    except Exception as e:
        return f"Error: {e}"

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

    button_clicked = st.button("Start Sending Certificates")

    if button_clicked:
        if not sendgrid_api_key or not from_email:
            st.error("Please enter your SendGrid API key and sender email.")
        else:
            valid_df = df[df["Valid Name"] & df["Valid Email"]]
            os.makedirs("output", exist_ok=True)
            for _, row in valid_df.iterrows():
                name = row["Name"]
                email = row["Email"]
                first_name = name.split()[0]
                cert_path = f"output/{name}.pdf"
                generate_certificate(name, uploaded_pdf.read(), cert_path)

                html_message = email_body.replace("{first_name}", first_name).replace("\n", "<br>")
                status = send_email_via_sendgrid(
                    api_key=sendgrid_api_key,
                    from_email=from_email,
                    to_email=email,
                    subject=email_subject.replace("{first_name}", first_name),
                    html_content=html_message,
                    file_path=cert_path
                )
                st.write(f"{name} ({email}) → Status: {status}")
            st.success("All certificates sent!")
