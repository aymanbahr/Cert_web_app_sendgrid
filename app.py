import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import os
import re
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail, Attachment, FileContent, FileName, FileType, Disposition
import base64

st.set_page_config(page_title="Certificate Sender (SendGrid)", layout="centered")

# Login credentials
LOGIN_EMAIL = st.secrets.get("LOGIN_EMAIL")
LOGIN_PASSWORD = st.secrets.get("LOGIN_PASSWORD")  # Change this to your desired password

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    st.title("Login")
    email = st.text_input("Email")
    password = st.text_input("Password", type="password")
    if st.button("Login"):
        if email == LOGIN_EMAIL and password == LOGIN_PASSWORD:
            st.session_state.logged_in = True
            st.success("Login successful!")
            st.rerun()
        else:
            st.error("Invalid email or password.")
    st.stop()  # Prevents the rest of the app from running if not logged in

st.title("Volaris Certificate Cleaner & Sender")

uploaded_excel = st.file_uploader("Upload Attendee Excel (.xlsx)", type=["xlsx"])
uploaded_pdf = st.file_uploader("Upload Certificate Template (.pdf)", type=["pdf"])

sendgrid_api_key = st.secrets.get("SENDGRID_API_KEY")
from_email = st.secrets.get("FROM_EMAIL")

if not sendgrid_api_key or not from_email:
    st.error("SendGrid API Key or Sender Email not set in environment variables.")
    st.stop()

event_name = st.text_input("Event Name")
event_date = st.text_input("Event Date")
client_company = st.text_input("Client Company")


# Ensure required fields are filled
if not event_name or not event_date or not client_company:
    st.warning('Please fill in all required fields: Event Name, Event Date, and Client Company.')
    st.stop()
email_subject = st.text_input("Email Subject", value=f"Your Attendance Certificate {event_name}_{event_date}")
email_body = st.text_area("Email Body (use {first_name})", value=f"""Dear Dr {{first_name}},

Thank you for attending {event_name}.  

Your certificate is attached.

Best Regards,  

Volaris Team on behalf of {client_company}""")

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
            rect = fitz.Rect(inst)
            font_size = 24
            # Calculate the width of the name
            name_width = fitz.get_text_length(name, fontsize=font_size)
            # Center the name in the placeholder rectangle
            name_x = rect.x0 + (rect.width - name_width) / 2
            name_y = rect.y0 + rect.height / 2 + font_size / 2  # Vertically center
            # Redact the placeholder
            page.add_redact_annot(rect)
            page.apply_redactions()
            # Insert the name
            page.insert_text((name_x, name_y), name, fontsize=font_size, color=(0, 0, 0))
    doc.save(output_path)
    doc.close()

if uploaded_excel and uploaded_pdf:
    template_bytes = uploaded_pdf.read()  # Read once here
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
        valid_df = df[df["Valid Name"] & df["Valid Email"]]
        os.makedirs("output", exist_ok=True)
        for _, row in valid_df.iterrows():
            name = row["Name"]
            email = row["Email"]
            first_name = name.split()[0]
            cert_path = f"output/{name}.pdf"
            generate_certificate(name, template_bytes, cert_path)
            # Read and encode PDF
            with open(cert_path, "rb") as f:
                data = f.read()
                encoded = base64.b64encode(data).decode()
            attachment = Attachment(
                FileContent(encoded),
                FileName(f"{name}.pdf"),
                FileType("application/pdf"),
                Disposition("attachment")
            )
            message = Mail(
                from_email=from_email,
                to_emails=email,
                subject=email_subject.format(first_name=first_name),
                html_content=email_body.format(first_name=first_name)
            )
            message.attachment = attachment
            try:
                sg = SendGridAPIClient(sendgrid_api_key)
                sg.send(message)
                st.write(f"Sent to {name} ({email})")
            except Exception as e:
                st.error(f"Failed to send to {name} ({email}): {e}")
        st.success("All certificates sent!")


# 📥 Download All Certificates
st.markdown("---")
if os.path.isdir("output"):
    import zipfile
    from io import BytesIO

    def zip_output_folder():
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for root, _, files in os.walk("output"):
                for file in files:
                    zipf.write(os.path.join(root, file), arcname=file)
        zip_buffer.seek(0)
        return zip_buffer

    zip_data = zip_output_folder()
    st.download_button(
        label="📦 Download All Certificates (ZIP)",
        data=zip_data,
        file_name="All_Certificates.zip",
        mime="application/zip"
    )
