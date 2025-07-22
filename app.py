import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import os
import re
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail, Attachment, FileContent, FileName, FileType, Disposition
import base64

st.set_page_config(page_title="Certificate Sender (SendGrid)", layout="centered")

# Login credentials with fallback for local testing
try:
    LOGIN_EMAIL = st.secrets["LOGIN_EMAIL"]
    LOGIN_PASSWORD = st.secrets["LOGIN_PASSWORD"]
    sendgrid_api_key = st.secrets["SENDGRID_API_KEY"]
    from_email = st.secrets["FROM_EMAIL"]
except Exception:
    LOGIN_EMAIL = "Marketing@volaris-global.com"
    LOGIN_PASSWORD = "CER@VoL#20&GO"
    sendgrid_api_key = "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"
    from_email = "ahmed.mahmoud@volaris-global.com"

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

# Add color picker and font size selection for the name
name_color_hex = st.color_picker("Pick a color for the name", "#000000")
font_size = st.slider("Font size for the name", min_value=10, max_value=240, value=120)

# Remove font family selection, always use uploaded font if provided, else use helv
uploaded_font = st.file_uploader("Upload a custom font (.ttf) for the name (optional)", type=["ttf"])
fontfile = None
if uploaded_font is not None:
    fontfile = "custom_font.ttf"
    with open(fontfile, "wb") as f:
        f.write(uploaded_font.read())
default_fontname = "helv"  # Must be exactly this, with a capital B
# Extra robustness: force correct case
if default_fontname.lower() == "helv":
    default_fontname = "helv"

# Custom font upload (Google Fonts or any TTF)
# Font family selection (case-sensitive for PyMuPDF)
font_options = [
    ("Helvetica", "helv"),
    ("Helvetica Bold", "helv"),
    ("Helvetica Italic", "helvI"),
    ("Helvetica Bold Italic", "helvI"),
    ("Times", "times"),
    ("Times Bold", "timesB"),
    ("Times Italic", "timesI"),
    ("Times Bold Italic", "timesBI"),
    ("Courier", "cour"),
    ("Courier Bold", "courB"),
    ("Courier Italic", "courI"),
    ("Courier Bold Italic", "courBI"),
]
font_display = [f[0] for f in font_options]
font_map = {f[0]: f[1] for f in font_options}  # No lowercasing
# selected_font_display = st.selectbox("Font family for the name", font_display, index=1)
# selected_font = font_map[selected_font_display]
# Fix for common case errors
font_case_map = {
    "helv": "helv", "helvi": "helvI", "helvi": "helvI",
    "timesb": "timesB", "timesi": "timesI", "timesbi": "timesBI",
    "courb": "courB", "couri": "courI", "courbi": "courBI"
}
# selected_font = font_case_map.get(selected_font, selected_font)

# Read uploaded PDF only once and reuse bytes using session_state
if uploaded_pdf and "template_bytes" not in st.session_state:
    st.session_state["template_bytes"] = uploaded_pdf.read()
template_bytes = st.session_state.get("template_bytes", None)

if template_bytes:
    import fitz
    doc = fitz.open(stream=template_bytes, filetype="pdf")
    page = doc[0]
    page_width = page.rect.width
    # Centered horizontally, 400 from top, height 60, width 60% of page
    rect_width = int(page_width * 0.6)
    x0_default = int((page_width - rect_width) / 2)
    x1_default = int(x0_default + rect_width)
    y0_default = 850
    y1_default = y0_default + 60
    doc.close()
else:
    x0_default = 200
    x1_default = 1000
    y0_default = 400
    y1_default = 460

# Manual placement sliders with calculated defaults
st.subheader("Manual Name Placement (if needed)")
x0 = st.slider("X (left)", 0, 1200, x0_default)
y0 = st.slider("Y (top)", 0, 8000, y0_default)
x1 = st.slider("X (right)", 0, 1200, x1_default)
y1 = st.slider("Y (bottom)", 0, 8000, y1_default)
default_rect = (x0, y0, x1, y1)

def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16)/255 for i in (0, 2, 4))

def generate_certificate(name, template_bytes, output_path, name_color=(0, 0, 0), font_size=60, default_rect=None, fontname="helv", fontfile=None):
    doc = fitz.open(stream=template_bytes, filetype="pdf")
    for page in doc:
        found = False
        for inst in page.search_for("<fullName>"):
            rect = fitz.Rect(inst)
            found = True
            break
        if not found and default_rect:
            rect = fitz.Rect(*default_rect)
        if found or default_rect:
            if fontfile:
                name_width = fitz.get_text_length(name, fontsize=font_size, fontfile=fontfile)
            else:
                name_width = fitz.get_text_length(name, fontsize=font_size, fontname=fontname)
            name_x = rect.x0 + (rect.width - name_width) / 2
            name_y = rect.y0 + rect.height / 2 + font_size / 2
            page.insert_text(
                (name_x, name_y),
                name,
                fontsize=font_size,
                color=name_color,
                fontname=None if fontfile else fontname,
                fontfile=fontfile
            )
    doc.save(output_path)
    doc.close()
    if fontfile and os.path.exists(fontfile):
        os.remove(fontfile)

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
        name_color = hex_to_rgb(name_color_hex)
        # Remove debug print and preview
        for _, row in valid_df.iterrows():
            name = row["Name"]
            email = row["Email"]
            first_name = name.split()[0]
            cert_path = f"output/{name}.pdf"
            # Always use correct case for fontname
            generate_certificate(name, template_bytes, cert_path, name_color=name_color, font_size=font_size, default_rect=default_rect, fontname=default_fontname, fontfile=fontfile)
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
