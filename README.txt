Certificate Sender with SendGrid

How to Use:
-----------
1. Install dependencies:
   pip install -r requirements.txt

2. Run the app:
   streamlit run app.py

3. In the browser:
   - Upload attendee Excel file with 'Name' and 'Email' columns
   - Upload a certificate template PDF with <fullName> as the placeholder
   - Paste your SendGrid API key
   - Enter sender email (must be verified in your SendGrid dashboard)
   - Customize subject and message
   - Click 'Start Sending Certificates'

Requires:
- Python 3.8+
- SendGrid account with API key
