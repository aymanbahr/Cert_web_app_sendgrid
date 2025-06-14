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
   - Customize subject and message
   - Click 'Start Sending Certificates'

Configuration:
--------------
- The SendGrid API key and sender email are set directly in the code.
- To change them, open `app.py` and edit the following lines (around lines 17-18):
  ```python
  sendgrid_api_key = "your_sendgrid_api_key"
  from_email = "your_verified_sender_email"
  ```
- Make sure the sender email is verified in your SendGrid dashboard.

Requires:
- Python 3.8+
- SendGrid account with API key
