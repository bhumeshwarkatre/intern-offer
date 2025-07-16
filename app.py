import os
import re
import qrcode
import random
import string
import tempfile
import base64
import streamlit as st
from datetime import date
from smtplib import SMTP
from docxtpl import DocxTemplate
from docx.shared import Inches
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email import encoders
import pandas as pd
import logging

# ✅ Aspose Cloud
import asposewordscloud
from asposewordscloud.apis.words_api import WordsApi
from asposewordscloud.models.requests import UploadFileRequest, SaveAsRequest, DownloadFileRequest
from asposewordscloud.models import PdfSaveOptionsData

# ✅ Google Sheets
from google.oauth2.service_account import Credentials
import gspread

# --- Logging ---
logging.basicConfig(level=logging.INFO)

# --- Streamlit Config ---
st.set_page_config("Intern Offer Generator", layout="wide")
EMAIL = st.secrets["email"]["user"]
PASSWORD = st.secrets["email"]["password"]
ADMIN_KEY = st.secrets["admin"]["key"]
TEMPLATE_FILE = os.path.join(tempfile.gettempdir(), "offer_template.docx")
LOGO = "logo.png"

# --- Aspose Setup ---
api_sid = st.secrets["aspose"]["app_sid"]
api_key = st.secrets["aspose"]["app_key"]
words_api = WordsApi(api_sid, api_key)

# --- Decode Template from Base64 on First Run ---
if not os.path.exists(TEMPLATE_FILE):
    encoded_template = st.secrets["template_base64"]["template_base64"]
    with open(TEMPLATE_FILE, "wb") as f:
        f.write(base64.b64decode(encoded_template))

# --- Google Sheet Setup ---
def get_gsheet():
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scope
    )
    client = gspread.authorize(creds)
    return client.open("intern_offers").sheet1

def save_to_gsheet(data):
    try:
        sheet = get_gsheet()
        row = [data['i_id'], data['intern_name'], data['domain'],
               data['start_date'], data['end_date'], data['offer_date'], data['email']]
        sheet.append_row(row)
    except Exception as e:
        st.warning("⚠️ Could not log to Google Sheet.")
        logging.exception("Google Sheet error:")

# --- Utility Functions ---
def format_date(d):
    return d.strftime("%A, %d %B %Y")

def generate_certificate_key():
    return ''.join(random.choices(string.ascii_uppercase + string.digits, k=9))

def generate_qr(data):
    qr = qrcode.QRCode(box_size=10, border=0)
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    path = os.path.join(tempfile.gettempdir(), "qr.png")
    img.save(path)
    return path

def send_email(receiver, pdf_path, data):
    msg = MIMEMultipart()
    msg['From'] = EMAIL
    msg['To'] = receiver
    msg['Subject'] = f"🎉 Congratulations {data['intern_name']}! Your Internship Offer"

    html = f"""
    <html>
    <body>
        <p>Dear {data["intern_name"]},</p>
        <p>We are delighted to offer you an <strong>Internship Opportunity</strong> at <strong>SkyHighes Technology</strong>!</p>
        <p>Your offer letter is attached. Kindly review and confirm your acceptance.</p>
        <p>Regards,<br><strong>SkyHighes Team</strong></p>
    </body>
    </html>
    """
    msg.attach(MIMEText(html, 'html'))

    with open(pdf_path, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(pdf_path)}")
        msg.attach(part)

    with SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(EMAIL, PASSWORD)
        server.send_message(msg)

# --- UI Header ---
st.markdown("""
<style>
.title-text { font-size: 2rem; font-weight: 700; }
.stButton>button { background-color: #1E88E5; color: white; padding: 0.5rem 1.5rem; border-radius: 8px; font-weight: 600; }
</style>
""", unsafe_allow_html=True)

with st.container():
    col_logo, col_title = st.columns([1, 6])
    with col_logo:
        if os.path.exists(LOGO):
            st.image(LOGO, width=80)
    with col_title:
        st.markdown('<div class="title-text">SkyHighes Technologies Internship Letter Portal</div>', unsafe_allow_html=True)

st.divider()

# --- Form ---
with st.form("offer_form"):
    st.subheader("🎓 Internship Offer Letter Generator")

    col1, col2, col3 = st.columns(3)
    intern_name = col1.text_input("Intern Name")
    domain = col2.text_input("Domain")
    email = col3.text_input("Recipient Email")

    col4, col5, col6 = st.columns(3)
    start_date = col4.date_input("Start Date", value=date.today())
    end_date = col5.date_input("End Date", value=date.today())
    offer_date = col6.date_input("Offer Date", value=date.today())

    submit = st.form_submit_button("🚀 Generate & Send Offer Letter")

# --- On Submit ---
if submit:
    if not all([intern_name, domain, email]):
        st.error("❌ Please fill all fields.")
    elif not re.match(r"[^@]+@[^@]+\.[^@]+", email):
        st.warning("⚠️ Invalid email.")
    elif end_date < start_date:
        st.warning("⚠️ End date cannot be before start date.")
    else:
        intern_id = generate_certificate_key()
        data = {
            "intern_name": intern_name.title().strip(),
            "domain": domain.title().strip(),
            "start_date": format_date(start_date),
            "end_date": format_date(end_date),
            "offer_date": format_date(offer_date),
            "i_id": intern_id,
            "email": email.lower().strip()
        }

        try:
            save_to_gsheet(data)
            doc = DocxTemplate(TEMPLATE_FILE)
            doc.render(data)

            qr_path = generate_qr(", ".join(data.values()))
            try:
                doc.tables[0].rows[0].cells[2].paragraphs[0].add_run().add_picture(qr_path, width=Inches(1.4))
            except:
                st.warning("⚠️ QR insertion failed.")

            docx_path = os.path.join(tempfile.gettempdir(), f"{intern_id}.docx")
            doc.save(docx_path)

            cloud_path = f"temp/{intern_id}.docx"
            cloud_pdf_name = f"Offer_{intern_name}.pdf"
            local_pdf_path = os.path.join(tempfile.gettempdir(), cloud_pdf_name)

            with open(docx_path, "rb") as f:
                words_api.upload_file(UploadFileRequest(f, cloud_path))

            save_opts = PdfSaveOptionsData(file_name=cloud_pdf_name)
            words_api.save_as(SaveAsRequest(name=cloud_path, save_options_data=save_opts))

            pdf_stream = words_api.download_file(DownloadFileRequest(path=cloud_pdf_name))
            with open(local_pdf_path, "wb") as f:
                f.write(pdf_stream)

            send_email(email, local_pdf_path, data)
            st.success(f"✅ Offer letter sent to {email}")

            with open(local_pdf_path, "rb") as f:
                st.download_button("📥 Download Offer Letter", f, file_name=os.path.basename(local_pdf_path))

        except Exception as e:
            logging.exception("Error during document generation or upload:")
            st.error(f"❌ Error occurred during processing. Please check the logs.")

# --- Admin Panel ---
st.divider()
with st.expander("🔐 Admin Panel"):
    admin_key = st.text_input("Enter Admin Key", type="password")
    if admin_key == ADMIN_KEY:
        st.success("✅ Access granted.")
        try:
            sheet = get_gsheet()
            df = pd.DataFrame(sheet.get_all_records())
            st.dataframe(df if not df.empty else pd.DataFrame(columns=["No Records"]))
        except Exception as e:
            logging.exception("Failed to load sheet:")
            st.error("❌ Failed to load Google Sheet.")
    elif admin_key:
        st.error("❌ Invalid key.")

st.markdown("<hr><center><small>© 2025 SkyHighes Technologies. All Rights Reserved.</small></center>", unsafe_allow_html=True)
