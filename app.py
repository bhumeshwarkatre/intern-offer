import os
import re
import csv
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

# ✅ Aspose Cloud
import asposewordscloud
from asposewordscloud.apis.words_api import WordsApi
from asposewordscloud.models.requests import UploadFileRequest, SaveAsRequest, DownloadFileRequest
from asposewordscloud.models import PdfSaveOptionsData

# --- Setup Aspose Cloud Client ---
api_sid = st.secrets["aspose"]["app_sid"]
api_key = st.secrets["aspose"]["app_key"]
words_api = WordsApi(api_sid, api_key)

# --- Config ---
st.set_page_config("Intern Offer Generator", layout="wide")
EMAIL = st.secrets["email"]["user"]
PASSWORD = st.secrets["email"]["password"]
ADMIN_KEY = st.secrets["admin"]["key"]
CSV_FILE = "intern_offers.csv"
TEMPLATE_FILE = os.path.join(tempfile.gettempdir(), "offer_template.docx")
LOGO = "logo.png"

# --- Write template from base64 ---
if not os.path.exists(TEMPLATE_FILE):
    encoded_template = st.secrets["template_base64"]["template_base64"]
    with open(TEMPLATE_FILE, "wb") as f:
        f.write(base64.b64decode(encoded_template))

# --- Style ---
st.markdown("""
<style>
    .title-text {
        font-size: 2rem;
        font-weight: 700;
    }
    .stButton>button {
        background-color: #1E88E5;
        color: white;
        padding: 0.5rem 1.5rem;
        border-radius: 8px;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# --- Header ---
with st.container():
    col_logo, col_title = st.columns([1, 6])
    with col_logo:
        if os.path.exists(LOGO):
            st.image(LOGO, width=80)
    with col_title:
        st.markdown('<div class="title-text">SkyHighes Technologies Internship Letter Portal</div>', unsafe_allow_html=True)

st.divider()

# --- Utils ---
def format_date(d):
    return d.strftime("%A, %d %B %Y")

def generate_certificate_key():
    return ''.join(random.choices(string.ascii_uppercase + string.digits, k=9))

def generate_qr(data):
    qr = qrcode.QRCode(box_size=10, border=4)
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    path = os.path.join(tempfile.gettempdir(), "qr.png")
    img.save(path)
    return path

def save_to_csv(data):
    exists = os.path.exists(CSV_FILE)
    with open(CSV_FILE, mode='a', newline='') as f:
        writer = csv.writer(f)
        if not exists:
            writer.writerow(data.keys())
        writer.writerow(data.values())

def send_email(receiver, pdf_path, data, filename):
    msg = MIMEMultipart()
    msg['From'] = EMAIL
    msg['To'] = receiver
    msg['Subject'] = f"🎉 Your Internship Offer - {data['intern_name']}"

    html = f"""
    <html><body>
    <p>Dear {data['intern_name']},</p>
    <p>We are pleased to offer you an internship at <strong>SkyHighes Technology</strong>.</p>
    <ul>
        <li><b>Domain:</b> {data['domain']}</li>
        <li><b>Start Date:</b> {data['start_date']}</li>
        <li><b>End Date:</b> {data['end_date']}</li>
        <li><b>Offer Date:</b> {data['offer_date']}</li>
    </ul>
    <p>Your offer letter is attached.</p>
    </body></html>
    """
    msg.attach(MIMEText(html, 'html'))

    with open(pdf_path, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={filename}")
        msg.attach(part)

    with SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(EMAIL, PASSWORD)
        server.send_message(msg)

# --- Form ---
with st.form("offer_form"):
    st.subheader("🎓 Internship Offer Letter Generator")

    col1, col2, col3 = st.columns(3)
    with col1:
        intern_name = st.text_input("Intern Name")
    with col2:
        domain = st.text_input("Domain")
    with col3:
        email = st.text_input("Recipient Email")

    col4, col5, col6 = st.columns(3)
    with col4:
        start_date = st.date_input("Start Date", value=date.today())
    with col5:
        end_date = st.date_input("End Date", value=date.today())
    with col6:
        offer_date = st.date_input("Offer Date", value=date.today())

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
            "intern_name": intern_name.strip(),
            "domain": domain.strip(),
            "start_date": format_date(start_date),
            "end_date": format_date(end_date),
            "offer_date": format_date(offer_date),
            "i_id": intern_id,
            "email": email.strip()
        }

        save_to_csv(data)
        doc = DocxTemplate(TEMPLATE_FILE)
        doc.render(data)

        qr_path = generate_qr(f"{intern_name}, {domain}, {start_date}, {end_date}, {offer_date}, {intern_id}")
        try:
            doc.tables[0].rows[0].cells[2].paragraphs[0].add_run().add_picture(qr_path, width=Inches(1.5))
        except:
            st.warning("⚠️ QR insertion failed.")

        # ✅ Use file name like: Offer_Ananya Naresh Dube.pdf
        clean_name = intern_name.strip()
        docx_filename = f"Offer_{clean_name}.docx"
        pdf_filename = f"Offer_{clean_name}.pdf"

        docx_path = os.path.join(tempfile.gettempdir(), docx_filename)
        pdf_path = os.path.join(tempfile.gettempdir(), pdf_filename)

        doc.save(docx_path)

        try:
            # Upload to Aspose Cloud
            with open(docx_path, "rb") as f:
                words_api.upload_file(UploadFileRequest(f, docx_filename))

            # Convert to PDF
            save_opts = PdfSaveOptionsData(file_name=pdf_filename)
            save_as_request = SaveAsRequest(name=docx_filename, save_options_data=save_opts)
            words_api.save_as(save_as_request)

            # Download PDF
            result_stream = words_api.download_file(DownloadFileRequest(pdf_filename))
            with open(pdf_path, "wb") as f:
                f.write(result_stream)

            # Send via Email
            send_email(email, pdf_path, data, pdf_filename)

            # Success + Download
            st.success(f"✅ Offer Letter Sent to {email}")
            with open(pdf_path, "rb") as f:
                st.download_button("📥 Download Offer Letter", f, file_name=pdf_filename)

        except Exception as e:
            st.error(f"❌ Error occurred: {e}")

# --- Admin Panel ---
st.divider()
with st.expander("🔐 Admin Panel"):
    admin_key = st.text_input("Enter Admin Key", type="password")
    if admin_key == ADMIN_KEY:
        st.success("✅ Access granted.")
        if os.path.exists(CSV_FILE):
            with open(CSV_FILE, "r") as f:
                rows = list(csv.reader(f))
                if rows:
                    st.dataframe(rows[1:], columns=rows[0])
                    with open(CSV_FILE, "rb") as f_dl:
                        st.download_button("📥 Download CSV", f_dl, file_name="intern_offers.csv")
                else:
                    st.info("No data available.")
        else:
            st.info("CSV log not found.")
    elif admin_key:
        st.error("❌ Invalid key.")
