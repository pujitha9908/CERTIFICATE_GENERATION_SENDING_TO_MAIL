import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import os
import smtplib
import ssl
from email.message import EmailMessage

# =====================
# CONFIG
# =====================
EXCEL_FILE = "students.xlsx"
TEMPLATE = "certificate_.png"
OUTPUT_DIR = "certificates"

# Gmail credentials (use App Password, not your real password!)
FROM_EMAIL = "ponugotipujitha918@gmail.com"
APP_PASSWORD = "oeja tsoi dswy zgik"  # generate from Google Account > Security > App Passwords

# Create output folder if missing
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)

# Load Excel
students = pd.read_excel(EXCEL_FILE)

# Function to generate certificate PDF
def generate_certificate(name, roll, dept):
    file_path = os.path.join(OUTPUT_DIR, f"{roll}_{name}.pdf")
    c = canvas.Canvas(file_path, pagesize=A4)

    # Draw template background
    if os.path.exists(TEMPLATE):
        c.drawImage(TEMPLATE, 0, 0, width=A4[0], height=A4[1])

    c.setFont("Helvetica-Bold", 28)
    c.drawCentredString(A4[0]/2, A4[1]/2 + 50, name)

    c.setFont("Helvetica", 16)
    c.drawCentredString(A4[0]/2, A4[1]/2, f"Roll No: {roll} | Dept: {dept}")

    c.save()
    return file_path

# Function to send email using Gmail + App Password
def send_email(to_email, subject, body, attachment_path):
    msg = EmailMessage()
    msg["From"] = FROM_EMAIL
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.set_content(body, subtype="html")

    # Attach PDF
    with open(attachment_path, "rb") as f:
        file_data = f.read()
        file_name = os.path.basename(attachment_path)
    msg.add_attachment(file_data, maintype="application", subtype="pdf", filename=file_name)

    # Send email via Gmail SMTP
    try:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
            server.login(FROM_EMAIL, APP_PASSWORD)
            server.send_message(msg)
        print(f"✅ Sent to {to_email}")
    except Exception as e:
        print(f"❌ Failed to send to {to_email}: {str(e)}")

# =====================
# Process students
# =====================
for _, row in students.iterrows():
    name = row["Name"]
    roll = row["Roll number"]
    dept = row["Department"]
    email = row["Email"]

    pdf_path = generate_certificate(name, roll, dept)

    subject = "Certificate of Participation"
    body = f"""
    <p>Dear {name},</p>
    <p>Please find attached your certificate of participation.</p>
    <p>Best regards,<br>Team</p>
    """

    send_email(email, subject, body, pdf_path)