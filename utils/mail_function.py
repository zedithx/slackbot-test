import logging
import os
import smtplib
import time

import pytz
import schedule
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv

# Configure logging
logging.basicConfig(level=logging.INFO)  # Set to DEBUG for more details
logger = logging.getLogger(__name__)

# Load environment variables from .env file
load_dotenv()

# Email configuration
EMAIL_ADDRESS = os.getenv("EMAIL_ADDRESS")  # Replace with your email
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")          # Replace with your email password
RECIPIENT_EMAIL = os.getenv("RECIPIENT_EMAIL")   # Replace with recipient's email
SMTP_SERVER = "smtp.gmail.com"            # SMTP server for Gmail
SMTP_PORT = 587                           # Port for TLS

# Singapore timezone
SGT = pytz.timezone("Asia/Singapore")

def send_email_with_excel():
    try:
        current_date = datetime.now().strftime("%d-%m_CIT")
        filepath = f"checkin_records/{current_date}.xlsx"
        # Create the email
        subject = "End of Day Check-in Report"
        body = "Please find attached the check-in report for today."

        msg = MIMEMultipart()
        msg["From"] = EMAIL_ADDRESS
        msg["To"] = RECIPIENT_EMAIL
        msg["Subject"] = subject

        msg.attach(MIMEText(body, "plain"))

        # Attach the Excel file
        with open(filepath, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f"attachment; filename={filepath}",
        )
        msg.attach(part)

        # Connect to the email server and send the email
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()  # Secure the connection
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.sendmail(EMAIL_ADDRESS, RECIPIENT_EMAIL, msg.as_string())

        logger.info("Email sent successfully!")
    except Exception as e:
        logger.info(f"Error sending email: {e}")

# Function to send the email at 11:59 PM SGT
def schedule_email():
    now = datetime.now(SGT)
    logger.info(f"Scheduler running. Current SGT time: {now.strftime('%Y-%m-%d %H:%M:%S')}")

    # Schedule the task at 23:59 SGT
    schedule.every().day.at("12:00").do(send_email_with_excel)

    while True:
        schedule.run_pending()
        time.sleep(60)  # Check every 60 seconds
