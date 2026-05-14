import os
import smtplib
from email.message import EmailMessage
from dotenv import load_dotenv
from pathlib import Path

# 1. Load variables from .env
env_path = Path('.') / '.env'
load_dotenv(dotenv_path=env_path)

EMAIL_ADDRESS = os.getenv("EMAIL_USER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

print(EMAIL_PASSWORD)

# 2. Define the recipient and content
recipient = "sangwal77@gmail.com"

msg = EmailMessage()
msg['Subject'] = "Secure Email via Python"
msg['From'] = EMAIL_ADDRESS
msg['To'] = recipient
msg.set_content("This script now uses environment variables for better security!")

# 3. Execute the sending process
try:
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        smtp.send_message(msg)
    print(f"Success! Email sent to {recipient}")
except Exception as e:
    print(f"Failed to send email. Error: {e}")
