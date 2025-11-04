import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import getpass
# import json

# --- Configuration ---
# You'll need to set up an App Password for Gmail or enable less secure apps
# for other providers.
# Security best practice is NOT to hardcode your password here.

# Change this for other providers (e.g., 'smtp.office365.com' for Outlook)
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
SENDER_EMAIL = 'sunil.sangwal@gmail.com' 
# ---------------------


class Teacher:
    """Represents a teacher with their contact and timetable information."""
    def __init__(self, name: str, email: str, subject: str, timetable: str):
        self.name = name
        self.email = email
        self.subject = subject
        self.timetable = timetable

    def get_email_body(self) -> str:
        """Generates the personalized HTML email body."""
        return f"""\
        <html>
          <body>
            <p>Dear Mr./Ms. **{self.name}** ({self.subject} Teacher),</p>
            <p>Please find your updated teaching schedule for the coming week/term below:</p>
            <hr>
            <pre style="font-family: monospace; background-color: #f0f0f0; padding: 10px; border: 1px solid #ccc;">{self.timetable}</pre>
            <hr>
            <p>If you have any questions, please contact the administration office.</p>
            <p>Sincerely,<br>
               The Administration Team</p>
          </body>
        </html>
        """

class TimetableSender:
    """Handles the connection and logic for sending emails."""
    def __init__(self, smtp_server: str, smtp_port: int, sender_email: str):
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.sender_email = sender_email
        self.password = None
        self.smtp_session = None

    def get_credentials(self):
        """Securely prompts for the sender's email password."""
        print(f"Logging in as: {self.sender_email}")
        self.password = getpass.getpass("Enter your email password (or App Password): ")

    def connect(self):
        """Establishes and starts a secure SMTP connection."""
        try:
            self.smtp_session = smtplib.SMTP(self.smtp_server, self.smtp_port)
            self.smtp_session.ehlo()
            self.smtp_session.starttls()  # Upgrade to a secure connection
            self.smtp_session.ehlo()
            self.smtp_session.login(self.sender_email, self.password)
            print("SMTP connection and login successful.")
            return True
        except Exception as e:
            print(f"Error connecting or logging in: {e}")
            self.smtp_session = None
            return False

    def disconnect(self):
        """Closes the SMTP connection."""
        if self.smtp_session:
            self.smtp_session.quit()
            print("SMTP session closed.")

    def send_timetable(self, teacher: Teacher):
        """Constructs and sends a personalized timetable email to a single teacher."""
        if not self.smtp_session:
            print(f"Error: Not connected to SMTP server. Cannot send email to {teacher.email}")
            return

        try:
            # Create message container
            msg = MIMEMultipart('alternative')
            msg['Subject'] = f"Your New Teaching Timetable - {teacher.subject}"
            msg['From'] = self.sender_email
            msg['To'] = teacher.email

            # Create the body (HTML version)
            html_body = teacher.get_email_body()
            part = MIMEText(html_body, 'html')
            msg.attach(part)

            # Send the email
            self.smtp_session.sendmail(self.sender_email, teacher.email, msg.as_string())
            print(f"✅ Successfully sent timetable to **{teacher.name}** at {teacher.email}")

        except Exception as e:
            print(f"❌ Failed to send email to {teacher.name}: {e}")

# --- Example Usage ---

def run_sender():
    """Main function to run the timetable script."""
    # 1. DEFINE TEACHER DATA (Ideally, this would be loaded from a CSV or database)
    teachers_data = [
        Teacher(
            name="Sunil Sangwal",
            email="sunil.sangwal@gmail.com", # Replace with actual email
            subject="Mathematics",
            timetable="""
Mon: 8:00-9:00 (Grade 10), 9:00-10:00 (Grade 12)
Tue: 10:00-11:00 (Grade 9)
Wed: FREE
"""
        ),
        Teacher(
            name="Amit Arora",
            email="amit.arora@xgxmxaxixlx.com", # Replace with actual email
            subject="English",
            timetable="""
Mon: 11:00-12:00 (Grade 8)
Tue: 8:00-9:00 (Grade 11)
Fri: 13:00-14:00 (Grade 7)
"""
        )
    ]

    # 2. INITIALIZE SENDER
    sender = TimetableSender(SMTP_SERVER, SMTP_PORT, SENDER_EMAIL)
    sender.get_credentials()
    
    # 3. CONNECT AND SEND
    if sender.connect():
        for teacher in teachers_data:
            # In a real scenario, you might add a small delay here to avoid being flagged as spam
            sender.send_timetable(teacher)
        
        # 4. DISCONNECT
        sender.disconnect()
    else:
        print("Script aborted due to connection failure.")

if __name__ == "__main__":
    run_sender()
