import os
import base64
from email.message import EmailMessage
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Scope for sending emails
SCOPES = ['https://www.googleapis.com/auth/gmail.send']

def get_gmail_service():
    creds = None
    # Load existing tokens if they exist
    if os.path.exists('email_sender/token.json'):
        creds = Credentials.from_authorized_user_file('email_sender/token.json', SCOPES)

    # If no valid tokens, prompt the user to "Sign in with Google"
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'email_sender/credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('email_sender/token.json', 'w') as token:
            token.write(creds.to_json())

    return build('gmail', 'v1', credentials=creds)

def send_message(to_email, subject, body):
    try:
        service = get_gmail_service()

        # Create the email structure
        message = EmailMessage()
        message.set_content(body)
        message['To'] = to_email
        message['From'] = 'me' # Google automatically uses your authenticated email
        message['Subject'] = subject

        # The Gmail API requires the message to be base64url encoded
        encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

        create_message = {
            'raw': encoded_message
        }

        # Send the email
        send_receipt = service.users().messages().send(userId="me", body=create_message).execute()
        print(f'Message sent successfully! Message Id: {send_receipt["id"]}')

    except HttpError as error:
        print(f'An error occurred: {error}')

if __name__ == '__main__':
    send_message(
        to_email="sunil.sangwal@gmail.com",
        subject="Hello from Python OAuth!",
        body="This email was sent using the Gmail API and 'Sign in with Google' authentication."
    )
