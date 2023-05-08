import os
import smtplib
from typing import List
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv
load_dotenv()


class EmailController:
    def __init__(self, smtp_server="smtp.gmail.com", smtp_port=587, smtp_username=os.getenv("MAIL_ID"), smtp_password=os.getenv("MAIL_PASS")):
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.smtp_username = smtp_username
        self.smtp_password = smtp_password

    def send_email(self, email_template):
        message = email_template.create_message()
        self._send_email(message, email_template.hod_name)

    def _send_email(self, message, recipient):
        try:
            smtp_server = smtplib.SMTP(self.smtp_server, self.smtp_port)
            smtp_server.starttls()
            smtp_server.login(self.smtp_username, self.smtp_password)
            message['From'] = self.smtp_username
            message['To'] = recipient

            smtp_server.sendmail(self.smtp_username,
                                 recipient, message.as_string())

            smtp_server.quit()
            print("Email sent successfully!")
        except Exception as e:
            print("Failed to send email.")
            print(e)
