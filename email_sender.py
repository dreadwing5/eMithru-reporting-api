import smtplib
import os
# from email.message import EmailMessage
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.text import MIMEText
import ssl


class EmailSender:
    def __init__(self, email: str, password: str, subject: str, body: str, recipients: list, attachment: str):
        self.email = email
        self.password = password
        self.subject = subject
        self.body = body
        self.recipients = recipients
        self.attachment = attachment

    def send_email(self):
        msg = MIMEMultipart()
        msg['From'] = self.email
        msg['To'] = ', '.join(self.recipients)
        msg['Subject'] = self.subject
        msg.attach(MIMEText(self.body, 'plain'))
        if self.attachment:
            with open(self.attachment, 'rb') as attachment_file:
                attachment_payload = MIMEBase('application', 'octet-stream')

                attachment_payload.set_payload(attachment_file.read())
            encoders.encode_base64(attachment_payload)
            attachment_payload.add_header(
                'Content-Disposition',
                f'attachment; filename={os.path.basename(self.attachment)}',
            )
            msg.attach(attachment_payload)

        context = ssl.create_default_context()

        with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as server:
            server.login(self.email, self.password)
            server.sendmail(self.email, self.recipients, msg.as_string())

        print("Email sent successfully.")
