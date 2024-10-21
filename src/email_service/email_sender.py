# src/email_service/email_sender.py
import os
import smtplib
from email.message import EmailMessage

class EmailSender:
    def __init__(self, user_email, app_password):
        self.user_email = user_email
        self.app_password = app_password

    def send_email(self, recipient, subject, body_html, attachment_path=None):
        try:
            msg = EmailMessage()
            msg['From'] = self.user_email
            msg['To'] = recipient
            msg['Subject'] = subject
            msg.set_content(body_html, subtype='html')

            if attachment_path:
                with open(attachment_path, 'rb') as f:
                    file_data = f.read()
                    file_name = os.path.basename(attachment_path)
                    msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(self.user_email, self.app_password)
                smtp.send_message(msg)
            return True, f"El archivo '{os.path.basename(attachment_path)}' se envió correctamente a {recipient}."
        except Exception as e:
            return False, str(e)
