# src/email_service/email_sender.py
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import os
from typing import Tuple

from utils.config_manager import ConfigManager

class EmailSender:
    """
    Clase para enviar correos electrónicos con adjuntos.
    """

    def __init__(self):
        """
        Inicializa una instancia de EmailSender.
        """
        self.config_manager = ConfigManager()
        self.user_email = self.config_manager.get_email()
        self.app_password = self.config_manager.get_app_password()
        self.smtp_server = 'smtp.gmail.com'
        self.smtp_port = 587

    def send_email(self, to_email: str, subject: str, html_content: str, attachment_path: str) -> Tuple[bool, str]:
        """
        Envía un correo electrónico con un archivo adjunto.

        Args:
            to_email (str): Dirección de correo del destinatario.
            subject (str): Asunto del correo.
            html_content (str): Contenido HTML del correo.
            attachment_path (str): Ruta al archivo adjunto.

        Returns:
            Tuple[bool, str]: (Éxito, Mensaje de error si falla).
        """
        try:
            # Crear un mensaje de tipo "multipart"
            msg = MIMEMultipart()
            msg['Subject'] = subject
            msg['From'] = self.user_email
            msg['To'] = to_email

            # Agregar el contenido HTML al correo
            msg.attach(MIMEText(html_content, 'html'))

            # Adjuntar el archivo
            with open(attachment_path, 'rb') as f:
                file_data = f.read()
                file_name = os.path.basename(attachment_path)
                attachment = MIMEApplication(file_data, Name=file_name)
                attachment['Content-Disposition'] = f'attachment; filename="{file_name}"'
                msg.attach(attachment)

            # Enviar el correo
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()
                server.login(self.user_email, self.app_password)
                server.send_message(msg)

            return (True, "Correo enviado exitosamente.")
        except Exception as e:
            return (False, str(e))
