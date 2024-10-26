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
    Clase para enviar correos electrónicos con adjuntos para múltiples proveedores.
    """

    def __init__(self):
        """
        Inicializa una instancia de EmailSender.
        """
        self.config_manager = ConfigManager()
        self.user_email = self.config_manager.get_email()
        self.app_password = self.config_manager.get_app_password()

    def get_smtp_details(self) -> Tuple[str, int]:
        """
        Determina los detalles del servidor SMTP en función del dominio del correo del remitente.

        Returns:
            Tuple[str, int]: Servidor SMTP y puerto.
        """
        if 'gmail.com' in self.user_email:
            return 'smtp.gmail.com', 587
        elif 'outlook.com' in self.user_email or 'hotmail.com' in self.user_email or 'office365.com' in self.user_email:
            return 'smtp.office365.com', 587
        else:
            raise ValueError("Proveedor de correo no soportado.")

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
            # Obtener detalles del servidor SMTP
            smtp_server, smtp_port = self.get_smtp_details()

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
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(self.user_email, self.app_password)
                server.send_message(msg)

            return (True, "Correo enviado exitosamente.")
        except Exception as e:
            return (False, str(e))