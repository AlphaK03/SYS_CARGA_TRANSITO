# src/email_service/email_sender.py

import os
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from typing import Tuple
import logging
import requests
from msal import PublicClientApplication

from utils.config_manager import ConfigManager


class EmailSender:
    MS_SCOPES = [
        'User.Read',
        'Mail.Send'
    ]

    def __init__(self, msal_app: PublicClientApplication):
        """
        Inicializa una instancia de EmailSender con detalles de configuración de correo.
        """
        self.msal_app = msal_app
        self.config_manager = ConfigManager()
        self.user_email = self.config_manager.get_email()
        logging.info("EmailSender inicializado correctamente. (Solo Microsoft)")

    def get_microsoft_access_token(self):
        """
        Intenta adquirir un token de acceso para Microsoft (Graph) en modo silencioso.
        Si no está disponible en caché, devuelve None para requerir reautenticación.
        """
        accounts = self.msal_app.get_accounts()
        if accounts:
            logging.debug(f"Encontradas cuentas en MSAL: {accounts}")
            result = self.msal_app.acquire_token_silent(self.MS_SCOPES, account=accounts[0])
            if result and 'access_token' in result:
                logging.debug("Token adquirido silenciosamente.")
                return result['access_token']
            else:
                logging.info("No se pudo adquirir token silenciosamente. Requiere reautenticación.")
        else:
            logging.info("No hay cuentas en MSAL. Requiere autenticación.")
        return None

    def create_email_message(
            self, to_email: str, subject: str, html_content: str, attachment_path: str
    ) -> MIMEMultipart:
        """
        Crea un mensaje de correo electrónico con un archivo adjunto.
        """
        logging.debug(f"Creando el mensaje de correo para: {to_email}, Asunto: {subject}")
        msg = MIMEMultipart()
        msg['Subject'] = subject
        msg['From'] = self.user_email
        msg['To'] = to_email

        # Agregar el contenido HTML al correo
        msg.attach(MIMEText(html_content, 'html'))

        # Adjuntar el archivo, si existe
        if os.path.exists(attachment_path):
            logging.info(f"Adjuntando el archivo: {attachment_path}")
            with open(attachment_path, 'rb') as f:
                file_data = f.read()
                file_name = os.path.basename(attachment_path)
                attachment = MIMEApplication(file_data, Name=file_name)
                attachment['Content-Disposition'] = f'attachment; filename="{file_name}"'
                msg.attach(attachment)
        else:
            error_msg = f"El archivo adjunto '{attachment_path}' no existe."
            logging.error(error_msg)
            raise FileNotFoundError(error_msg)

        return msg

    def send_email(
            self,
            to_email: str,
            subject: str,
            html_content: str,
            attachment_path: str,
            cc_emails: str = "",
            bcc_emails: str = "",
            from_email: str = None  # <-- Nuevo parámetro
    ) -> Tuple[bool, str]:
        """
        Envía un correo utilizando Microsoft Graph.
        Retorna (True, "OK") en caso de éxito, o (False, "Mensaje de Error") en caso de falla.
        Se aceptan parámetros opcionales cc_emails y bcc_emails (como cadena con correos separados por comas).
        """
        try:
            return self._send_email_microsoft(to_email, subject, html_content, attachment_path, cc_emails, bcc_emails, from_email)
        except Exception as e:
            logging.error(f"Error al enviar el correo: {str(e)}")
            return False, f"Error al enviar el correo: {str(e)}"

    def _send_email_microsoft(
            self,
            to_email: str,
            subject: str,
            html_content: str,
            attachment_path: str,
            cc_emails: str = "",
            bcc_emails: str = "",
            from_email: str = None  # <-- Nuevo
    ) -> Tuple[bool, str]:
        """
        Envía un correo usando Microsoft Graph.
        Se procesan los campos to, cc y bcc de modo que, si contienen múltiples direcciones separadas por comas,
        se conviertan en listas separadas.
        """
        try:
            access_token = self.get_microsoft_access_token()
            if not access_token:
                logging.warning("No se encontró un access_token válido para Microsoft. Se requiere reautenticación.")
                return False, "Se requiere reautenticación con Microsoft."

            headers = {
                'Authorization': f'Bearer {access_token}',
                'Content-Type': 'application/json'
            }

            # Procesar el campo 'to_email'
            to_list = [email.strip() for email in to_email.split(',') if email.strip()]
            # Si no se pudo dividir (por ejemplo, si no hay comas) se usa la lista con el valor original.
            if not to_list:
                to_list = [to_email.strip()]

            # Construir la sección toRecipients
            message = {
                "message": {
                    "subject": subject,
                    "body": {
                        "contentType": "HTML",
                        "content": html_content
                    },
                    "toRecipients": [{"emailAddress": {"address": email}} for email in to_list]
                },
                "saveToSentItems": "true"
            }

            # 👇 Agregar el campo From si se especificó
            if from_email:
                message["message"]["from"] = {"emailAddress": {"address": from_email}}

            # Agregar ccRecipients si se proporcionan
            if cc_emails:
                cc_list = [email.strip() for email in cc_emails.split(',') if email.strip()]
                if cc_list:
                    message["message"]["ccRecipients"] = [{"emailAddress": {"address": email}} for email in cc_list]

            # Agregar bccRecipients si se proporcionan
            if bcc_emails:
                bcc_list = [email.strip() for email in bcc_emails.split(',') if email.strip()]
                if bcc_list:
                    message["message"]["bccRecipients"] = [{"emailAddress": {"address": email}} for email in bcc_list]

            # Adjuntar archivo si existe
            if os.path.exists(attachment_path):
                with open(attachment_path, 'rb') as f:
                    file_data = f.read()
                    file_name = os.path.basename(attachment_path)
                    encoded_content = base64.b64encode(file_data).decode()
                    attachment = {
                        "@odata.type": "#microsoft.graph.fileAttachment",
                        "Name": file_name,
                        "ContentBytes": encoded_content
                    }
                message["message"]["attachments"] = [attachment]
            else:
                logging.error(f"El archivo adjunto '{attachment_path}' no existe.")
                raise FileNotFoundError(f"El archivo adjunto '{attachment_path}' no existe.")

            # Enviar el correo a través de Microsoft Graph
            response = requests.post(
                'https://graph.microsoft.com/v1.0/me/sendMail',
                headers=headers,
                json=message
            )

            if response.status_code == 202:
                logging.info(f"Correo enviado exitosamente a {to_list} usando Microsoft Graph.")
                return True, "Correo enviado exitosamente."
            else:
                error_msg = f"Error al enviar correo con Microsoft: {response.text}"
                logging.error(error_msg)
                return False, error_msg

        except Exception as e:
            logging.error(f"Error al enviar correo con Microsoft: {str(e)}")
            return False, f"Error al enviar correo con Microsoft: {str(e)}"

