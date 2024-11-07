# src/email_service/email_sender.py

import os
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from typing import Tuple
import logging
from googleapiclient.discovery import build
import requests
from msal import PublicClientApplication

from utils.config_manager import ConfigManager

class EmailSender:
    GOOGLE_SCOPES = [
        'https://www.googleapis.com/auth/gmail.send',
        'https://www.googleapis.com/auth/userinfo.email',
        'https://www.googleapis.com/auth/userinfo.profile',
        'openid'
    ]
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
        self.email_provider = self.config_manager.get_email_provider()
        self.google_creds = self.config_manager.get_google_credentials()
        logging.info("EmailSender inicializado correctamente.")

    def get_microsoft_access_token(self):
        accounts = self.msal_app.get_accounts()
        if accounts:
            logging.debug(f"Encontrados accounts en MSAL: {accounts}")
            result = self.msal_app.acquire_token_silent(self.MS_SCOPES, account=accounts[0])
            if result and 'access_token' in result:
                logging.debug("Token adquirido silenciosamente.")
                return result['access_token']
            else:
                logging.info("No se pudo adquirir token silenciosamente. Requiere reautenticación.")
        else:
            logging.info("No hay cuentas en MSAL. Requiere autenticación.")
        return None

    def create_email_message(self, to_email: str, subject: str, html_content: str,
                             attachment_path: str) -> MIMEMultipart:
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

    def send_email(self, to_email: str, subject: str, html_content: str, attachment_path: str) -> Tuple[bool, str]:
        try:
            if self.email_provider == 'Google' and self.google_creds:
                return self._send_email_google(to_email, subject, html_content, attachment_path)
            elif self.email_provider == 'Microsoft':
                return self._send_email_microsoft(to_email, subject, html_content, attachment_path)
            else:
                raise ValueError("Credenciales OAuth no disponibles o proveedor de correo no soportado.")
        except Exception as e:
            logging.error(f"Error al enviar el correo: {str(e)}")
            return False, f"Error al enviar el correo: {str(e)}"

    def _send_email_google(self, to_email: str, subject: str, html_content: str, attachment_path: str) -> Tuple[bool, str]:
        """
        Envía un correo usando el API de Gmail.
        """
        try:
            service = build('gmail', 'v1', credentials=self.google_creds)
            message = self.create_email_message(to_email, subject, html_content, attachment_path)
            raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

            send_message = {
                'raw': raw_message
            }
            service.users().messages().send(userId='me', body=send_message).execute()
            logging.info(f"Correo enviado exitosamente a {to_email} usando Gmail API.")
            return (True, "Correo enviado exitosamente.")
        except Exception as e:
            logging.error(f"Error al enviar correo con Google: {str(e)}")
            return (False, f"Error al enviar correo con Google: {str(e)}")

    def _send_email_microsoft(self, to_email: str, subject: str, html_content: str, attachment_path: str) -> Tuple[bool, str]:
        try:
            access_token = self.get_microsoft_access_token()
            if not access_token:
                # No se encontró un access_token válido en la caché
                logging.warning("No se encontró un access_token válido para Microsoft. Se requiere reautenticación.")
                return False, "Se requiere reautenticación con Microsoft."

            headers = {
                'Authorization': f'Bearer {access_token}',
                'Content-Type': 'application/json'
            }
            message = {
                "message": {
                    "subject": subject,
                    "body": {
                        "contentType": "HTML",
                        "content": html_content
                    },
                    "toRecipients": [
                        {"emailAddress": {"address": to_email}}
                    ]
                },
                "saveToSentItems": "true"
            }

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

            # Usar el punto de acceso de Microsoft Graph para enviar el correo
            response = requests.post(
                'https://graph.microsoft.com/v1.0/me/sendMail',
                headers=headers,
                json=message
            )

            if response.status_code == 202:
                logging.info(f"Correo enviado exitosamente a {to_email} usando Microsoft Graph.")
                return True, "Correo enviado exitosamente."
            else:
                error_msg = f"Error al enviar correo con Microsoft: {response.text}"
                logging.error(error_msg)
                return False, error_msg
        except Exception as e:
            logging.error(f"Error al enviar correo con Microsoft: {str(e)}")
            return False, f"Error al enviar correo con Microsoft: {str(e)}"
