# src/utils/helpers.py

import sys
import os
import logging
import requests
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from googleapiclient.discovery import build

logger = logging.getLogger(__name__)

def check_sharepoint_credentials(username, password):
    try:
        # URL base de tu sitio de SharePoint
        site_url = 'https://chcconsultorescr.sharepoint.com/sites/CONSULTORAS'

        ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))

        # Intentar obtener información del sitio para validar las credenciales
        web = ctx.web
        ctx.load(web)
        ctx.execute_query()

        # Si la consulta se ejecuta correctamente, las credenciales son válidas
        return True
    except Exception as e:
        logging.error(f"Error de autenticación en SharePoint: {e}")
        return False


def get_resource_path(relative_path):
    """Obtiene la ruta del recurso empaquetado o del entorno de desarrollo."""
    base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
    return os.path.join(base_path, relative_path)


def get_email_google(creds):
    try:
        service = build('people', 'v1', credentials=creds)
        profile = service.people().get(resourceName='people/me', personFields='emailAddresses').execute()
        email_addresses = profile.get('emailAddresses', [])
        if email_addresses:
            return email_addresses[0]['value']
        else:
            return None
    except Exception as e:
        logging.error(f"Error al obtener el correo electrónico: {str(e)}")
        return None


def get_microsoft_user_info(access_token):
    """
    Obtiene la información del usuario autenticado en Microsoft Graph.
    """
    try:
        headers = {'Authorization': f'Bearer {access_token}'}
        response = requests.get('https://graph.microsoft.com/v1.0/me', headers=headers)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        logging.error(f"Error al obtener información del usuario de Microsoft: {str(e)}")
        return {}
