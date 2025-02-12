# src/utils/helpers.py

import sys
import os
import logging
import requests

logger = logging.getLogger(__name__)

def get_resource_path(relative_path):
    """Obtiene la ruta del recurso empaquetado o del entorno de desarrollo."""
    base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
    return os.path.join(base_path, relative_path)



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
