# src/utils/config_manager.py
import os
import json
import keyring
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request

def credentials_to_dict(creds):
    return {
        'token': creds.token,
        'refresh_token': creds.refresh_token,
        'token_uri': creds.token_uri,
        'client_id': creds.client_id,
        'client_secret': creds.client_secret,
        'scopes': creds.scopes
    }

class ConfigManager:
    """
    Clase para manejar la configuración de la aplicación.
    """
    CONFIG_FILE = 'data/config.json'

    def __init__(self):
        # Crear el directorio si no existe
        config_dir = os.path.dirname(self.CONFIG_FILE)
        if not os.path.exists(config_dir):
            os.makedirs(config_dir)
        # Cargar la configuración existente o crear una nueva
        if os.path.exists(self.CONFIG_FILE):
            try:
                with open(self.CONFIG_FILE, 'r') as f:
                    self.config = json.load(f)
            except json.JSONDecodeError:
                print("Error: el archivo de configuración está corrupto o vacío. Creando uno nuevo.")
                self.config = self.create_default_config()
                self.save_config()
        else:
            self.config = self.create_default_config()
            self.save_config()

    def create_default_config(self):
        return {
            "email": "",
            "email_provider": "",
            "google_credentials": None,
            "microsoft_credentials": None,
            "sharepoint_site_url": "",
            "sharepoint_username": "",
            "sharepoint_file_url": ""
        }

    # Métodos para manejar el proveedor de correo y el email
    def get_email_provider(self) -> str:
        return self.config.get('email_provider', '')

    def set_email_provider(self, provider: str):
        self.config['email_provider'] = provider
        self.save_config()

    def get_email(self) -> str:
        return self.config.get('email', '')

    def set_email(self, email: str):
        # Guarda el correo electrónico en la configuración
        self.config['email'] = email
        self.save_config()

    # Métodos para manejar las contraseñas usando keyring
    def get_app_password(self) -> str:
        return keyring.get_password('ProyectoASW_AppPassword', self.get_email()) or ''

    def set_app_password(self, password: str):
        keyring.set_password('ProyectoASW_AppPassword', self.get_email(), password)

    def get_sharepoint_password(self) -> str:
        return keyring.get_password('ProyectoASW_SharePointPassword', self.get_sharepoint_username()) or ''

    def set_sharepoint_password(self, password: str):
        keyring.set_password('ProyectoASW_SharePointPassword', self.get_sharepoint_username(), password)

    # Métodos para SharePoint
    def get_sharepoint_site_url(self) -> str:
        return self.config.get('sharepoint_site_url', '')

    def set_sharepoint_site_url(self, site_url: str):
        self.config['sharepoint_site_url'] = site_url
        self.save_config()

    def get_sharepoint_username(self) -> str:
        return self.config.get('sharepoint_username', '')

    def set_sharepoint_username(self, username: str):
        self.config['sharepoint_username'] = username
        self.save_config()

    def get_sharepoint_file_url(self) -> str:
        return self.config.get('sharepoint_file_url', '')

    def set_sharepoint_file_url(self, file_url: str):
        self.config['sharepoint_file_url'] = file_url
        self.save_config()

    def save_config(self):
        """
        Guarda la configuración en el archivo JSON.
        """
        with open(self.CONFIG_FILE, 'w') as f:
            json.dump(self.config, f, indent=4)

    # Métodos para credenciales de Google OAuth2
    def set_google_credentials(self, creds):
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        self.config['google_credentials'] = credentials_to_dict(creds)
        self.save_config()

    def get_google_credentials(self):
        creds_dict = self.config.get('google_credentials')
        if creds_dict:
            creds = Credentials(
                token=creds_dict.get('token'),
                refresh_token=creds_dict.get('refresh_token'),
                token_uri=creds_dict.get('token_uri'),
                client_id=creds_dict.get('client_id'),
                client_secret=creds_dict.get('client_secret'),
                scopes=creds_dict.get('scopes')
            )
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
                self.set_google_credentials(creds)  # Actualizar el token si se ha refrescado
            return creds
        return None


    def clear_credentials(self):
        # Implementa la lógica para limpiar todas las credenciales almacenadas
        self.config.pop('microsoft_access_token', None)
        self.config.pop('email', None)
        self.save_config()

