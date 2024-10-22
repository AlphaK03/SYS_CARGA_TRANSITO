# src/utils/config_manager.py
import os
import json
import keyring

class ConfigManager:
    """
    Clase para manejar la configuración de la aplicación.
    """
    CONFIG_FILE = 'data/config.json'
    SERVICE_NAME = 'EmailSenderApp'  # Nombre del servicio para keyring

    def __init__(self):
        # Crear el directorio si no existe
        config_dir = os.path.dirname(self.CONFIG_FILE)
        if not os.path.exists(config_dir):
            os.makedirs(config_dir)
        # Cargar la configuración existente o crear una nueva
        if os.path.exists(self.CONFIG_FILE):
            with open(self.CONFIG_FILE, 'r') as f:
                self.config = json.load(f)
        else:
            self.config = {}

    def get_email(self) -> str:
        """
        Obtiene el correo electrónico almacenado.

        Returns:
            str: Correo electrónico o cadena vacía si no está configurado.
        """
        return self.config.get('email', '')

    def set_email(self, email: str):
        """
        Establece el correo electrónico y guarda la configuración.

        Args:
            email (str): Correo electrónico.
        """
        self.config['email'] = email
        self.save_config()

    def get_app_password(self) -> str:
        """
        Obtiene la contraseña de aplicación almacenada en keyring.

        Returns:
            str: Contraseña de aplicación o None si no está configurada.
        """
        return keyring.get_password(self.SERVICE_NAME, self.get_email())

    def set_app_password(self, password: str):
        """
        Establece la contraseña de aplicación en keyring.

        Args:
            password (str): Contraseña de aplicación.
        """
        keyring.set_password(self.SERVICE_NAME, self.get_email(), password)

    def save_config(self):
        """
        Guarda la configuración en el archivo JSON.
        """
        with open(self.CONFIG_FILE, 'w') as f:
            json.dump(self.config, f)
