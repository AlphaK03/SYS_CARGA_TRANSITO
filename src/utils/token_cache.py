# src/utils/token_cache.py
import msal
import os

class FileTokenCache:
    def __init__(self, cache_file='microsoft_token_cache.json'):
        self.cache_file = cache_file
        self.cache = msal.SerializableTokenCache()

        # Cargar el caché desde el archivo si existe
        if os.path.exists(self.cache_file):
            with open(self.cache_file, 'r') as f:
                try:
                    self.cache.deserialize(f.read())
                except Exception as e:
                    print(f"Error al deserializar el caché de tokens desde {self.cache_file}: {e}")

    def save(self, app):
        # Guardar el caché en el archivo después de actualizarlo en la memoria
        with open(self.cache_file, 'w') as f:
            f.write(app.token_cache.serialize())

    def clear_cache(self):
        """Limpia el caché de tokens y elimina el archivo de caché."""
        self.cache.deserialize("")  # Limpia el caché en memoria
        if os.path.exists(self.cache_file):
            os.remove(self.cache_file)  # Elimina el archivo de caché del disco

