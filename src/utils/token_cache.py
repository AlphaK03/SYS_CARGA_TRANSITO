# src/utils/token_cache.py
import msal
import os
import json

class FileTokenCache:
    def __init__(self, cache_file='microsoft_token_cache.json'):
        self.cache_file = cache_file
        self.cache = msal.SerializableTokenCache()
        if os.path.exists(self.cache_file):
            with open(self.cache_file, 'r') as f:
                try:
                    self.cache.deserialize(f.read())
                except Exception as e:
                    print(f"Error al deserializar el caché de tokens: {e}")

    def load(self, app):
        app.token_cache.deserialize(self.cache.serialize())

    def save(self, app):
        with open(self.cache_file, 'w') as f:
            f.write(app.token_cache.serialize())

    def clear_cache(self):
        """Clear the token cache and delete the cache file."""
        self.cache.deserialize("")  # Clear the in-memory cache
        if os.path.exists(self.cache_file):
            os.remove(self.cache_file)  # Delete the cache file from disk