import msal
import keyring
import logging
import keyring.errors


class SecureTokenCache:
    def __init__(self, cache_name='default_token_cache', fragment_size=1000):
        """
        Inicializa el caché de tokens utilizando keyring.

        :param cache_name: Nombre único para identificar el caché en keyring.
        :param fragment_size: Tamaño de cada fragmento de datos.
        """
        self.cache_name = cache_name
        self.cache = msal.SerializableTokenCache()
        self.fragment_size = fragment_size
        self.load()
        logging.debug(
            f"SecureTokenCache inicializado con cache_name='{self.cache_name}' y fragment_size={self.fragment_size}.")

    def load(self):
        """
        Carga el caché de tokens desde keyring y reensambla los fragmentos.
        """
        logging.debug("Iniciando la carga del caché de tokens desde keyring.")
        serialized_cache = ""
        fragment_index = 0

        while True:
            fragment_key = f"{self.cache_name}_{fragment_index}"
            try:
                fragment = keyring.get_password('my_app', fragment_key)
                if not fragment:
                    logging.debug(f"No se encontró el fragmento {fragment_index} en keyring. Finalizando carga.")
                    break
                serialized_cache += fragment
                logging.debug(f"Fragmento {fragment_index} cargado desde keyring.")
                fragment_index += 1
            except keyring.errors.KeyringError as e:
                logging.error(f"Error al obtener el fragmento {fragment_index} de keyring: {e}")
                break

        if serialized_cache:
            try:
                self.cache.deserialize(serialized_cache)
                logging.info(f"Caché de tokens cargado exitosamente desde keyring para '{self.cache_name}'.")
            except Exception as e:  # Excepción general en lugar de SerializationError
                logging.error(f"Error al deserializar el caché de tokens desde keyring: {e}")
        else:
            logging.info(f"No se encontró caché de tokens existente para '{self.cache_name}' en keyring.")

    def save(self, app):
        """
        Guarda el caché de tokens en keyring, dividido en fragmentos si es necesario.

        :param app: Instancia de MSAL PublicClientApplication.
        """
        logging.debug("Iniciando el guardado del caché de tokens en keyring.")
        serialized_cache = self.cache.serialize()
        num_fragments = len(serialized_cache) // self.fragment_size + (
            1 if len(serialized_cache) % self.fragment_size else 0)
        logging.debug(
            f"Serialized cache length: {len(serialized_cache)} bytes, dividido en {num_fragments} fragmentos.")

        # Eliminar fragmentos existentes que ya no sean necesarios
        existing_fragments = []
        fragment_index = 0
        while True:
            fragment_key = f"{self.cache_name}_{fragment_index}"
            try:
                fragment = keyring.get_password('my_app', fragment_key)
                if fragment is None:
                    break
                existing_fragments.append(fragment_index)
                fragment_index += 1
            except keyring.errors.KeyringError as e:
                logging.error(f"Error al verificar fragmento {fragment_index} en keyring: {e}")
                break

        # Eliminar fragmentos sobrantes
        for i in range(num_fragments, fragment_index):
            fragment_key = f"{self.cache_name}_{i}"
            try:
                keyring.delete_password('my_app', fragment_key)
                logging.debug(f"Fragmento sobrante {i} eliminado de keyring.")
            except keyring.errors.PasswordDeleteError:
                logging.warning(f"No se pudo eliminar el fragmento {i} de keyring. Puede que no exista.")
            except keyring.errors.KeyringError as e:
                logging.error(f"Error al eliminar fragmento {i} de keyring: {e}")

        # Guardar nuevos fragmentos
        for i in range(num_fragments):
            fragment = serialized_cache[i * self.fragment_size:(i + 1) * self.fragment_size]
            fragment_key = f"{self.cache_name}_{i}"
            try:
                keyring.set_password('my_app', fragment_key, fragment)
                logging.debug(f"Fragmento {i} guardado en keyring bajo la clave '{fragment_key}'.")
            except keyring.errors.KeyringError as e:
                logging.error(f"Error al guardar el fragmento {i} en keyring: {e}")

        logging.info(f"Caché de tokens guardado exitosamente en keyring para '{self.cache_name}'.")

    def clear_cache(self):
        """
        Limpia el caché de tokens en keyring y en memoria.
        """
        logging.debug("Iniciando la limpieza del caché de tokens.")
        # Limpiar el caché en memoria
        try:
            self.cache.deserialize("")  # Limpia el caché en memoria
            logging.debug("Caché en memoria deserializado y limpiado.")
        except Exception as e:  # Excepción general para deserialización
            logging.error(f"Error al limpiar el caché en memoria: {e}")

        # Eliminar todos los fragmentos almacenados en keyring
        fragment_index = 0
        while True:
            fragment_key = f"{self.cache_name}_{fragment_index}"
            try:
                fragment = keyring.get_password('my_app', fragment_key)
                if not fragment:
                    logging.debug(f"No se encontró el fragmento {fragment_index} en keyring durante la limpieza.")
                    break
                keyring.delete_password('my_app', fragment_key)
                logging.debug(f"Fragmento {fragment_index} eliminado de keyring.")
                fragment_index += 1
            except keyring.errors.PasswordDeleteError:
                logging.warning(f"No se pudo eliminar el fragmento {fragment_index} de keyring. Puede que no exista.")
                fragment_index += 1
            except keyring.errors.KeyringError as e:
                logging.error(f"Error al eliminar el fragmento {fragment_index} de keyring: {e}")
                break

        logging.info(f"Caché de tokens eliminado exitosamente de keyring para '{self.cache_name}'.")
