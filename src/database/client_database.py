import os
import sqlite3
from typing import List, Optional, Tuple

class ClientDatabase:
    """
    Clase para manejar la base de datos de clientes y correos electrónicos.
    """

    def __init__(self, db_path: str = 'data/clientes.db'):
        """
        Inicializa la conexión a la base de datos.

        Args:
            db_path (str): Ruta al archivo de la base de datos SQLite.
        """
        self.db_path = os.path.abspath(db_path)
        db_dir = os.path.dirname(self.db_path)

        # Crear el directorio si no existe
        if not os.path.exists(db_dir):
            os.makedirs(db_dir)

        # Establecer la conexión a la base de datos
        try:
            self.connection = sqlite3.connect(self.db_path)
            self.create_tables()
            self.initialize_default_list()
        except sqlite3.OperationalError as e:
            raise Exception(f"No se pudo abrir la base de datos: {e}")

    def create_tables(self):
        cursor = self.connection.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS clientes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT UNIQUE NOT NULL,
                correo TEXT DEFAULT '(correo sin definir)',
                correo_cc TEXT DEFAULT '',
                correo_bcc TEXT DEFAULT ''
            )
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS aliases (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                client_id INTEGER NOT NULL,
                alias TEXT NOT NULL,
                FOREIGN KEY(client_id) REFERENCES clientes(id)
            )
        ''')
        # Nueva tabla 'listas'
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS listas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT UNIQUE NOT NULL
            )
        ''')
        # Nueva tabla 'client_list' para la relación muchos a muchos
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS client_list (
                client_id INTEGER NOT NULL,
                list_id INTEGER NOT NULL,
                FOREIGN KEY(client_id) REFERENCES clientes(id),
                FOREIGN KEY(list_id) REFERENCES listas(id),
                PRIMARY KEY (client_id, list_id)
            )
        ''')
        self.connection.commit()

    # Métodos para manejar clientes y aliases

    def add_client(self, nombre: str, correo: str = '(correo sin definir)',
                   correo_cc: str = '', correo_bcc: str = '', client_id: Optional[int] = None):
        """
        Agrega un nuevo cliente a la base de datos.
        Si se proporciona client_id, lo utiliza; de lo contrario, se auto-incrementa.

        Args:
            nombre (str): Nombre del cliente.
            correo (str): Correo electrónico principal del cliente.
            correo_cc (str): Correos para CC, separados por comas.
            correo_bcc (str): Correos para CCO (BCC), separados por comas.
            client_id (Optional[int]): ID del cliente. Si se proporciona, se usa para la inserción.
        """
        cursor = self.connection.cursor()
        if client_id is not None:
            try:
                cursor.execute('''
                    INSERT INTO clientes (id, nombre, correo, correo_cc, correo_bcc)
                    VALUES (?, ?, ?, ?, ?)
                ''', (client_id, nombre, correo, correo_cc, correo_bcc))
            except sqlite3.IntegrityError as e:
                raise Exception(f"No se pudo agregar el cliente con ID {client_id}: {e}")
        else:
            cursor.execute('''
                INSERT INTO clientes (nombre, correo, correo_cc, correo_bcc)
                VALUES (?, ?, ?, ?)
            ''', (nombre, correo, correo_cc, correo_bcc))
        self.connection.commit()

    def update_client_name_and_email(self, client_id: int, nombre: str, correo: str,
                                       correo_cc: str = '', correo_bcc: str = ''):
        """
        Actualiza el nombre, el correo principal, el correo CC y el correo CCO de un cliente.

        Args:
            client_id (int): ID del cliente.
            nombre (str): Nuevo nombre del cliente.
            correo (str): Nuevo correo principal del cliente.
            correo_cc (str): Nuevo valor para el campo de CC (correos separados por comas).
            correo_bcc (str): Nuevo valor para el campo de CCO (correos separados por comas).
        """
        cursor = self.connection.cursor()
        cursor.execute('''
            UPDATE clientes
            SET nombre = ?, correo = ?, correo_cc = ?, correo_bcc = ?
            WHERE id = ?
        ''', (nombre, correo, correo_cc, correo_bcc, client_id))
        self.connection.commit()

    def delete_client(self, client_id: int):
        """
        Elimina un cliente, sus aliases y sus asociaciones en listas.

        Args:
            client_id (int): ID del cliente.
        """
        cursor = self.connection.cursor()
        cursor.execute('DELETE FROM aliases WHERE client_id = ?', (client_id,))
        cursor.execute('DELETE FROM client_list WHERE client_id = ?', (client_id,))
        cursor.execute('DELETE FROM clientes WHERE id = ?', (client_id,))
        self.connection.commit()

    def get_client_by_id(self, client_id: int) -> Optional[Tuple[int, str, str, str, str]]:
        """
        Obtiene un cliente por su ID.

        Args:
            client_id (int): ID del cliente.

        Returns:
            Optional[Tuple[int, str, str, str, str]]: Tupla con (id, nombre, correo, correo_cc, correo_bcc) o None si no se encuentra.
        """
        cursor = self.connection.cursor()
        cursor.execute('SELECT id, nombre, correo, correo_cc, correo_bcc FROM clientes WHERE id = ?', (client_id,))
        return cursor.fetchone()

    def add_alias(self, client_id: int, alias: str):
        """
        Agrega un alias a un cliente.

        Args:
            client_id (int): ID del cliente.
            alias (str): Alias a agregar.
        """
        cursor = self.connection.cursor()
        cursor.execute('''
            INSERT INTO aliases (client_id, alias)
            VALUES (?, ?)
        ''', (client_id, alias))
        self.connection.commit()

    def remove_alias(self, alias_id: int):
        """
        Elimina un alias de la base de datos.

        Args:
            alias_id (int): ID del alias.
        """
        cursor = self.connection.cursor()
        cursor.execute('DELETE FROM aliases WHERE id = ?', (alias_id,))
        self.connection.commit()

    def get_aliases_for_client(self, client_id: int) -> List[Tuple[int, str]]:
        """
        Obtiene los aliases de un cliente.

        Args:
            client_id (int): ID del cliente.

        Returns:
            List[Tuple[int, str]]: Lista de tuplas con (alias_id, alias).
        """
        cursor = self.connection.cursor()
        cursor.execute('SELECT id, alias FROM aliases WHERE client_id = ?', (client_id,))
        return cursor.fetchall()

    def get_all_clients(self) -> List[Tuple[int, str, str, str, str]]:
        """
        Obtiene todos los clientes de la base de datos.

        Returns:
            List[Tuple[int, str, str, str, str]]: Lista de tuplas con (id, nombre, correo, correo_cc, correo_bcc).
        """
        cursor = self.connection.cursor()
        cursor.execute('SELECT id, nombre, correo, correo_cc, correo_bcc FROM clientes')
        return cursor.fetchall()

    def get_clients_with_email(self) -> List[Tuple[int, str, str]]:
        """
        Obtiene una lista de clientes que tienen correo asociado.

        Returns:
            List[Tuple[int, str, str]]: Lista de tuplas con (id, nombre, correo).
        """
        cursor = self.connection.cursor()
        cursor.execute('''
            SELECT id, nombre, correo FROM clientes
            WHERE correo IS NOT NULL AND correo != '' AND LOWER(TRIM(correo)) != '(correo sin definir)'
        ''')
        return cursor.fetchall()

    def close(self):
        """
        Cierra la conexión a la base de datos.
        """
        self.connection.close()

    # Métodos para manejar listas

    def add_list(self, nombre_lista: str):
        """
        Agrega una nueva lista a la base de datos.

        Args:
            nombre_lista (str): Nombre de la lista.
        """
        cursor = self.connection.cursor()
        cursor.execute('''
            INSERT INTO listas (nombre)
            VALUES (?)
        ''', (nombre_lista,))
        self.connection.commit()

    def delete_list(self, list_id: int):
        """
        Elimina una lista y sus asociaciones.

        Args:
            list_id (int): ID de la lista.
        """
        cursor = self.connection.cursor()
        cursor.execute('DELETE FROM client_list WHERE list_id = ?', (list_id,))
        cursor.execute('DELETE FROM listas WHERE id = ?', (list_id,))
        self.connection.commit()

    def get_all_lists(self) -> List[Tuple[int, str]]:
        """
        Obtiene todas las listas.

        Returns:
            List[Tuple[int, str]]: Lista de tuplas con (id, nombre).
        """
        cursor = self.connection.cursor()
        cursor.execute('SELECT id, nombre FROM listas')
        return cursor.fetchall()

    def add_client_to_list(self, client_id: int, list_id: int):
        """
        Asocia un cliente a una lista si la asociación no existe y list_id no es None.

        Args:
            client_id (int): ID del cliente.
            list_id (int): ID de la lista.
        """
        # Verificar que list_id no sea None
        if list_id is None:
            raise ValueError(f"list_id no puede ser None al agregar cliente con ID {client_id} a la lista")

        cursor = self.connection.cursor()

        # Verificar si ya existe la asociación entre el cliente y la lista
        cursor.execute('''
            SELECT 1 FROM client_list WHERE client_id = ? AND list_id = ?
        ''', (client_id, list_id))
        result = cursor.fetchone()

        # Solo agregar si la asociación no existe
        if not result:
            cursor.execute('''
                INSERT INTO client_list (client_id, list_id)
                VALUES (?, ?)
            ''', (client_id, list_id))
            self.connection.commit()
        else:
            print(f"La asociación del cliente con ID {client_id} y la lista con ID {list_id} ya existe.")

    def remove_client_from_list(self, client_id: int, list_id: int):
        """
        Elimina la asociación de un cliente con una lista.

        Args:
            client_id (int): ID del cliente.
            list_id (int): ID de la lista.
        """
        cursor = self.connection.cursor()
        cursor.execute('''
            DELETE FROM client_list
            WHERE client_id = ? AND list_id = ?
        ''', (client_id, list_id))
        self.connection.commit()

    def get_clients_by_list(self, list_id: int) -> List[Tuple[int, str, str, str, str]]:
        """
        Obtiene los clientes asociados a una lista.

        Args:
            list_id (int): ID de la lista.

        Returns:
            List[Tuple[int, str, str, str, str]]: Lista de tuplas con (id, nombre, correo, correo_cc, correo_bcc).
        """
        cursor = self.connection.cursor()
        cursor.execute('''
            SELECT c.id, c.nombre, c.correo, c.correo_cc, c.correo_bcc
            FROM clientes c
            INNER JOIN client_list cl ON c.id = cl.client_id
            WHERE cl.list_id = ?
        ''', (list_id,))
        return cursor.fetchall()

    def get_client_id_by_name(self, nombre: str) -> Optional[int]:
        cursor = self.connection.cursor()
        cursor.execute('SELECT id FROM clientes WHERE nombre = ?', (nombre,))
        result = cursor.fetchone()
        return result[0] if result else None

    def initialize_default_list(self):
        """
        Inicializa la lista 'Predeterminada' si no existe.
        """
        cursor = self.connection.cursor()
        cursor.execute('SELECT id FROM listas WHERE nombre = ?', ('Predeterminada',))
        result = cursor.fetchone()
        if not result:
            self.add_list('Predeterminada')
        self.connection.commit()

    def get_client_lists(self, client_id: int) -> List[int]:
        """
        Obtiene las listas a las que pertenece un cliente.

        Args:
            client_id (int): ID del cliente.

        Returns:
            List[int]: Lista de IDs de listas.
        """
        cursor = self.connection.cursor()
        cursor.execute('SELECT list_id FROM client_list WHERE client_id = ?', (client_id,))
        rows = cursor.fetchall()
        return [row[0] for row in rows]

    def get_list_names(self) -> List[str]:
        """
        Obtiene los nombres de todas las listas.

        Returns:
            List[str]: Lista de nombres de listas.
        """
        cursor = self.connection.cursor()
        cursor.execute('SELECT nombre FROM listas')
        return [row[0] for row in cursor.fetchall()]

    def client_exists(self, client_id: int) -> bool:
        cursor = self.connection.cursor()
        cursor.execute('SELECT 1 FROM clientes WHERE id = ?', (client_id,))
        return cursor.fetchone() is not None

    def list_exists(self, list_id: int) -> bool:
        cursor = self.connection.cursor()
        cursor.execute('SELECT 1 FROM listas WHERE id = ?', (list_id,))
        return cursor.fetchone() is not None

    def get_list_id_by_name(self, list_name: str) -> Optional[int]:
        """
        Obtiene el ID de una lista según su nombre.

        Args:
            list_name (str): Nombre de la lista.

        Returns:
            Optional[int]: ID de la lista si existe, None en caso contrario.
        """
        cursor = self.connection.cursor()
        cursor.execute("SELECT id FROM listas WHERE nombre = ?", (list_name,))
        result = cursor.fetchone()
        return result[0] if result else None

    def alias_exists(self, client_id: int, alias: str) -> bool:
        """
        Verifica si un alias específico ya existe para un cliente dado.

        Args:
            client_id (int): ID del cliente.
            alias (str): Alias a verificar.

        Returns:
            bool: True si el alias existe, False en caso contrario.
        """
        cursor = self.connection.cursor()
        cursor.execute("SELECT 1 FROM aliases WHERE client_id = ? AND alias = ?", (client_id, alias))
        return cursor.fetchone() is not None

    def list_exists_by_name(self, list_name: str) -> bool:
        """
        Verifica si una lista con un nombre específico ya existe en la base de datos.

        Args:
            list_name (str): Nombre de la lista a verificar.

        Returns:
            bool: True si la lista existe, False en caso contrario.
        """
        cursor = self.connection.cursor()
        cursor.execute("SELECT 1 FROM listas WHERE nombre = ?", (list_name,))
        return cursor.fetchone() is not None

    def client_exists_by_name(self, client_name: str) -> bool:
        """
        Verifica si un cliente con un nombre específico ya existe en la base de datos.

        Args:
            client_name (str): Nombre del cliente a verificar.

        Returns:
            bool: True si el cliente existe, False en caso contrario.
        """
        cursor = self.connection.cursor()
        cursor.execute("SELECT 1 FROM clientes WHERE nombre = ?", (client_name,))
        return cursor.fetchone() is not None

    def is_client_in_list(self, client_id: int, list_id: int) -> bool:
        """
        Verifica si un cliente está en una lista específica.

        Args:
            client_id (int): ID del cliente.
            list_id (int): ID de la lista.

        Returns:
            bool: True si el cliente está en la lista, False en caso contrario.
        """
        cursor = self.connection.cursor()
        cursor.execute('''
            SELECT 1 FROM client_list WHERE client_id = ? AND list_id = ?
        ''', (client_id, list_id))
        return cursor.fetchone() is not None

    def update_alias(self, alias_id: int, new_alias: str):
        """
        Actualiza un alias específico para un cliente.

        Args:
            alias_id (int): ID del alias a actualizar.
            new_alias (str): Nuevo valor del alias.
        """
        cursor = self.connection.cursor()
        cursor.execute('''
            UPDATE aliases
            SET alias = ?
            WHERE id = ?
        ''', (new_alias, alias_id))
        self.connection.commit()
