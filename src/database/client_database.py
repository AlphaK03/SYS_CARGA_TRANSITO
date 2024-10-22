# src/database/client_database.py
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
            self.create_table()
        except sqlite3.OperationalError as e:
            raise Exception(f"No se pudo abrir la base de datos: {e}")

    def create_table(self):
        """
        Crea la tabla de clientes si no existe.
        """
        cursor = self.connection.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS clientes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT UNIQUE NOT NULL,
                correo TEXT DEFAULT '(correo sin definir)'
            )
        ''')
        self.connection.commit()

    # [Resto de los métodos de la clase]

    def add_clients(self, clients: List[str]):
        """
        Agrega nuevos clientes a la base de datos.

        Args:
            clients (List[str]): Lista de nombres de clientes a agregar.
        """
        cursor = self.connection.cursor()
        for client in clients:
            try:
                cursor.execute('''
                    INSERT INTO clientes (nombre)
                    VALUES (?)
                ''', (client,))
            except sqlite3.IntegrityError:
                # El cliente ya existe, continuar
                continue
        self.connection.commit()

    def get_all_clients(self) -> List[Tuple[int, str, str]]:
        """
        Obtiene todos los clientes de la base de datos.

        Returns:
            List[Tuple[int, str, str]]: Lista de tuplas con (id, nombre, correo).
        """
        cursor = self.connection.cursor()
        cursor.execute('SELECT id, nombre, correo FROM clientes')
        return cursor.fetchall()

    def get_client_by_name(self, name: str) -> Optional[Tuple[int, str, str]]:
        """
        Obtiene un cliente por su nombre.

        Args:
            name (str): Nombre del cliente.

        Returns:
            Optional[Tuple[int, str, str]]: Tupla con (id, nombre, correo) o None si no se encuentra.
        """
        cursor = self.connection.cursor()
        cursor.execute('SELECT id, nombre, correo FROM clientes WHERE nombre = ?', (name,))
        return cursor.fetchone()

    def update_client_email(self, client_id: int, email: str):
        """
        Actualiza el correo electrónico de un cliente.

        Args:
            client_id (int): ID del cliente.
            email (str): Nuevo correo electrónico.
        """
        cursor = self.connection.cursor()
        cursor.execute('UPDATE clientes SET correo = ? WHERE id = ?', (email, client_id))
        self.connection.commit()

    def get_clients_without_email(self) -> List[Tuple[int, str]]:
        """
        Obtiene una lista de clientes que no tienen correo asociado.

        Returns:
            List[Tuple[int, str]]: Lista de tuplas con (id, nombre).
        """
        cursor = self.connection.cursor()
        cursor.execute('''
            SELECT id, nombre FROM clientes
            WHERE correo IS NULL OR correo = '' OR LOWER(TRIM(correo)) = '(correo sin definir)'
        ''')
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
