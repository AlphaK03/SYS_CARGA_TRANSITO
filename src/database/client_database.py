# src/database/client_database.py
import os
import sqlite3
from typing import List, Optional, Tuple
import pandas as pd
import json

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
        except sqlite3.OperationalError as e:
            raise Exception(f"No se pudo abrir la base de datos: {e}")

    def create_tables(self):
        """
        Crea las tablas necesarias si no existen.
        """
        cursor = self.connection.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS clientes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre TEXT UNIQUE NOT NULL,
                correo TEXT DEFAULT '(correo sin definir)'
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
        self.connection.commit()

    # Métodos para manejar clientes y aliases

    def add_client(self, nombre: str, correo: str = '(correo sin definir)'):
        """
        Agrega un nuevo cliente a la base de datos.

        Args:
            nombre (str): Nombre del cliente.
            correo (str): Correo electrónico del cliente.
        """
        cursor = self.connection.cursor()
        cursor.execute('''
            INSERT INTO clientes (nombre, correo)
            VALUES (?, ?)
        ''', (nombre, correo))
        self.connection.commit()

    def update_client_name_and_email(self, client_id: int, nombre: str, correo: str):
        """
        Actualiza el nombre y el correo de un cliente.

        Args:
            client_id (int): ID del cliente.
            nombre (str): Nuevo nombre del cliente.
            correo (str): Nuevo correo electrónico del cliente.
        """
        cursor = self.connection.cursor()
        cursor.execute('''
            UPDATE clientes
            SET nombre = ?, correo = ?
            WHERE id = ?
        ''', (nombre, correo, client_id))
        self.connection.commit()

    def delete_client(self, client_id: int):
        """
        Elimina un cliente y sus aliases de la base de datos.

        Args:
            client_id (int): ID del cliente.
        """
        cursor = self.connection.cursor()
        cursor.execute('DELETE FROM aliases WHERE client_id = ?', (client_id,))
        cursor.execute('DELETE FROM clientes WHERE id = ?', (client_id,))
        self.connection.commit()

    def get_client_by_id(self, client_id: int) -> Optional[Tuple[int, str, str]]:
        """
        Obtiene un cliente por su ID.

        Args:
            client_id (int): ID del cliente.

        Returns:
            Optional[Tuple[int, str, str]]: Tupla con (id, nombre, correo) o None si no se encuentra.
        """
        cursor = self.connection.cursor()
        cursor.execute('SELECT id, nombre, correo FROM clientes WHERE id = ?', (client_id,))
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

    def get_all_clients(self) -> List[Tuple[int, str, str]]:
        """
        Obtiene todos los clientes de la base de datos.

        Returns:
            List[Tuple[int, str, str]]: Lista de tuplas con (id, nombre, correo).
        """
        cursor = self.connection.cursor()
        cursor.execute('SELECT id, nombre, correo FROM clientes')
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