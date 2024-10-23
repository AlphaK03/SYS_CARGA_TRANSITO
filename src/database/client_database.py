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
            self.create_table()
        except sqlite3.OperationalError as e:
            raise Exception(f"No se pudo abrir la base de datos: {e}")


    def import_clients(self, file_path: str, file_type: str):
        """
        Importa clientes desde un archivo especificado.

        Args:
            file_path (str): Ruta al archivo a importar.
            file_type (str): Tipo de archivo ('excel', 'csv', 'json').
        """
        if file_type == 'excel':
            df = pd.read_excel(file_path)
        elif file_type == 'csv':
            df = pd.read_csv(file_path)
        elif file_type == 'json':
            with open(file_path, 'r') as f:
                data = json.load(f)
            df = pd.DataFrame(data)
        else:
            raise ValueError("Tipo de archivo no soportado.")

        # Se espera que el DataFrame tenga columnas 'nombre' y 'correo'
        if 'nombre' not in df.columns or 'correo' not in df.columns:
            raise ValueError("El archivo debe contener las columnas 'nombre' y 'correo'.")

        # Insertar o actualizar los registros
        cursor = self.connection.cursor()
        for _, row in df.iterrows():
            nombre = row['nombre']
            correo = row['correo']
            cursor.execute('''
                INSERT INTO clientes (nombre, correo)
                VALUES (?, ?)
                ON CONFLICT(nombre) DO UPDATE SET correo=excluded.correo
            ''', (nombre, correo))
        self.connection.commit()

    def export_clients(self, file_path: str, file_type: str):
        """
        Exporta la lista de clientes a un archivo especificado.

        Args:
            file_path (str): Ruta al archivo de destino.
            file_type (str): Tipo de archivo ('excel', 'csv', 'json').
        """
        cursor = self.connection.cursor()
        cursor.execute('SELECT nombre, correo FROM clientes')
        rows = cursor.fetchall()
        df = pd.DataFrame(rows, columns=['nombre', 'correo'])

        if file_type == 'excel':
            df.to_excel(file_path, index=False)
        elif file_type == 'csv':
            df.to_csv(file_path, index=False)
        elif file_type == 'json':
            df.to_json(file_path, orient='records', indent=4)
        else:
            raise ValueError("Tipo de archivo no soportado.")

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
