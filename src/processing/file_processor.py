# src/processing/file_processor.py
import os
import pandas as pd
from typing import List, Optional

from database.client_database import ClientDatabase

class FileProcessor:
    """
    Clase para procesar archivos Excel y generar reportes por cliente, y manejar la base de datos de clientes.
    """

    def __init__(self, file_path: str):
        """
        Inicializa una instancia de FileProcessor.

        Args:
            file_path (str): Ruta al archivo Excel a procesar.
        """
        self.file_path = file_path
        self.output_dir = os.path.join('data', 'Clientes_Separados')
        self.db = ClientDatabase()

    def process_file(self) -> List[str]:
        """
        Procesa el archivo Excel, genera reportes separados por cliente y actualiza la base de datos de clientes.

        Returns:
            List[str]: Lista de clientes únicos encontrados en el archivo.
        """
        # Leer el archivo Excel con la fila correcta como encabezado
        df = pd.read_excel(self.file_path, sheet_name='Report SI SYSTEM - Customers', header=7)
        # Remover espacios en blanco de los nombres de las columnas
        df.columns = df.columns.str.strip()

        # Verificar si 'Consignee' está en las columnas
        if 'Consignee' not in df.columns:
            raise ValueError(
                f"La columna 'Consignee' no se encontró. Las columnas disponibles son: {df.columns.tolist()}")

        # Filtrar por clientes únicos en la columna 'Consignee'
        clientes_unicos = df['Consignee'].dropna().unique().tolist()

        # Crear un directorio para guardar los archivos
        os.makedirs(self.output_dir, exist_ok=True)

        # Crear un archivo separado por cada cliente
        for cliente in clientes_unicos:
            df_cliente = df[df['Consignee'] == cliente]
            output_file = os.path.join(self.output_dir, f'{cliente}.xlsx')
            df_cliente.to_excel(output_file, index=False)

        # Actualizar la base de datos de clientes
        self.db.add_clients(clientes_unicos)

        return clientes_unicos

    def get_clients_without_email(self) -> List[str]:
        """
        Obtiene una lista de nombres de clientes que no tienen correo asociado.

        Returns:
            List[str]: Lista de nombres de clientes sin correo.
        """
        clients = self.db.get_clients_without_email()
        return [client[1] for client in clients]  # client[1] es el nombre

    def get_clients_with_email(self) -> List[tuple]:
        """
        Obtiene una lista de clientes con correo asociado.

        Returns:
            List[tuple]: Lista de tuplas (id, nombre, correo).
        """
        return self.db.get_clients_with_email()

    def get_client_file(self, cliente: str) -> Optional[str]:
        """
        Obtiene la ruta del archivo correspondiente a un cliente.

        Args:
            cliente (str): Nombre del cliente.

        Returns:
            Optional[str]: Ruta al archivo del cliente, o None si no existe.
        """
        archivo = os.path.join(self.output_dir, f'{cliente}.xlsx')
        if os.path.exists(archivo):
            return archivo
        else:
            return None

    def close(self):
        """
        Cierra la conexión a la base de datos.
        """
        self.db.close()
