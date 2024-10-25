# src/processing/file_processor.py
import os
import pandas as pd
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill
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
            List[str]: Lista de nombres de clientes procesados.
        """
        # Leer el archivo Excel desde la hoja "Report SI SYSTEM - ASW", comenzando en la fila 8
        df = pd.read_excel(self.file_path, sheet_name='Report SI SYSTEM - Customers', header=7)

        # Remover espacios en blanco de los nombres de las columnas
        df.columns = df.columns.str.strip()

        # Verificar si 'Consignee' está en las columnas
        if 'Consignee' not in df.columns:
            raise ValueError("La columna 'Consignee' no se encontró.")

        # Obtener los clientes registrados y sus aliases
        clients = self.db.get_all_clients()
        clients_dict = {}
        for client_id, nombre, correo in clients:
            aliases = self.db.get_aliases_for_client(client_id)
            alias_list = [alias for alias_id, alias in aliases]
            clients_dict[client_id] = {
                'nombre': nombre,
                'correo': correo,
                'aliases': alias_list
            }

        if not clients_dict:
            raise ValueError("No hay clientes registrados en la base de datos.")

        os.makedirs(self.output_dir, exist_ok=True)

        for client_id, client_data in clients_dict.items():
            nombre = client_data['nombre']
            aliases = client_data['aliases']

            if not aliases:
                continue  # Si el cliente no tiene aliases, continuar

            # Filtrar los datos que coincidan con cualquiera de los aliases
            df_cliente = df[df['Consignee'].isin(aliases)]

            if df_cliente.empty:
                continue  # Si no hay datos, continuar

            output_file = os.path.join(self.output_dir, f'{nombre}.xlsx')

            # Crear el archivo Excel con formato
            self._save_excel_with_format(df_cliente, output_file)

        return [client_data['nombre'] for client_id, client_data in clients_dict.items()]

    def _save_excel_with_format(self, df_cliente: pd.DataFrame, output_file: str):
        """
        Guarda el DataFrame en un archivo Excel con formato personalizado.

        Args:
            df_cliente (pd.DataFrame): DataFrame del cliente a guardar.
            output_file (str): Ruta del archivo Excel a generar.
        """
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_cliente.to_excel(writer, index=False, startrow=7, sheet_name='Clientes')

            # Obtener el workbook y el worksheet
            workbook = writer.book
            worksheet = writer.sheets['Clientes']

            # Ajustar el ancho de columnas
            for col_num, column in enumerate(df_cliente.columns, 1):
                max_length = df_cliente[column].astype(str).map(len).max() + 2
                worksheet.column_dimensions[get_column_letter(col_num)].width = max(max_length, 15)

            # Aplicar color de fondo en la fila de encabezado
            header_fill = PatternFill(start_color='63B9FF', end_color='63B9FF', fill_type='solid')
            for cell in worksheet[8]:  # La fila 8, ya que los datos empiezan en la fila 8 después del encabezado
                cell.fill = header_fill
                cell.font = Font(bold=True, color='FFFFFF')

            # Ajustar el ancho de la columna "Consignee" si existe
            if 'Consignee' in df_cliente.columns:
                consignee_col = df_cliente.columns.get_loc('Consignee') + 1
                worksheet.column_dimensions[get_column_letter(consignee_col)].width = 30

            # Aplicar el formato de fuente a las celdas de datos
            for row in worksheet.iter_rows(min_row=8, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
                for cell in row:
                    cell.font = Font(name='Calibri')

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
