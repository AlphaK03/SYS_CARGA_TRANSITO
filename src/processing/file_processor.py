# src/processing/file_processor.py

import os
import sys

import requests
import win32com.client
import pandas as pd
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from openpyxl.styles import Font, PatternFill
from typing import List, Optional
from openpyxl import load_workbook
import win32com.client  # Importar para interactuar con Excel
import time

from database.client_database import ClientDatabase
from utils.config_manager import ConfigManager

def get_resource_path(relative_path):
    """
    Obtiene la ruta del recurso empaquetado o del entorno de desarrollo.
    """
    # Si está empaquetado, sys._MEIPASS estará disponible
    base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
    return os.path.join(base_path, relative_path)

def _save_excel_with_format(df_cliente: pd.DataFrame, template_path: str, output_file: str):
    """
    Guarda el DataFrame en una copia de la plantilla de Excel,
    agregando los datos a partir de la fila 5.

    Args:
        df_cliente (pd.DataFrame): DataFrame del cliente a guardar.
        template_path (str): Ruta de la plantilla de Excel.
        output_file (str): Ruta del archivo Excel a generar.
    """
    # Cargar la plantilla
    workbook = load_workbook(template_path)
    worksheet = workbook.active  # Suponiendo que la plantilla tiene una sola hoja

    # Escribir los nombres de las columnas en la fila 4 (encabezado)
    for col_idx, column_name in enumerate(df_cliente.columns, start=1):
        cell = worksheet.cell(row=4, column=col_idx)
        cell.value = column_name
        # Aplicar formato al encabezado si es necesario
        cell.fill = PatternFill(start_color='63B9FF', end_color='63B9FF', fill_type='solid')
        cell.font = Font(bold=True, color='FFFFFF')

    # Escribir los datos a partir de la fila 5
    for row_idx, row in enumerate(df_cliente.values, start=5):
        for col_idx, value in enumerate(row, start=1):
            worksheet.cell(row=row_idx, column=col_idx, value=value)

    # Guardar el nuevo archivo Excel
    workbook.save(output_file)

class FileProcessor:
    def __init__(self, site_url: str, sharepoint_file_url: str, access_token: str):
        self.site_url = site_url
        self.sharepoint_file_url = sharepoint_file_url
        self.access_token = access_token
        self.output_dir = os.path.join('data', 'Clientes_Separados')
        self.db = ClientDatabase()
        self.config_manager = ConfigManager()
        self.local_file_path = None  # Ruta al archivo descargado

    def download_file(self):
        if not self.access_token:
            raise Exception("El token de acceso no está disponible.")

        # Crear carpeta temporal si no existe
        os.makedirs('data/temp', exist_ok=True)
        file_name = os.path.basename(self.sharepoint_file_url)
        self.local_file_path = os.path.join('data/temp', file_name)

        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Accept': 'application/json;odata=verbose'
        }

        # Construir la URL completa del archivo
        download_url = f"{self.site_url}/_api/web/GetFileByServerRelativeUrl('{self.sharepoint_file_url}')/$value"

        try:
            response = requests.get(download_url, headers=headers)
            response.raise_for_status()
            with open(self.local_file_path, 'wb') as local_file:
                local_file.write(response.content)
            print(f"Archivo descargado exitosamente en {self.local_file_path}.")
        except requests.exceptions.RequestException as e:
            raise Exception(f"No se pudo descargar el archivo: {e}")


    def authenticate_sharepoint(self, username: str, password: str):
        """
        Autentica en SharePoint y crea el contexto.

        Args:
            username (str): Nombre de usuario de SharePoint.
            password (str): Contraseña de SharePoint.
        """
        if not username or not password:
            raise ValueError("Las credenciales de SharePoint no están configuradas.")

        ctx_auth = AuthenticationContext(self.site_url)
        if ctx_auth.acquire_token_for_user(username, password):
            self.ctx = ClientContext(self.site_url, ctx_auth)
        else:
            raise Exception("Error de autenticación en SharePoint. Por favor, verifica tus credenciales.")


    def run_macro(self):
        """
        Ejecuta la macro 'ActualizarPBI' en una nueva instancia de Excel y la cierra después de un retraso de 2 segundos.
        """
        if not self.local_file_path:
            raise Exception("El archivo no ha sido descargado aún.")

        try:
            # Crear una nueva instancia de Excel
            excel_app = win32com.client.DispatchEx("Excel.Application")
            excel_app.Visible = False  # Mantener Excel no visible
            excel_app.DisplayAlerts = False
            excel_app.AskToUpdateLinks = False
            excel_app.EnableEvents = False

            print("Excel iniciado...")

            # Intentar abrir el archivo de Excel
            workbook = excel_app.Workbooks.Open(os.path.abspath(self.local_file_path), ReadOnly=False)
            print("Archivo de Excel abierto...")

            # Añadir un tiempo de espera antes de ejecutar la macro
            time.sleep(1)

            macro_name = 'ActualizarPBI'

            # Reintentar la ejecución de la macro hasta 3 veces si falla
            for intento in range(3):
                try:
                    print(f"Ejecutando macro '{macro_name}', intento {intento + 1}...")
                    excel_app.Run(f"'{workbook.Name}'!{macro_name}")
                    print("Macro ejecutada exitosamente.")
                    break  # Salir del bucle si la macro se ejecuta sin errores
                except Exception as e:
                    print(f"Error al ejecutar la macro en el intento {intento + 1}: {str(e)}")
                    if intento < 2:
                        time.sleep(2)  # Esperar antes de reintentar
                    else:
                        raise Exception(f"Error al ejecutar la macro '{macro_name}': {str(e)}")

            # Esperar 2 segundos antes de cerrar

            # Cerrar el archivo de Excel y la instancia
            workbook.Close(SaveChanges=True)
            excel_app.Quit()
            print("Instancia de Excel cerrada.")

        except Exception as e:
            raise Exception(f"Error al iniciar o manejar Excel: {str(e)}")

    def process_file(self) -> List[str]:
        """
        Procesa el archivo Excel, genera reportes separados por cliente y actualiza la base de datos de clientes.

        Returns:
            List[str]: Lista de nombres de clientes procesados.
        """
        # Leer el archivo Excel desde la hoja correspondiente
        df = pd.read_excel(self.local_file_path, sheet_name='Report SI SYSTEM - Customers', header=3)

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

        template_path = get_resource_path('template/plantilla_Excel.xlsx')

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
            _save_excel_with_format(df_cliente, template_path, output_file)

        return [client_data['nombre'] for client_id, client_data in clients_dict.items()]

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
