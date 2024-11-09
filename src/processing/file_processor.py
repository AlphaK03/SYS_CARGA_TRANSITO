# src/processing/file_processor.py

import os
import sys
import time
import logging
from typing import List, Optional

import requests
import win32com.client
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from urllib.parse import quote, unquote

from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext

from database.client_database import ClientDatabase
from utils.config_manager import ConfigManager

from src.processing.custom_auth_provider import TokenAuthenticationContext

# Configuración del logger
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)  # Ajusta el nivel según tus necesidades
handler = logging.StreamHandler(sys.stdout)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)


def get_resource_path(relative_path: str) -> str:
    """
    Obtiene la ruta del recurso empaquetado o del entorno de desarrollo.

    Args:
        relative_path (str): Ruta relativa al recurso.

    Returns:
        str: Ruta absoluta al recurso.
    """
    try:
        # Si está empaquetado, sys._MEIPASS estará disponible
        base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
        resource_path = os.path.join(base_path, relative_path)
        logger.debug(f"Ruta de recurso obtenida: {resource_path}")
        return resource_path
    except Exception as e:
        logger.exception(f"Error al obtener la ruta del recurso: {e}")
        raise


def _save_excel_with_format(df_cliente: pd.DataFrame, template_path: str, output_file: str) -> None:
    """
    Guarda el DataFrame en una copia de la plantilla de Excel,
    agregando los datos a partir de la fila 5.

    Args:
        df_cliente (pd.DataFrame): DataFrame del cliente a guardar.
        template_path (str): Ruta de la plantilla de Excel.
        output_file (str): Ruta del archivo Excel a generar.
    """
    try:
        if not os.path.exists(template_path):
            logger.error(f"La plantilla de Excel no existe en la ruta: {template_path}")
            raise FileNotFoundError(f"La plantilla de Excel no existe en la ruta: {template_path}")

        # Cargar la plantilla
        workbook = load_workbook(template_path)
        worksheet = workbook.active  # Suponiendo que la plantilla tiene una sola hoja
        logger.debug(f"Plantilla de Excel cargada desde: {template_path}")

        # Escribir los nombres de las columnas en la fila 4 (encabezado)
        for col_idx, column_name in enumerate(df_cliente.columns, start=1):
            cell = worksheet.cell(row=4, column=col_idx)
            cell.value = column_name
            # Aplicar formato al encabezado si es necesario
            cell.fill = PatternFill(start_color='63B9FF', end_color='63B9FF', fill_type='solid')
            cell.font = Font(bold=True, color='FFFFFF')
            logger.debug(f"Columna '{column_name}' añadida y formateada en la celda ({4}, {col_idx}).")

        # Escribir los datos a partir de la fila 5
        for row_idx, row in enumerate(df_cliente.itertuples(index=False), start=5):
            for col_idx, value in enumerate(row, start=1):
                worksheet.cell(row=row_idx, column=col_idx, value=value)
            logger.debug(f"Fila {row_idx} añadida al archivo Excel.")

        # Guardar el nuevo archivo Excel
        workbook.save(output_file)
        logger.info(f"Archivo Excel guardado exitosamente en: {output_file}")

    except Exception as e:
        logger.exception(f"Error al guardar el archivo Excel con formato: {e}")
        raise



class FileProcessor:
    def __init__(self, site_url: str, sharepoint_file_url: str, access_token: str):
        """
        Initializes the FileProcessor with the necessary configuration.

        Args:
            site_url (str): The SharePoint site URL.
            sharepoint_file_url (str): The relative URL of the file in SharePoint.
            access_token (str): Access token for authentication.
        """
        self.site_url = site_url
        self.sharepoint_file_url = sharepoint_file_url
        self.access_token = access_token
        self.local_file_path = None

        self.output_dir = os.path.join('data', 'Clientes_Separados')
        self.db = ClientDatabase()
        self.config_manager = ConfigManager()

        # Asegurar que el directorio de salida exista
        os.makedirs(self.output_dir, exist_ok=True)
        logger.debug(f"Directorio de salida asegurado en: {self.output_dir}")

    def download_file(self) -> str:
        """
        Downloads the file from SharePoint and saves it locally using the ClientContext with the access token.

        Returns:
            str: The local file path of the downloaded file.
        """
        # Create a temporary directory if it doesn't exist
        os.makedirs('data/temp', exist_ok=True)
        file_name = os.path.basename(self.sharepoint_file_url)
        self.local_file_path = os.path.join('data/temp', file_name)

        try:
            # Set up authentication and ClientContext
            ctx_auth = TokenAuthenticationContext(self.site_url, self.access_token)
            ctx = ClientContext(self.site_url, ctx_auth)

            # Retrieve the file from SharePoint
            file = ctx.web.get_file_by_server_relative_url(self.sharepoint_file_url)
            with open(self.local_file_path, "wb") as local_file:
                file.download(local_file)
                ctx.execute_query()

            logger.info(f"File downloaded successfully to {self.local_file_path}")
            return self.local_file_path
        except Exception as e:
            logger.error(f"Failed to download file from SharePoint: {e}")
            raise Exception(f"Failed to download file from SharePoint: {e}")



    def authenticate_sharepoint(self, username: str, password: str) -> None:
        """
        Autentica en SharePoint y crea el contexto.

        Args:
            username (str): Nombre de usuario de SharePoint.
            password (str): Contraseña de SharePoint.
        """
        if not username or not password:
            logger.error("Las credenciales de SharePoint no están configuradas.")
            raise ValueError("Las credenciales de SharePoint no están configuradas.")

        try:
            ctx_auth = AuthenticationContext(self.site_url)
            if ctx_auth.acquire_token_for_user(username, password):
                self.ctx = ClientContext(self.site_url, ctx_auth)
                logger.info("Autenticación en SharePoint exitosa.")
            else:
                logger.error("Error de autenticación en SharePoint. Verifica tus credenciales.")
                raise Exception("Error de autenticación en SharePoint. Por favor, verifica tus credenciales.")
        except Exception as e:
            logger.exception(f"Error durante la autenticación en SharePoint: {e}")
            raise

    def run_macro(self, macro_name: str = 'ActualizarPBI') -> None:
        """
        Ejecuta una macro en el archivo de Excel descargado y maneja la instancia de Excel.

        Args:
            macro_name (str): Nombre de la macro a ejecutar. Por defecto, 'ActualizarPBI'.
        """
        if not self.local_file_path:
            logger.error("El archivo no ha sido descargado aún.")
            raise Exception("El archivo no ha sido descargado aún.")

        excel_app = None
        workbook = None
        try:
            # Crear una nueva instancia de Excel
            excel_app = win32com.client.DispatchEx("Excel.Application")
            excel_app.Visible = False  # Mantener Excel no visible
            excel_app.DisplayAlerts = False
            excel_app.AskToUpdateLinks = False
            excel_app.EnableEvents = False
            logger.debug("Instancia de Excel iniciada.")

            # Abrir el archivo de Excel
            workbook = excel_app.Workbooks.Open(os.path.abspath(self.local_file_path), ReadOnly=False)
            logger.info(f"Archivo de Excel abierto: {self.local_file_path}")

            # Ejecutar la macro
            logger.debug(f"Ejecutando macro '{macro_name}'...")
            excel_app.Run(f"'{workbook.Name}'!{macro_name}")
            logger.info(f"Macro '{macro_name}' ejecutada exitosamente.")

            # Opcional: Guardar cambios si la macro los realiza
            workbook.Save()
            logger.debug("Cambios en el archivo de Excel guardados.")

        except Exception as e:
            logger.exception(f"Error al ejecutar la macro '{macro_name}': {e}")
            raise Exception(f"Error al ejecutar la macro '{macro_name}': {e}")
        finally:
            # Cerrar el archivo de Excel y la instancia
            if workbook:
                workbook.Close(SaveChanges=True)
                logger.debug("Archivo de Excel cerrado.")
            if excel_app:
                excel_app.Quit()
                logger.debug("Instancia de Excel cerrada.")

    def process_file(self) -> List[str]:
        """
        Procesa el archivo Excel, genera reportes separados por cliente y actualiza la base de datos de clientes.

        Returns:
            List[str]: Lista de nombres de clientes procesados.
        """
        if not self.local_file_path or not os.path.exists(self.local_file_path):
            logger.error("El archivo local no existe o no ha sido descargado.")
            raise Exception("El archivo local no existe o no ha sido descargado.")

        try:
            # Leer el archivo Excel desde la hoja correspondiente
            df = pd.read_excel(self.local_file_path, sheet_name='Report SI SYSTEM - Customers', header=3)
            logger.debug(f"Datos leídos del archivo Excel: {self.local_file_path}")

            # Remover espacios en blanco de los nombres de las columnas
            df.columns = df.columns.str.strip()
            logger.debug("Espacios en blanco removidos de los nombres de las columnas.")

            # Verificar si 'Consignee' está en las columnas
            if 'Consignee' not in df.columns:
                logger.error("La columna 'Consignee' no se encontró en el archivo Excel.")
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
            logger.debug(f"Clientes obtenidos de la base de datos: {clients_dict.keys()}")

            if not clients_dict:
                logger.error("No hay clientes registrados en la base de datos.")
                raise ValueError("No hay clientes registrados en la base de datos.")

            template_path = get_resource_path('template/plantilla_Excel.xlsx')
            if not os.path.exists(template_path):
                logger.error(f"La plantilla de Excel no existe en la ruta: {template_path}")
                raise FileNotFoundError(f"La plantilla de Excel no existe en la ruta: {template_path}")

            procesados = []

            for client_id, client_data in clients_dict.items():
                nombre = client_data['nombre']
                aliases = client_data['aliases']

                if not aliases:
                    logger.warning(f"El cliente '{nombre}' no tiene aliases registrados. Se omitirá.")
                    continue  # Si el cliente no tiene aliases, continuar

                # Filtrar los datos que coincidan con cualquiera de los aliases
                df_cliente = df[df['Consignee'].isin(aliases)]

                if df_cliente.empty:
                    logger.info(f"No se encontraron datos para el cliente '{nombre}'.")
                    continue  # Si no hay datos, continuar

                output_file = os.path.join(self.output_dir, f'{nombre}.xlsx')

                # Crear el archivo Excel con formato
                _save_excel_with_format(df_cliente, template_path, output_file)
                logger.info(f"Reporte generado para el cliente '{nombre}': {output_file}")

                procesados.append(nombre)

            logger.info(f"Procesamiento completado para los clientes: {procesados}")
            return procesados

        except ValueError as ve:
            logger.error(f"Valor inválido: {ve}")
            raise
        except FileNotFoundError as fnf:
            logger.error(f"Archivo no encontrado: {fnf}")
            raise
        except Exception as e:
            logger.exception(f"Error al procesar el archivo Excel: {e}")
            raise

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
            logger.debug(f"Archivo encontrado para el cliente '{cliente}': {archivo}")
            return archivo
        else:
            logger.warning(f"No se encontró el archivo para el cliente '{cliente}'.")
            return None

    def close(self) -> None:
        """
        Cierra la conexión a la base de datos.
        """
        try:
            self.db.close()
            logger.debug("Conexión a la base de datos cerrada correctamente.")
        except Exception as e:
            logger.exception(f"Error al cerrar la conexión a la base de datos: {e}")
            raise
