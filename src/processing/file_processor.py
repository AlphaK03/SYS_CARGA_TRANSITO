# src/processing/file_processor.py

import os
import shutil
import sys
from tkinter import filedialog, messagebox
from datetime import datetime
import requests
import win32com.client
import pandas as pd
from openpyxl.styles import Font, PatternFill
from typing import List, Optional
from openpyxl import load_workbook
import win32com.client  # Importar para interactuar con Excel
import time
from urllib.parse import quote


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
    """
    Clase para procesar archivos Excel y generar reportes por cliente, y manejar la base de datos de clientes.
    """

    def __init__(self, site_url: str, sharepoint_file_url: str):
        """
        Inicializa una instancia de FileProcessor.

        Args:
            site_url (str): URL del sitio de SharePoint.
            sharepoint_file_url (str): Ruta relativa al archivo en SharePoint.
        """
        self.loading_frames = None  # Asegúrate de que loading_frames esté definido en __init__

        self.site_url = site_url
        self.sharepoint_file_url = sharepoint_file_url
        self.output_dir = os.path.join('data', 'reportes')
        self.db = ClientDatabase()
        self.config_manager = ConfigManager()
        self.ctx = None  # SharePoint context
        self.local_file_path = "data/temp/Informe de cargas Power BI.xlsm"  # Ruta al archivo descargado

    def run_macro(self):
        """
        Ejecuta la macro 'ActualizarPBI' en una nueva instancia de Excel y la cierra después de un retraso de 2 segundos.
        """
        if not self.local_file_path:
            raise Exception("El archivo no ha sido descargado aún.")

        try:
            # Crear una nueva instancia de Excel
            excel_app = win32com.client.DispatchEx("Excel.Application")
            excel_app.Visible = True  # Mantener Excel visible para depuración
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
            time.sleep(2)

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

        # Obtener la fecha actual en formato DD-MM-YYYY
        today_date = datetime.now().strftime("%d-%m-%Y")

        # Obtener la ruta del directorio raíz del proyecto (un nivel arriba de `src`)
        root_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))  # 📌 Sube un nivel desde `src`
        historico_dir = os.path.join(root_dir, "histórico")  # 📌 `historico/` estará al lado de `data/`

        # Crear la carpeta "histórico" si no existe
        os.makedirs(historico_dir, exist_ok=True)

        # Definir la ruta del archivo en la carpeta "histórico"
        historico_file_path = os.path.join(historico_dir, f"Informe de cargas Power BI_{today_date}.xlsm")

        # Guardar una copia del archivo original en "histórico"
        shutil.copy2(self.local_file_path, historico_file_path)

        # Leer el archivo Excel desde la hoja correspondiente
        df = pd.read_excel(self.local_file_path, sheet_name='Report SI SYSTEM - Customers', header=3)

        # Remover espacios en blanco de los nombres de las columnas
        df.columns = df.columns.str.strip()

        # Verificar si 'Consignee' está en las columnas
        if 'Consignee' not in df.columns:
            raise ValueError("La columna 'Consignee' no se encontró en el archivo.")

        # Obtener los clientes registrados junto con sus correos y aliases
        clients = self.db.get_all_clients()

        if not clients:
            raise ValueError("No hay clientes registrados en la base de datos.")

        os.makedirs(self.output_dir, exist_ok=True)

        # Ruta inicial de la plantilla general
        default_template_path = get_resource_path('template/plantilla_Excel.xlsx')

        # Mapeo de clientes y plantillas específicas
        special_clients = {
            'EPA': 'EPA Report',
            'ALPEMUSA': 'ALPEMUSA Report',
            'SAN NICOLAS': 'San Nicolas Report',
            'PLASTIMEX': 'Plastimex Report'
        }

        for client_id, nombre, correo, correo_cc, correo_bcc in clients:
            # Si el cliente tiene una hoja de reporte específica
            if nombre in special_clients:
                sheet_name = special_clients[nombre]
                try:
                    df_cliente = pd.read_excel(self.local_file_path, sheet_name=sheet_name, header=3)
                except ValueError:
                    raise ValueError(f"La hoja '{sheet_name}' no se encontró en el archivo Excel.")

                # Plantilla específica para el cliente
                template_path = get_resource_path(f'template/plantilla_{nombre.upper()}.xlsx')

                # Definir el archivo de salida
                output_file = os.path.join(self.output_dir, f'{nombre}.xlsx')

                # Guardar el archivo con formato
                _save_excel_with_format(df_cliente, template_path, output_file)
                continue  # Pasar al siguiente cliente

            # Filtrar los datos por los aliases del cliente
            aliases = self.db.get_aliases_for_client(client_id)
            alias_list = [alias for alias_id, alias in aliases]

            if not alias_list:
                continue  # Si el cliente no tiene aliases, continuar

            df_cliente = df[df['Consignee'].isin(alias_list)]

            if df_cliente.empty:
                continue  # Si no hay datos, continuar

            output_file = os.path.join(self.output_dir, f'{nombre}.xlsx')

            # Guardar el archivo con formato usando la plantilla general
            _save_excel_with_format(df_cliente, default_template_path, output_file)

        return [nombre for client_id, nombre, _, _, _ in clients]

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

    def get_last_update_date(self) -> Optional[str]:
        """
        Obtiene la fecha de la última actualización desde la hoja 'Hora Actualización' del Excel.

        Returns:
            Optional[str]: La fecha de la última actualización como cadena de texto.
        """
        if not self.local_file_path:
            raise Exception("El archivo no ha sido descargado aún.")

        try:
            # Cargar el libro de Excel
            workbook = load_workbook(filename=self.local_file_path, data_only=True)

            # Seleccionar la hoja 'Hora Actualización'
            if 'Hora Actualización' in workbook.sheetnames:
                sheet = workbook['Hora Actualización']

                # Buscar la columna con el encabezado 'Fecha de actualización'
                for col in sheet.iter_cols(min_row=1, max_row=1):
                    if col[0].value == 'Fecha de actualización':
                        # Obtener el valor en la fila 2 de esa columna
                        update_date = sheet.cell(row=2, column=col[0].column).value

                        if update_date:
                            # Guardar la fecha en config.json
                            self.config_manager.set_last_update_date(str(update_date))
                            return str(update_date)

            return None  # Si la hoja no tiene la fecha

        except Exception as e:
            print(f"Error al obtener la fecha de actualización: {e}")
            return None

    def download_file_with_token(self, access_token: str) -> str:
        """
        Descarga el archivo de SharePoint usando peticiones REST y un token OAuth2 (autenticación moderna).
        Si el archivo ya existe en la carpeta de destino y es válido, lo reutiliza.
        Si la descarga falla, permite al usuario seleccionar el archivo manualmente.
        """

        # 1) Crear carpeta local de destino si no existe
        os.makedirs('data/temp', exist_ok=True)

        # 2) Determinar el nombre del archivo local
        file_name = os.path.basename(self.sharepoint_file_url)
        self.local_file_path = os.path.join('data', 'temp', file_name)

        # 3) Verificar si el archivo ya existe y es válido
        if os.path.exists(self.local_file_path):
            if os.path.getsize(self.local_file_path) > 1024:  # 🔹 Verifica que el archivo tenga contenido (más de 1 KB)
                print(f"✅ Archivo ya existente encontrado en: {self.local_file_path}")
                return self.local_file_path  # Devolver el archivo sin descargarlo nuevamente
            else:
                print(f"⚠️ Archivo encontrado pero inválido (posible error previo en descarga). Eliminando...")
                os.remove(self.local_file_path)  # Eliminar archivo vacío o corrupto

        # 4) Escapar la URL para caracteres especiales
        server_relative_path = quote(self.sharepoint_file_url)
        endpoint = f"{self.site_url}/_api/web/GetFileByServerRelativeUrl('{server_relative_path}')/$value"

        # 5) Hacer la petición GET con Bearer token
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json;odata=nometadata"
        }
        response = requests.get(endpoint, headers=headers, stream=True)

        if response.status_code == 200:
            # 6) Guardar contenido en local_file_path
            with open(self.local_file_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            print(f"📥 Archivo descargado en: {self.local_file_path}")
            return self.local_file_path
        else:
            print(f"❌ Error al descargar: {response.status_code} - {response.text}")
            messagebox.showerror("Error de descarga", "No se pudo descargar el archivo de SharePoint.")

            # 7) Opción de permitir al usuario seleccionar el archivo manualmente
            if not os.path.exists(self.local_file_path):
                messagebox.showinfo("Seleccionar archivo",
                                    "El archivo no fue descargado. Por favor, selecciona el archivo localmente.")

                # Abre un cuadro de diálogo para que el usuario seleccione un archivo local
                selected_file_path = filedialog.askopenfilename(
                    title="Selecciona el archivo",
                    filetypes=[("Archivos Excel", "*.xlsx;*.xlsm"), ("Todos los archivos", "*.*")]
                )

                if selected_file_path:
                    try:
                        # Copiar el archivo seleccionado en la carpeta destino
                        shutil.copy(selected_file_path, self.local_file_path)
                        print(f"📂 Archivo copiado a: {self.local_file_path}")
                        return self.local_file_path
                    except Exception as e:
                        messagebox.showerror("Error", f"No se pudo copiar el archivo: {e}")
                        return None
                else:
                    messagebox.showwarning("Sin archivo seleccionado",
                                           "No se seleccionó ningún archivo. No se puede continuar.")
                    return None