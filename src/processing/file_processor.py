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
    import re
    from datetime import datetime, date
    from openpyxl.styles import Font, PatternFill
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo
    import pandas as pd
    import numpy as np
    import uuid

    HIDDEN_COLS = {"Consignee_norm", "__tmp_norm"}
    df_export = df_cliente.drop(columns=[c for c in HIDDEN_COLS if c in df_cliente.columns],
                                errors="ignore").copy()

    # --- Detectar columnas tipo fecha (acepta / o - y 2 o 4 dígitos de año) ---
    date_regex = re.compile(r"^\s*\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\s*/?\s*$")

    def looks_like_date_series(s: pd.Series) -> bool:
        if s.dropna().empty:
            return False
        as_str = s.dropna().astype(str).str.strip()
        return (as_str.str.match(date_regex)).mean() >= 0.6  # un poco más laxo

    date_cols = [c for c in df_export.columns if looks_like_date_series(df_export[c])]

    # --- Normalizador robusto para '9/8/25', '09-8-2025', '09/08/25/' -> '09/08/2025'
    def _normalize_date_text(txt: str, pivot_year: int = 49) -> str | None:
        if txt is None:
            return None
        s = str(txt).strip()
        s = re.sub(r"/\s*$", "", s)  # quitar barra final
        m = re.match(r"^\s*(\d{1,2})[/-](\d{1,2})[/-](\d{2}|\d{4})\s*$", s)
        if not m:
            return None
        d, mth, yr = m.groups()
        d = d.zfill(2)
        mth = mth.zfill(2)
        if len(yr) == 2:
            yy = int(yr)
            yr = str(2000 + yy) if yy <= pivot_year else str(1900 + yy)
        return f"{d}/{mth}/{yr}"

    # Intento de conversión vectorial; si no llega al umbral, igual mantengo la columna en date_cols
    for col in list(date_cols):
        s = df_export[col]
        if not pd.api.types.is_datetime64_any_dtype(s):
            norm = s.astype(str).map(lambda x: _normalize_date_text(x))
            parsed = pd.to_datetime(norm, dayfirst=True, errors="coerce")
            # Si parsea una proporción razonable, reemplazo la serie
            if parsed.notna().mean() >= 0.5:
                df_export[col] = parsed

    # --- Cargar plantilla ---
    wb = load_workbook(template_path)
    ws = wb.active

    # Limpiar tablas/auto-filtros previos del template (evita "reparaciones" de Excel)
    try:
        for t in list(ws._tables):
            ws._tables.remove(t)
    except Exception:
        pass
    ws.auto_filter.ref = None

    # Encabezados en fila 4
    for col_idx, name in enumerate(df_export.columns, start=1):
        cell = ws.cell(row=4, column=col_idx)
        cell.value = name
        cell.fill = PatternFill(start_color='63B9FF', end_color='63B9FF', fill_type='solid')
        cell.font = Font(bold=True, color='FFFFFF')

    # --- Helper: coerción por celda a datetime ---
    def _coerce_to_datetime(v):
        if v is None:
            return None
        if isinstance(v, (datetime,)):
            return v
        if isinstance(v, date):
            return datetime(v.year, v.month, v.day)
        if isinstance(v, pd.Timestamp):
            return v.to_pydatetime()
        if isinstance(v, np.datetime64):
            try:
                return pd.to_datetime(v).to_pydatetime()
            except Exception:
                return None
        # Texto -> normalizar y convertir
        s = _normalize_date_text(v)
        if s is None:
            return None
        try:
            return pd.to_datetime(s, dayfirst=True, errors="raise").to_pydatetime()
        except Exception:
            return None

    # --- Datos desde fila 5 (forzando fechas reales en columnas de fecha) ---
    excel_date_format = 'DD/MM/YYYY'
    start_row = 5
    for r, row in enumerate(df_export.itertuples(index=False, name=None), start=start_row):
        for c, val in enumerate(row, start=1):
            col_name = df_export.columns[c - 1]
            cell = ws.cell(row=r, column=c)  # no pasar value aquí

            if col_name in date_cols:
                dt = _coerce_to_datetime(val)
                if dt is not None:
                    cell.value = dt
                    cell.number_format = excel_date_format
                else:
                    # si no se pudo parsear, deja la celda vacía en vez de texto
                    cell.value = None
            else:
                cell.value = val

    # --- Tabla con filtros (si hay datos) ---
    n_rows = len(df_export)
    n_cols = len(df_export.columns)
    ws.freeze_panes = "A5"

    if n_cols > 0 and n_rows > 0:
        last_row = (start_row - 1) + n_rows
        last_col_letter = get_column_letter(n_cols)
        table_ref = f"A4:{last_col_letter}{last_row}"

        table_name = f"Tbl_{uuid.uuid4().hex[:8]}"
        tbl = Table(displayName=table_name, ref=table_ref)
        tbl.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )
        ws.add_table(tbl)

    wb.save(output_file)

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

        import unicodedata
        import pandas as pd

        def normalize_text(text):
            if not isinstance(text, str):
                return text
            text = unicodedata.normalize('NFKD', text)
            text = ''.join(c for c in text if not unicodedata.combining(c))
            return text.strip().upper()

        today_date = datetime.now().strftime("%d-%m-%Y")

        root_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
        historico_dir = os.path.join(root_dir, "data/histórico")
        os.makedirs(historico_dir, exist_ok=True)

        historico_file_path = os.path.join(historico_dir, f"Informe de cargas Power BI_{today_date}.xlsm")
        shutil.copy2(self.local_file_path, historico_file_path)

        df = pd.read_excel(self.local_file_path, sheet_name='Report SI SYSTEM - Customers', header=3)
        df.columns = df.columns.str.strip()

        if 'Consignee' not in df.columns:
            raise ValueError("La columna 'Consignee' no se encontró en el archivo.")

        # Normalizar columna Consignee
        df['Consignee_norm'] = df['Consignee'].apply(normalize_text)

        clients = self.db.get_all_clients()
        if not clients:
            raise ValueError("No hay clientes registrados en la base de datos.")

        os.makedirs(self.output_dir, exist_ok=True)

        default_template_path = get_resource_path('template/plantilla_Excel.xlsx')

        special_clients = {
            'OLO': 'OLO Report',
            'ALPEMUSA': 'ALPEMUSA Report',
            'SAN NICOLAS': 'San Nicolas Report',
            'PLASTIMEX': 'Plastimex Report'
        }

        # 🔸 Acumuladores para "Reporte Sin Reporte"
        aliases_usados = set()
        consignees_personalizados = set()

        for client_id, nombre, correo, correo_cc, correo_bcc in clients:
            if nombre in special_clients:
                sheet_name = special_clients[nombre]
                try:
                    df_cliente = pd.read_excel(self.local_file_path, sheet_name=sheet_name, header=3)
                    if 'Consignee' in df_cliente.columns:
                        df_cliente['Consignee_norm'] = df_cliente['Consignee'].apply(normalize_text)
                        consignees_personalizados.update(df_cliente['Consignee_norm'].dropna().unique())
                except ValueError:
                    raise ValueError(f"La hoja '{sheet_name}' no se encontró en el archivo Excel.")

                template_path = get_resource_path(f'template/plantilla_{nombre.upper()}.xlsx')
                output_file = os.path.join(self.output_dir, f'{nombre}.xlsx')

                _save_excel_with_format(df_cliente, template_path, output_file)
                continue

            aliases = self.db.get_aliases_for_client(client_id)
            alias_list = [normalize_text(alias) for alias_id, alias in aliases]

            if not alias_list:
                continue

            aliases_usados.update(alias_list)

            # Filtrar normalizado
            df_cliente = df[df['Consignee_norm'].isin(alias_list)]

            if df_cliente.empty:
                continue

            output_file = os.path.join(self.output_dir, f'{nombre}.xlsx')
            _save_excel_with_format(df_cliente, default_template_path, output_file)

        # 🔸 Crear "Reporte Sin Reporte"
        consignees_excluir = aliases_usados.union(consignees_personalizados)
        df_sin_reporte = df[~df['Consignee_norm'].isin(consignees_excluir)].copy()

        if not df_sin_reporte.empty:
            output_file = os.path.join(self.output_dir, 'Sin_Reporte.xlsx')
            _save_excel_with_format(df_sin_reporte, default_template_path, output_file)
            print("✅ Reporte 'Sin Reporte' generado.")
        else:
            print("ℹ️ No se encontraron registros para 'Reporte Sin Reporte'.")

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