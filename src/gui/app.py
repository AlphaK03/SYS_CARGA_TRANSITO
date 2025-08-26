# src/gui/app.py
import ctypes
import http.server
import json
import os
import shutil
import socketserver
import sqlite3
import sys
import time
from datetime import datetime
import threading
from typing import Optional
from urllib.parse import urlparse, parse_qs
import ttkbootstrap as ttk
import tkinter as tk
from tkinter import messagebox
import webbrowser
import requests
import logging
from PIL import Image, ImageTk, ImageSequence
from email_service.email_sender import EmailSender
from msal import PublicClientApplication

from processing.file_processor import FileProcessor
from database.client_database import ClientDatabase
from tkhtmlview import HTMLLabel
from utils.config_manager import ConfigManager

from utils.token_cache import SecureTokenCache

from manager.client_manager import open_client_manager, on_tab_change, on_list_select_por_lista, load_all_clients, \
    add_client_to_existing_list, on_list_select, load_clients_in_selected_list, on_client_select, on_all_client_select, \
    add_client, edit_client, delete_client, add_alias, delete_alias, on_closing, create_new_list, delete_selected_list, \
    update_list_combobox_ui, add_client_to_selected_list, remove_client_from_selected_list, export_clients, \
    import_clients

from manager.client_manager import context_menu_add_to_list, show_context_menu


from src.manager.client_manager import delete_client_from_context_menu, edit_alias_inline, undo_action, \
    send_reports_to_selected_clients, open_reports_in_excel, select_all_with_reports_in_all_clients, filter_all_clients

# Microsoft OAuth2 - ASW SAN JOSE
#MICROSOFT_CLIENT_ID = "3c6abf1d-1a49-4b55-9bf6-4ed5c3dcad51"
MICROSOFT_CLIENT_ID = "077e19a4-ab53-4e55-8286-ec02dc3cf36c"
MICROSOFT_TENANT_ID = "common"

# Establecer AppUserModelID
app_id = "CHC.Reportes"
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)

from gui.styles import init_styles
# Importar funciones y clases desde los módulos que creamos
from gui.helpers import (
    get_resource_path,
    get_microsoft_user_info
)


def open_user_manual_window():
    url = "https://alphak03.github.io/CHC_REPORTES_MANUAL-DE-USUARIO/"  # Reemplaza con el URL del manual de usuario
    webbrowser.open(url)


class Application(ttk.Window):
    def __init__(self):
        # Ajuste de DPI para Windows
        self.send_no_button = None
        self.send_yes_button = None
        self.frame_opciones = None
        if sys.platform.startswith('win'):
            try:
                from ctypes import windll
                windll.shcore.SetProcessDpiAwareness(1)
            except Exception as e:
                print(f"No se pudo ajustar la conciencia de DPI: {e}")

        super().__init__(themename="flatly")  # Mantener el tema
        self.title("CHC REPORTES")

        # 📌 Ruta del icono
        if getattr(sys, 'frozen', False):
            # Si el programa está compilado en un .exe con PyInstaller
            base_path = sys._MEIPASS
        else:
            # Si se ejecuta en el entorno de desarrollo
            base_path = os.path.abspath(".")

        icon_path = os.path.join(base_path, "resources", "icono.ico")

        # 📌 Establecer el icono de la ventana
        if os.path.exists(icon_path):
            self.iconbitmap(icon_path)
        else:
            print(f"⚠️ Advertencia: No se encontró el icono en {icon_path}")

        # Inicializar la etiqueta de progreso usando grid
        self.progress_label = ttk.Label(self, text="", style='Info.TLabel')
        self.progress_label.grid(row=0, column=0, padx=10, pady=10, sticky='w')

        # Ajuste del factor de escalado
        self.scaling_factor = self.get_scaling_factor()
        self.tk.call('tk', 'scaling', self.scaling_factor)

        # Ajustar el tamaño inicial de la ventana según el escalado
        initial_width = int(500 * self.scaling_factor)
        initial_height = int(400 * self.scaling_factor)

        self.center_window(self, initial_width, initial_height)


        self.processor = None
        self.db = ClientDatabase()
        self.config_manager = ConfigManager()

        self.configure(background='white')  # Fondo blanco para la ventana

        # Crear cachés de tokens separados
        # Crear cachés de tokens separados utilizando SecureTokenCache
        self.email_token_cache = SecureTokenCache(cache_name='email_token_cache')
        self.sharepoint_token_cache = SecureTokenCache(cache_name='sharepoint_token_cache')

        # Configurar MSAL para la cuenta de origen (correo) con "common"
        self.email_msal_app = PublicClientApplication(
            client_id=MICROSOFT_CLIENT_ID,
            authority='https://login.microsoftonline.com/common',
            token_cache=self.email_token_cache.cache  # Usar el caché específico para correo
        )

        # Configurar MSAL para SharePoint con el TENANT_ID específico
        self.sharepoint_msal_app = PublicClientApplication(
            client_id=MICROSOFT_CLIENT_ID,
            authority=f'https://login.microsoftonline.com/{MICROSOFT_TENANT_ID}',
            token_cache=self.sharepoint_token_cache.cache  # Usar el caché específico para SharePoint
        )

        # Bindings globales para atajos de teclado
        self.bind_all("<Control-a>", lambda event: self.add_client())
        self.bind_all("<Control-A>", lambda event: self.add_client())
        self.bind_all("<Control-e>", lambda event: self.edit_client())
        self.bind_all("<Control-E>", lambda event: self.edit_client())
        self.bind_all("<Control-d>", lambda event: self.delete_client())
        self.bind_all("<Control-D>", lambda event: self.delete_client())
        self.bind_all("<Control-Shift-l>", lambda event: self.add_client_to_existing_list())
        self.bind_all("<Control-Shift-L>", lambda event: self.add_client_to_existing_list())

        self.undo_stack = []

        # Configuración para FileProcessor (en app.py o donde crees tu ventana principal)
        self.processor = FileProcessor(
            site_url=self.config_manager.get_sharepoint_site_url(),
            sharepoint_file_url=self.config_manager.get_sharepoint_file_url()
        )

        # Inicializar EmailSender pasando la instancia de msal_app
        self.email_sender = EmailSender(self.email_msal_app)

        init_styles()  # Inicializar estilos personalizados
        self.create_widgets()  # Crear widgets

    def get_scaling_factor(self):
        # Obtener el factor de escalado de DPI
        if sys.platform.startswith('win'):
            from ctypes import windll
            user32 = windll.user32
            user32.SetProcessDPIAware()
            dpi = user32.GetDpiForWindow(self.winfo_id())
            scaling = dpi / 96  # 96 DPI es el valor por defecto
            return scaling
        else:
            return self.tk.call('tk', 'scaling')

    def get_email_access_token(self):
        """Obtiene el token para la cuenta de origen (correo) usando el tenant 'common'."""
        scopes = ['Mail.Send']  # Ajusta el scope según sea necesario
        accounts = self.email_msal_app.get_accounts()
        result = None

        if accounts:
            result = self.email_msal_app.acquire_token_silent(scopes, account=accounts[0])

        if not result or 'access_token' not in result:
            result = self.email_msal_app.acquire_token_interactive(scopes=scopes, prompt='select_account')
            self.email_token_cache.save(self.email_msal_app)  # Guarda el token en caché después de autenticar

        if "access_token" in result:
            return result['access_token']
        else:
            error_message = result.get('error_description', 'No se pudo obtener el token de acceso para el correo.')
            messagebox.showerror("Error", f"Error al adquirir el token de acceso para el correo: {error_message}")
            return None

    def get_sharepoint_access_token(self):
        """
        Obtiene el token de acceso para SharePoint con persistencia en caché.
        Evita bloqueos si el usuario cierra la pestaña de autenticación sin iniciar sesión.
        """
        try:
            # Leer site_url desde configuración
            raw_site_url = self.config_manager.get_sharepoint_site_url()

            # Extraer dominio base de SharePoint
            from urllib.parse import urlparse
            parsed_url = urlparse(raw_site_url)
            base_domain = f"{parsed_url.scheme}://{parsed_url.netloc}"

            # Definir scopes correctamente
            scopes = [f"{base_domain}/.default"]

            # Intentar obtener el token en modo silencioso
            accounts = self.sharepoint_msal_app.get_accounts()
            result = None

            if accounts:
                result = self.sharepoint_msal_app.acquire_token_silent(scopes, account=accounts[0])

            if result and "access_token" in result:
                print("✅ Token obtenido de caché")
                return result["access_token"]

            print("🔄 No se encontró token en caché. Iniciando autenticación interactiva...")

            # Autenticación interactiva en un hilo separado
            def acquire_token():
                nonlocal result
                try:
                    result = self.sharepoint_msal_app.acquire_token_interactive(scopes=scopes, prompt='select_account')
                except Exception as e:
                    result = {'error_description': str(e)}

            auth_thread = threading.Thread(target=acquire_token, daemon=True)
            auth_thread.start()

            # Esperar con timeout de 35 segundos
            start_time = time.time()
            while auth_thread.is_alive():
                if time.time() - start_time > 35:
                    messagebox.showwarning("Tiempo agotado", "La autenticación tardó demasiado. Inténtalo de nuevo.")
                    return None
                time.sleep(1)

                # Si se obtuvo un token, almacenarlo en caché
            if result and "access_token" in result:
                self.sharepoint_token_cache.save(
                    self.sharepoint_msal_app)  # 📌 Guardar token en caché después de autenticación
                print("✅ Token guardado en caché correctamente")
                return result["access_token"]
            else:
                error_message = result.get("error_description",
                                           "No se pudo obtener el token de acceso para SharePoint.")
                messagebox.showerror("Error", f"Error al adquirir el token de acceso: {error_message}")
                return None

        except Exception as e:
            messagebox.showerror("Error", f"Error inesperado al obtener el token de SharePoint: {str(e)}")
            return None

    def create_widgets(self):

        # Ajustar la fuente del menú
        menu_font = ("Helvetica", 14)  # Aumentar tamaño de letra
        # Crear la barra de menú
        menu_bar = tk.Menu(self, font=menu_font)
        self.config(menu=menu_bar)

        #
        # Menú "Configuración"
        #
        config_menu = tk.Menu(menu_bar, tearoff=0, font=menu_font)

        # Sección relacionada con el correo
        config_menu.add_command(
            label="Configurar Cuenta de Envío (Correo)",
            command=self.open_config_window
        )
        config_menu.add_command(
            label="Configurar Plantilla de Correo",
            command=self.open_email_template_settings_window
        )
        config_menu.add_separator()  # Línea separadora

        config_menu.add_command(
            label="Configurar URL de Archivo en SharePoint",
            command=self.open_sharepoint_file_settings_window
        )
        config_menu.add_separator()  # Otra línea separadora

        # Opción para salir
        config_menu.add_command(
            label="Salir",
            command=self.on_closing
        )

        menu_bar.add_cascade(label="Configuración", menu=config_menu)

        #
        # Menú "Gestión"
        #
        gestion_menu = tk.Menu(menu_bar, tearoff=0, font=menu_font)
        gestion_menu.add_command(label="Gestor de Clientes", command=self.open_client_manager)
        gestion_menu.add_command(label="Importar Clientes (JSON)", command=self.import_clients)
        gestion_menu.add_command(label="Exportar Clientes (JSON)", command=self.export_clients)
        menu_bar.add_cascade(label="Gestión", menu=gestion_menu)

        #
        # Menú "Ayuda"
        #
        ayuda_menu = tk.Menu(menu_bar, tearoff=0, font=menu_font)
        ayuda_menu.add_command(label="Acerca de", command=self.open_about_window)
        menu_bar.add_cascade(label="Ayuda", menu=ayuda_menu)

        # Configuración del grid de la ventana principal
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self.rowconfigure(1, weight=0)
        self.rowconfigure(2, weight=1)
        self.rowconfigure(3, weight=0)
        self.rowconfigure(4, weight=1)
        self.rowconfigure(5, weight=0)
        self.rowconfigure(6, weight=0)  # Ajustado para el pie de página

        # Ruta de la imagen
        logo_path = get_resource_path('resources/Logo_CHC_new.png')
        print(f"Ruta del logo: {logo_path}")  # Añadir esta línea para depuración

        # Crear el label para el logo por adelantado
        self.logo_label = tk.Label(self, bg='white')
        self.logo_label.grid(row=0, column=0, padx=10, pady=(10, 5), sticky='n')

        # Cargar y redimensionar la imagen del logo si existe
        if os.path.exists(logo_path):
            try:
                # Abrir la imagen original
                original_logo_image = Image.open(logo_path)
                desired_width = 120  # Aumentado de 300 a 400 px

                # Determinar el filtro de resampling basado en la versión de Pillow
                try:
                    # Pillow >= 10.0.0
                    resample = Image.Resampling.LANCZOS
                except AttributeError:
                    # Pillow < 10.0.0
                    try:
                        resample = Image.LANCZOS
                    except AttributeError:
                        # Pillow < 10.0.0 sin LANCZOS, usar BICUBIC
                        resample = Image.BICUBIC

                # Mantener la relación de aspecto
                w_percent = (desired_width / float(original_logo_image.size[0]))
                h_size = int((float(original_logo_image.size[1]) * float(w_percent)))
                resized_image = original_logo_image.resize((desired_width, h_size), resample)

                # Convertir la imagen a PhotoImage
                self.logo_photo = ImageTk.PhotoImage(resized_image)

                # Configurar el label de la imagen
                self.logo_label.config(image=self.logo_photo)
                self.logo_label.image = self.logo_photo  # Mantener la referencia para evitar liberación

            except Exception as e:
                logging.error(f"Error al cargar la imagen: {e}")
                print(f"Error al cargar la imagen: {e}")
        else:
            logging.warning(f"Imagen no encontrada: {logo_path}")
            print(f"Imagen no encontrada: {logo_path}")

        # Etiqueta de texto "CHC CONSULTORES" debajo de la imagen con fuente grande
        consultores_label = ttk.Label(
            self,
            text="CHC CONSULTORES",
            style='Custom.TLabel',
            font=("Helvetica", 13)
        )
        consultores_label.grid(row=1, column=0, padx=10, pady=(0, 10), sticky='n')

        # Título principal
        titulo_principal = ttk.Label(
            self,
            text="Gestor de Reportes",
            style='Custom.TLabel',
            font=("Helvetica", 25, "bold")
        )
        titulo_principal.grid(row=2, column=0, padx=10, pady=10, sticky='n')

        # Frame de los botones
        frame_botones = tk.Frame(self, bg='white')
        frame_botones.grid(row=3, column=0, padx=10, pady=10, sticky='ew')

        # Reajustar para colocar el engranaje
        frame_botones.columnconfigure(0, weight=1)  # Espacio a la izquierda
        frame_botones.columnconfigure(1, weight=2)  # Botón principal
        frame_botones.columnconfigure(2, weight=0)  # Botón engranaje
        frame_botones.columnconfigure(3, weight=1)  # Espacio a la derecha

        # Botón para generar reportes (centrado)
        boton_generar = ttk.Button(
            frame_botones,
            text="Generar Reportes",
            command=self.filter_and_generate,
            style='BlueZodiac.TButton'
        )
        boton_generar.grid(row=0, column=1, padx=5, pady=5, sticky='ew')

        # Botón para generar y enviar reportes automáticamente
        boton_generar_enviar = ttk.Button(
            frame_botones,
            text="Generar y Enviar (Automático)",
            command=self.filter_generate_and_send,
            style='GreenSuccess.TButton'
        )
        boton_generar_enviar.grid(row=1, column=1, padx=5, pady=5, sticky='ew')

        # Agregar un botón de engranaje al lado del botón "Generar y Enviar (Automático)"
        gear_icon_path = get_resource_path('resources/gear_icon_b.png')  # Ajusta la ruta al ícono
        if os.path.exists(gear_icon_path):
            try:
                gear_image = Image.open(gear_icon_path).resize((20, 20), Image.Resampling.LANCZOS)
                self.gear_icon = ImageTk.PhotoImage(gear_image)
            except Exception as e:
                print(f"Error al cargar ícono de engranaje: {e}")
                self.gear_icon = None
        else:
            self.gear_icon = None

        boton_configurar_clientes = ttk.Button(
            frame_botones,
            command=self.open_client_selection_window,  # Tu función para abrir la ventana de checks
            style='Toolbutton'
        )

        if self.gear_icon:
            boton_configurar_clientes.config(image=self.gear_icon)
        else:
            # Si no hay imagen, usar un emoji
            boton_configurar_clientes.config(text="⚙")

        # Ubicar engranaje al lado del botón "Generar y Enviar (Automático)" (misma fila: 1, columna: 2)
        boton_configurar_clientes.grid(row=1, column=2, padx=(0, 5), pady=5, sticky='w')

        # Label para mostrar mensajes de progreso
        self.progress_label = ttk.Label(
            self,
            text="",
            style='Custom.TLabel'
        )
        self.progress_label.grid(row=4, column=0, padx=10, pady=5, sticky='n')

        # Cargar el GIF de carga
        loading_gif_path = get_resource_path('resources/loading.gif')
        if os.path.exists(loading_gif_path):
            try:
                self.loading_frames = [
                    ImageTk.PhotoImage(img)
                    for img in ImageSequence.Iterator(Image.open(loading_gif_path))
                ]
                self.loading_label = tk.Label(self, bg='white')
                self.loading_label.grid(row=5, column=0, padx=10, pady=5, sticky='n')
                self.loading_label.grid_remove()  # Ocultar el label inicialmente
            except Exception as e:
                logging.error(f"Error al cargar el GIF de carga: {e}")
                print(f"Error al cargar el GIF de carga: {e}")
        else:
            self.loading_frames = None
            self.loading_label = None

        # Pie de página
        footer = ttk.Label(
            self,
            text="© 2024 CHC-Consultores. Todos los derechos reservados.",
            style='Custom.TLabel'
        )
        footer.grid(row=6, column=0, padx=10, pady=10, sticky='s')

        # Configurar el evento de redimensionamiento

    def get_current_authenticated_user(self):
        """
        Obtiene la información del usuario actualmente autenticado.
        Retorna un diccionario con 'email' y 'name' si hay un usuario autenticado,
        o None si no hay ninguno.
        """
        # Log para indicar el inicio del proceso de obtención de información de usuario
        logging.info("Iniciando el proceso para obtener información del usuario autenticado.")

        accounts = self.email_msal_app.get_accounts()
        if not accounts:
            logging.warning("No hay cuentas autenticadas en MSAL.")
            return None  # No hay cuentas autenticadas

        # Intentar adquirir un token silencioso para obtener la información del usuario
        scopes = ['User.Read']  # Asegúrate de que este scope está permitido
        logging.info("Intentando adquirir un token de acceso en silencio.")
        result = self.email_msal_app.acquire_token_silent(scopes, account=accounts[0])

        if not result or 'access_token' not in result:
            logging.warning(
                "No se pudo obtener un token de acceso silencioso. Es posible que el usuario necesite autenticarse nuevamente.")
            return None

        access_token = result['access_token']
        logging.info("Token de acceso obtenido exitosamente.")

        # Llamar a Microsoft Graph API para obtener información del usuario
        try:
            headers = {
                'Authorization': f'Bearer {access_token}'
            }
            logging.info("Realizando solicitud a Microsoft Graph para obtener la información del usuario.")
            response = requests.get('https://graph.microsoft.com/v1.0/me', headers=headers)
            response.raise_for_status()  # Esto lanzará una excepción en caso de errores HTTP

            user_info = response.json()
            email = user_info.get('mail') or user_info.get('userPrincipalName')
            name = user_info.get('displayName')

            if email and name:
                logging.info(f"Información del usuario obtenida: email={email}, name={name}")
            else:
                logging.warning(
                    "La respuesta de Microsoft Graph no contiene la información esperada (email o name están ausentes).")

            return {'email': email, 'name': name}
        except requests.exceptions.HTTPError as http_err:
            logging.error(f"Error HTTP al obtener información del usuario de Microsoft Graph: {http_err}")
            logging.error(f"Detalle de la respuesta: {response.text if response else 'No hay respuesta'}")
        except Exception as e:
            logging.error(f"Error al obtener información del usuario de Microsoft Graph: {e}")
            logging.error(f"Detalle de la respuesta: {response.text if response else 'No hay respuesta'}")
        return None

    def open_about_window(self):
        about_window = tk.Toplevel(self)
        about_window.title("Acerca de")

        # Ajustar el tamaño de la ventana según la resolución de pantalla
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        scaling_factor = min(screen_width / 1920, screen_height / 1080)
        window_width = min(int(400 * scaling_factor), 600)
        window_height = min(int(300 * scaling_factor), 400)
        about_window.geometry(f"{window_width}x{window_height}")

        # Centrando la ventana en la pantalla
        x = int((screen_width / 2) - (window_width / 2))
        y = int((screen_height / 2) - (window_height / 2))
        about_window.geometry(f"+{x}+{y}")
        about_window.configure(background='white')

        # Información de la aplicación
        app_name = "Gestor de Reportes"
        version = "Versión 2.4.3"
        company = "CHC-Consultores"

        # Etiqueta de nombre de la aplicación con tamaño de fuente ajustado
        label_app_name = ttk.Label(
            about_window,
            text=app_name,
            style='Custom.TLabel'
        )
        label_app_name.pack(pady=int(10 * scaling_factor))

        # Etiqueta de versión con tamaño de fuente ajustado
        label_version = ttk.Label(
            about_window,
            text=version,
            style='Custom.TLabel'
        )
        label_version.pack(pady=int(5 * scaling_factor))

        # Etiqueta de empresa con tamaño de fuente ajustado
        label_company = ttk.Label(
            about_window,
            text=f"Desarrollado por {company}",
            style='Custom.TLabel'
        )
        label_company.pack(pady=int(5 * scaling_factor))

        # Frame para centrar los botones sin estirarlos
        button_frame = ttk.Frame(about_window)
        button_frame.pack(pady=int(20 * scaling_factor))

        # Botón para abrir las políticas y términos en una nueva ventana
        boton_privacidad = ttk.Button(
            button_frame,
            text="Política de Privacidad y Términos de Uso",
            command=self.open_privacy_terms_window,
            style='BlueZodiac.TButton'
        )
        boton_privacidad.pack(pady=(0, int(10 * scaling_factor)))

        # Botón para abrir el manual de usuario en una nueva ventana
        boton_manual_usuario = ttk.Button(
            button_frame,
            text="Manual de Usuario",
            command=open_user_manual_window,  # Asegúrate de tener esta función definida
            style='BlueZodiac.TButton'
        )
        boton_manual_usuario.pack()

        # Hacer que la ventana sea modal
        about_window.transient(self)
        about_window.grab_set()
        self.wait_window(about_window)

    def open_privacy_terms_window(self):
        privacy_window = tk.Toplevel(self)
        privacy_window.title("Política de Privacidad y Términos de Uso")

        # Ajustar el tamaño de la ventana según la resolución de pantalla
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        scaling_factor = min(screen_width / 1920, screen_height / 1080)

        window_width = min(int(screen_width * 0.6), 1500)
        window_height = min(int(screen_height * 0.8), 1000)

        # Calcular la posición x e y para centrar la ventana
        x = int((screen_width / 2) - (window_width / 2))
        y = int((screen_height / 2) - (window_height / 2))

        # Establecer la geometría de la ventana incluyendo la posición
        privacy_window.geometry(f"{window_width}x{window_height}+{x}+{y}")
        privacy_window.configure(background='white')

        # Título centrado
        title_label = ttk.Label(
            privacy_window,
            text="Política de Privacidad y Términos de Uso",
            anchor="center",
            style='Title.TLabel'
        )
        title_label.pack(pady=int(10 * scaling_factor))

        # Obtener HTML de la página de GitHub
        try:
            response = requests.get("https://alphak03.github.io/CHC-REPORTES---POLITICAS-DE-PRIVACIDAD/")
            response.raise_for_status()
            content_html = response.text
        except requests.RequestException as e:
            content_html = f"<p>Error al cargar las políticas y términos: {e}</p>"

        # Crear un HTMLLabel para mostrar el contenido HTML
        html_label = HTMLLabel(privacy_window, html=content_html)
        html_label.pack(fill="both", expand=True, padx=int(10 * scaling_factor), pady=int(10 * scaling_factor))

        # Frame para centrar el botón de cierre
        button_frame = ttk.Frame(privacy_window)
        button_frame.pack(pady=int(10 * scaling_factor))

        # Botón de cierre centrado
        close_button = ttk.Button(
            button_frame,
            text="Cerrar",
            command=privacy_window.destroy,
            style='RedColor.TButton'
        )
        close_button.pack()

        privacy_window.transient(self)
        privacy_window.grab_set()
        self.wait_window(privacy_window)

    def update_progress(self, message, style='Info.TLabel'):
        """
        Actualiza el texto de la etiqueta de progreso con el mensaje y el estilo proporcionado.

        Args:
            message (str): Mensaje a mostrar.
            style (str): Estilo de la etiqueta. Valores posibles: 'Info.TLabel', 'Success.TLabel', etc.
        """
        self.progress_label.config(style=style, text=message)
        self.update_idletasks()  # Actualiza la interfaz para reflejar el cambio inmediatamente

    def open_config_window(self):
        """
        Configuración de Correo y Autenticación, solo Microsoft
        """
        try:
            config_window = tk.Toplevel(self)
            config_window.title("Configuración de Correo y Autenticación")
            config_window.configure(background='#f0f0f0')

            screen_width = config_window.winfo_screenwidth()
            screen_height = config_window.winfo_screenheight()
            scaling_factor = min(screen_width / 1920, screen_height / 1080)

            window_width = min(int(screen_width * 0.3), 1200)
            window_height = min(int(screen_height * 0.4), 900)
            self.center_window(config_window, window_width, window_height)

            config_window.columnconfigure(0, weight=1)

            auth_frame = ttk.Frame(config_window, padding=(20, 10))
            auth_frame.grid(row=0, column=0, sticky='nsew')
            auth_frame.columnconfigure(0, weight=1)

            ttk.Label(
                auth_frame,
                text="Información de Autenticación de la Cuenta",
                style='CustomTitle.TLabel'
            ).grid(row=0, column=0, pady=(0, 10), sticky='w')

            ttk.Separator(auth_frame, orient='horizontal').grid(row=1, column=0, sticky='ew', pady=(0, 10))

            self.etiqueta_email_actual = ttk.Label(auth_frame, text="Cargando información...",
                                                   style='CustomValue.TLabel')
            self.etiqueta_email_actual.grid(row=2, column=0, padx=5, pady=5, sticky='w')

            self.etiqueta_nombre_value = ttk.Label(auth_frame, text="", style='CustomValue.TLabel')
            self.etiqueta_nombre_value.grid(row=3, column=0, padx=5, pady=5, sticky='w')

            # 🔥 Mostrar Alias Predeterminado Actual
            default_alias = self.config_manager.get_default_sender()
            if default_alias:
                alias_text = f"Alias Predeterminado: {default_alias}"
            else:
                alias_text = "Alias Predeterminado: No configurado"

            self.etiqueta_alias_predeterminado = ttk.Label(auth_frame, text=alias_text, style='CustomValue.TLabel')
            self.etiqueta_alias_predeterminado.grid(row=4, column=0, padx=5, pady=5, sticky='w')

            # Cargar información del usuario
            user_info = self.get_current_authenticated_user()
            if user_info:
                email_autenticado = user_info['email']
                nombre_usuario = user_info['name']
                self.etiqueta_email_actual.config(text=f"Cuenta Autenticada: {email_autenticado}")
                self.etiqueta_nombre_value.config(text=f"Usuario: {nombre_usuario}")
            else:
                self.etiqueta_email_actual.config(text="NO AUTENTICADO")
                self.etiqueta_nombre_value.config(text="")

            button_frame = ttk.Frame(config_window, padding=(20, 10))
            button_frame.grid(row=2, column=0, sticky='ew')
            button_frame.columnconfigure(0, weight=1)

            # -- ESTILOS
            style = ttk.Style()
            style.configure('CustomTitle.TLabel', font=('Helvetica', 16, 'bold'), foreground='#333333')
            style.configure('CustomValue.TLabel', font=('Helvetica', 12), foreground='#000000')
            style.configure('Warning.TButton', foreground='white',
                            background='#FFA500')  # Color naranja para el botón warning

            # Funciones de botones
            def iniciar_oauth2():
                try:
                    tenant_id = "common"
                    authority = f'https://login.microsoftonline.com/{tenant_id}'
                    redirect_uri = "http://localhost:62190/"
                    scopes = EmailSender.MS_SCOPES

                    app = self.email_msal_app
                    result = app.acquire_token_silent(scopes, account=None)

                    if not result:
                        auth_url = app.get_authorization_request_url(
                            scopes=scopes,
                            redirect_uri=redirect_uri,
                            response_mode='query',
                            prompt='select_account'
                        )
                        webbrowser.open(auth_url)
                        auth_code = []

                        class AuthHandler(http.server.SimpleHTTPRequestHandler):
                            def __init__(self, *args, auth_code_container, **kwargs):
                                self.auth_code_container = auth_code_container
                                super().__init__(*args, **kwargs)

                            def do_GET(self):
                                parsed_path = urlparse(self.path)
                                query = parse_qs(parsed_path.query)
                                if 'code' in query:
                                    self.auth_code_container.append(query['code'][0])
                                    html_content = '<html><body><h1>Autenticación exitosa. Puedes cerrar esta ventana.</h1></body></html>'
                                    self.send_response(200)
                                    self.send_header('Content-type', 'text/html; charset=utf-8')
                                    self.end_headers()
                                    self.wfile.write(html_content.encode('utf-8'))

                        with socketserver.TCPServer(("", 62190), lambda *args, **kwargs: AuthHandler(*args,
                                                                                                     auth_code_container=auth_code,
                                                                                                     **kwargs)) as httpd:
                            server_thread = threading.Thread(target=httpd.handle_request)
                            server_thread.start()
                            server_thread.join()

                        if not auth_code:
                            messagebox.showerror("Error", "No se recibió el código de autorización de Microsoft.")
                            return

                        result = app.acquire_token_by_authorization_code(auth_code[0], scopes=scopes,
                                                                         redirect_uri=redirect_uri)
                        self.email_token_cache.save(app)

                    if "access_token" in result:
                        user_info = self.get_current_authenticated_user()
                        if user_info:
                            email_autenticado = user_info['email']
                            nombre_usuario = user_info['name']
                            self.config_manager.set_email(email_autenticado)
                            self.etiqueta_email_actual.config(text=f"Cuenta autenticada: {email_autenticado}")
                            self.etiqueta_nombre_value.config(text=f"Usuario: {nombre_usuario}")
                            messagebox.showinfo("Éxito", "Autenticación de Microsoft exitosa.")
                        else:
                            messagebox.showerror("Error", "No se pudo obtener la información del usuario.")
                    else:
                        error_description = result.get('error_description', 'No se pudo obtener el token de acceso.')
                        messagebox.showerror("Error", f"Error al adquirir el token de acceso: {error_description}")

                except Exception as e:
                    messagebox.showerror("Error", f"Error durante la autenticación: {str(e)}")

            def cerrar_sesion():
                try:
                    self.config_manager.clear_credentials()
                    self.email_token_cache.clear_cache()
                    self.sharepoint_token_cache.clear_cache()
                    self.etiqueta_email_actual.config(text="NO AUTENTICADO")
                    self.etiqueta_nombre_value.config(text="")
                    messagebox.showinfo("Cierre de sesión", "Sesión cerrada. Por favor, vuelve a iniciar sesión.")
                    config_window.destroy()
                    self.open_config_window()
                except Exception as e:
                    messagebox.showerror("Error", f"No se pudo cerrar sesión: {str(e)}")

            def borrar_alias_predeterminado():
                confirm = messagebox.askyesno("Confirmar", "¿Deseas borrar el alias predeterminado?")
                if confirm:
                    self.config_manager.clear_default_sender()
                    self.etiqueta_alias_predeterminado.config(text="Alias Predeterminado: No configurado")
                    messagebox.showinfo("Éxito", "Alias predeterminado eliminado.")

            def cambiar_alias_predeterminado():
                nuevo_alias = self.select_sender_alias()
                if nuevo_alias:
                    self.etiqueta_alias_predeterminado.config(text=f"Alias Predeterminado: {nuevo_alias}")
                    messagebox.showinfo("Éxito", f"Alias predeterminado cambiado a: {nuevo_alias}")

            # -- BOTONES
            self.boton_autenticar = ttk.Button(
                button_frame,
                text="Iniciar Sesión con Microsoft",
                command=iniciar_oauth2,
                style='Success.TButton',
                width=25
            )
            self.boton_autenticar.grid(row=0, column=0, pady=(0, 10), sticky='ew')

            self.boton_cerrar_sesion = ttk.Button(
                button_frame,
                text="Cerrar Sesión",
                command=cerrar_sesion,
                style='Danger.TButton',
                width=25
            )
            self.boton_cerrar_sesion.grid(row=1, column=0, pady=(0, 10), sticky='ew')

            boton_borrar_alias = ttk.Button(
                button_frame,
                text="Borrar Alias Predeterminado",
                command=borrar_alias_predeterminado,
                style='Warning.TButton',
                width=25
            )
            boton_borrar_alias.grid(row=2, column=0, pady=(0, 10), sticky='ew')

            boton_cambiar_alias = ttk.Button(
                button_frame,
                text="Cambiar Alias Predeterminado",
                command=cambiar_alias_predeterminado,
                style='BlueZodiac.TButton',
                width=25
            )
            boton_cambiar_alias.grid(row=3, column=0, pady=(0, 10), sticky='ew')

            config_window.transient(self)
            config_window.grab_set()
            self.wait_window(config_window)

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir la ventana de configuración: {str(e)}")

    def open_email_template_settings_window(self):
        """
        Abre una ventana para configurar la plantilla del cuerpo de correo.
        El usuario puede usar {CLIENT_NAME} y {LAST_UPDATE_DATE} como marcadores de posición.
        """
        config_window = tk.Toplevel(self)
        config_window.title("Configurar Plantilla de Correo")

        # Calcular tamaño y centrar
        screen_width = config_window.winfo_screenwidth()
        screen_height = config_window.winfo_screenheight()
        scaling_factor = min(screen_width / 1920, screen_height / 1080)
        window_width = int(700 * scaling_factor)
        window_height = int(500 * scaling_factor)
        self.center_window(config_window, window_width, window_height)

        frame = ttk.Frame(config_window, padding=10)
        frame.pack(expand=True, fill="both")

        # Frame superior para botones
        top_frame = ttk.Frame(frame)
        top_frame.pack(fill='x', pady=(0, 10))

        # Botón de guardar
        def save_template():
            template = text_box.get("1.0", tk.END).strip()
            self.config_manager.set_email_body_template(template)
            messagebox.showinfo("Éxito", "Plantilla de correo guardada exitosamente.")
            config_window.destroy()

        save_button = ttk.Button(
            top_frame,
            text="Guardar",
            command=save_template,
            style='GreenSuccess.TButton'
        )
        save_button.pack(side='right', padx=5)

        # Botón para insertar {CLIENT_NAME}
        def insert_client_placeholder():
            text_box.insert(tk.INSERT, "{CLIENT_NAME}")

        insert_placeholder_button = ttk.Button(
            top_frame,
            text="Insertar {CLIENT_NAME}",
            command=insert_client_placeholder,
            style='BlueZodiac.TButton'
        )
        insert_placeholder_button.pack(side='left', padx=5)

        # Botón para insertar texto en negrita
        def insert_bold_tags():
            text_box.insert(tk.INSERT, "<b></b>")
            text_box.mark_set("insert", f"{text_box.index('insert')}-4c")

        bold_button = ttk.Button(
            top_frame,
            text="Negrita (<b>)",
            command=insert_bold_tags,
            style='BlueZodiac.TButton'
        )
        bold_button.pack(side='left', padx=5)

        # Botón para insertar salto de línea <br>
        def insert_br_tag():
            text_box.insert(tk.INSERT, "<br>")

        br_button = ttk.Button(
            top_frame,
            text="Salto de línea (<br>)",
            command=insert_br_tag,
            style='BlueZodiac.TButton'
        )
        br_button.pack(side='left', padx=5)

        # Botón para insertar {LAST_UPDATE_DATE} en la plantilla
        def insert_last_update_placeholder():
            text_box.insert(tk.INSERT, "{LAST_UPDATE_DATE}")

        last_update_button = ttk.Button(
            top_frame,
            text="Insertar Última Fecha",
            command=insert_last_update_placeholder,
            style='YellowWarning.TButton'  # 🔹 Usa un estilo amarillo para destacar este botón
        )
        last_update_button.pack(side='left', padx=5)

        # Explicación general
        instrucciones_label = ttk.Label(
            frame,
            text=(
                "Personaliza el cuerpo del correo. Puedes usar:\n"
                "• {CLIENT_NAME} para el nombre del cliente.\n"
                "• {LAST_UPDATE_DATE} para la última fecha de actualización de los datos.\n"
                "También puedes escribir texto plano o usar etiquetas HTML básicas como <b> para negrita y <br> para saltos de línea."
            ),
            style='Custom.TLabel',
            justify='left',
            wraplength=window_width - 20
        )
        instrucciones_label.pack(pady=5, anchor='w')

        # ScrolledText para editar la plantilla
        from tkinter.scrolledtext import ScrolledText
        text_box = ScrolledText(frame, wrap=tk.WORD, width=80, height=15)
        text_box.pack(expand=True, fill='both', padx=5, pady=5)

        # Cargar la plantilla actual desde config
        current_template = self.config_manager.get_email_body_template()
        text_box.insert(tk.END, current_template)

        # Configurar la ventana como modal
        config_window.transient(self)
        config_window.grab_set()
        self.wait_window(config_window)

    def filter_and_generate(self):

        # Ruta de la carpeta donde se generan los reportes
        report_folder = os.path.join(os.getcwd(), "data/reportes")  # Ajusta la ruta si es necesario

        # Verificar si existen reportes previos en la carpeta
        if os.path.exists(report_folder) and os.listdir(report_folder):
            answer = messagebox.askyesno(
                "Advertencia",
                "Existen reportes generados anteriormente en la carpeta 'data/reportes'. "
                "Si continúa, los reportes anteriores serán reemplazados.\n\n¿Desea continuar?"
            )
            if not answer:
                self.update_progress("Operación cancelada por el usuario.", style='Warning.TLabel')
                return

            # Vaciar la carpeta antes de generar los nuevos reportes
            try:
                for file in os.listdir(report_folder):
                    file_path = os.path.join(report_folder, file)
                    if os.path.isfile(file_path) or os.path.islink(file_path):
                        os.unlink(file_path)  # Eliminar archivos y enlaces simbólicos
                    elif os.path.isdir(file_path):
                        shutil.rmtree(file_path)  # Eliminar carpetas dentro de report_folder
                print("Se han eliminado los reportes previos.")
                # 🔹 Limpiar la fecha de última actualización
                self.config_manager.clear_last_update_date()
                print("Se eliminó la última fecha de modificación.")

            except Exception as e:
                messagebox.showerror("Error", f"No se pudo limpiar la carpeta 'data/reportes': {e}")
                return

        # Leer site_url/file_url desde config_manager
        site_url = self.config_manager.get_sharepoint_site_url()
        file_url = self.config_manager.get_sharepoint_file_url()

        try:
            # Crear instancia del FileProcessor
            self.processor = FileProcessor(site_url, file_url)

            # 1) Obtener token de SharePoint (OAuth2)
            self.update_progress("********************************************************", style='Delete.TLabel')
            self.update_progress("Esperando respuesta del servidor...", style='Info.TLabel')
            token = self.get_sharepoint_access_token()
            if not token:
                messagebox.showerror("Error", "No se pudo obtener el token de acceso para SharePoint.")
                return

            # 2) Descargar el archivo usando el token
            self.update_progress("********************************************************", style='Delete.TLabel')
            self.update_progress("Descargando el informe...", style='Info.TLabel')
            self.processor.download_file_with_token(token)

            # 3) Ejecutar la macro
            self.update_progress("********************************************************", style='Delete.TLabel')
            self.update_progress("Actualizando el informe...", style='Info.TLabel')
            self.processor.run_macro()

            # 4) Procesar/generar reportes
            self.update_progress("********************************************************", style='Delete.TLabel')
            self.update_progress("Generando reportes...", style='Info.TLabel')
            self.processor.process_file()

            # Obtener la fecha de la última actualización
            last_update_date = self.processor.get_last_update_date()
            self.config_manager.set_last_update_date(last_update_date)
            if last_update_date:
                # Formatear la fecha para eliminar los segundos
                last_update_date = datetime.strptime(last_update_date, "%Y-%m-%d %H:%M:%S.%f").strftime(
                    "%Y-%m-%d %H:%M")

                # Actualizar el mensaje de progreso con el texto centrado y la fecha formateada
                self.progress_label.config(anchor='center')  # Asegura que el texto esté centrado
                self.update_progress(f"Listo. Última actualización del BI: {last_update_date}",
                                     style='Success.TLabel')
            else:
                # Si no se pudo obtener la fecha, mostrar el mensaje sin la fecha
                self.update_progress("Reportes generados!!", style='Success.TLabel')


            # Mostrar opciones de envío debajo del mensaje de éxito
            messagebox.showinfo("Éxito", "Reportes generados correctamente.")
            answer = self.ask_yes_no_custom(
                "Enviar Reportes",
                "¿Desea enviar los reportes ahora?"
            )
            if answer:
                self.send_mails()
            else:
                # El usuario ha dicho NO
                self.update_progress("********************************************************", style='Delete.TLabel')
                self.update_progress("Los reportes no fueron enviados.", style='Info.TLabel')

        except Exception as e:
            self.update_progress(f"Error: {str(e)}", style='Error.TLabel')
        finally:
            # Asegurarse de cerrar el archivo si la instancia de processor existe
            if self.processor and hasattr(self.processor, 'close') and callable(self.processor.close):
                self.processor.close()

    def generate_reports_simplified(self):
        """
        Genera los reportes eliminando reportes anteriores sin confirmación
        y ejecutando tdo el proceso sin mostrar mensajes al usuario.
        """

        # Ruta de la carpeta donde se generan los reportes
        report_folder = os.path.join(os.getcwd(), "data/reportes")

        # Eliminar reportes previos sin preguntar
        if os.path.exists(report_folder):
            try:
                for file in os.listdir(report_folder):
                    file_path = os.path.join(report_folder, file)
                    if os.path.isfile(file_path) or os.path.islink(file_path):
                        os.unlink(file_path)
                    elif os.path.isdir(file_path):
                        shutil.rmtree(file_path)
                self.config_manager.clear_last_update_date()  # Limpiar la fecha de última actualización
            except Exception as e:
                print(f"Error al limpiar reportes previos: {e}")
                return

        # Obtener configuración de SharePoint
        site_url = self.config_manager.get_sharepoint_site_url()
        file_url = self.config_manager.get_sharepoint_file_url()

        try:
            # Instancia de FileProcessor
            self.processor = FileProcessor(site_url, file_url)

            # Obtener token de SharePoint
            token = self.get_sharepoint_access_token()
            if not token:
                return

            # Descargar el archivo
            self.processor.download_file_with_token(token)

            # Ejecutar macro en Excel
            self.processor.run_macro()

            # Procesar y generar reportes
            self.processor.process_file()

            # Actualizar fecha de última modificación
            last_update_date = self.processor.get_last_update_date()
            if last_update_date:
                self.config_manager.set_last_update_date(last_update_date)

        except Exception as e:
            print(f"Error en la generación de reportes: {e}")
        finally:
            if self.processor and hasattr(self.processor, 'close'):
                self.processor.close()

    def filter_generate_and_send(self):
        """
        Genera los reportes y los envía automáticamente sin preguntar.
        """
        self.generate_reports_simplified()
        self.send_mails_filtered()

    def open_client_selection_window(self):
        """
        Ventana de configuración para seleccionar clientes que recibirán correos.
        Guarda el estado en un archivo JSON para persistencia.
        Incluye opción de 'Seleccionar todos' y un scrollbar para muchos clientes.
        """
        import json  # Asegúrate de tenerlo importado al principio del archivo
        config_file = "clientes_config.json"

        # Cargar configuración previa si existe
        if os.path.exists(config_file):
            with open(config_file, "r", encoding="utf-8") as f:
                client_config = json.load(f)
        else:
            client_config = {}

        # Obtener la lista de clientes desde la base de datos
        clientes = self.db.get_all_clients()  # Retorna tuplas (id, nombre, correo, correo_cc, correo_bcc)

        if not clientes:
            messagebox.showwarning("Advertencia", "No hay clientes en la base de datos.")
            return

        # Crear ventana
        window = tk.Toplevel(self)
        window.title("Seleccionar Clientes para Envío")
        # Fijar tamaño y evitar que el usuario la redimensione
        window.geometry("500x500")
        window.resizable(False, False)
        # Centramos la ventana (opcional, según tu implementación)
        # self.center_window(window, 400, 500)  # si deseas usar tu función center_window

        # Frame principal con 3 filas:
        # [0] Encabezado + "Seleccionar todos"
        # [1] Zona con scroll para checkboxes
        # [2] Botón de Guardar
        main_frame = ttk.Frame(window, padding=10)
        main_frame.pack(expand=True, fill="both")

        main_frame.rowconfigure(0, weight=0)  # Encabezado, no se expande
        main_frame.rowconfigure(1, weight=1)  # Área con scroll, ocupa espacio
        main_frame.rowconfigure(2, weight=0)  # Botón "Guardar", fijo
        main_frame.columnconfigure(0, weight=1)

        # ─────────────────────────────────────────────
        # FILA 0: Encabezado + Checkbutton "Seleccionar Todos"
        # ─────────────────────────────────────────────
        top_frame = ttk.Frame(main_frame)
        top_frame.grid(row=0, column=0, sticky="ew", pady=(0, 5))

        label_instrucciones = ttk.Label(
            top_frame,
            text="Selecciona los clientes que recibirán correos:",
            style='Custom.TLabel'
        )
        label_instrucciones.pack(side="left", anchor="w")

        # Variable y función para "Seleccionar todos"
        select_all_var = tk.BooleanVar(value=False)

        def on_select_all_changed():
            new_val = select_all_var.get()
            # Asignar todos los checkboxes a new_val
            for var in check_vars.values():
                var.set(new_val)

        check_all_btn = ttk.Checkbutton(
            top_frame,
            text="Seleccionar todos",
            variable=select_all_var,
            command=on_select_all_changed
        )
        check_all_btn.pack(side="right", anchor="e")

        # ─────────────────────────────────────────────
        # FILA 1: Área con scrollbar que contendrá las checkboxes
        # ─────────────────────────────────────────────
        scroll_frame = ttk.Frame(main_frame)
        scroll_frame.grid(row=1, column=0, sticky="nsew")

        # Canvas + Scrollbar
        canvas = tk.Canvas(scroll_frame)
        canvas.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(scroll_frame, orient="vertical", command=canvas.yview)
        scrollbar.pack(side="right", fill="y")

        canvas.configure(yscrollcommand=scrollbar.set)

        # Frame interno donde colocar las checkboxes
        checklist_frame = ttk.Frame(canvas)
        # Crear ventana dentro del canvas
        canvas_window = canvas.create_window((0, 0), window=checklist_frame, anchor="nw")

        # Para actualizar la región de scroll cuando cambie el contenido
        def on_checklist_frame_configure(event):
            canvas.config(scrollregion=canvas.bbox("all"))

        checklist_frame.bind("<Configure>", on_checklist_frame_configure)

        # También, cuando se cambie el tamaño del canvas, reubicar la ventana para no cortar contenido
        def on_canvas_resize(event):
            canvas.itemconfig(canvas_window, width=event.width)

        canvas.bind("<Configure>", on_canvas_resize)

        # Diccionario para almacenar estados de checkboxes
        check_vars = {}

        def on_individual_check():
            """
            Evento llamado cuando se modifica un checkbox individual.
            Si uno se desactiva, desactivamos 'Seleccionar todos'.
            Si todos están activos, activamos 'Seleccionar todos'.
            """
            all_selected = all(var.get() for var in check_vars.values())
            select_all_var.set(all_selected)

        # Crear checkboxes dentro de checklist_frame
        for client_id, client_name, _, _, _ in clientes:
            var = tk.BooleanVar(value=client_config.get(client_name, False))
            check_vars[client_name] = var
            c = ttk.Checkbutton(
                checklist_frame,
                text=client_name,
                variable=var,
                command=on_individual_check  # Para monitorear cambios
            )
            c.pack(anchor="w", padx=5, pady=2)

        # ─────────────────────────────────────────────
        # FILA 2: Botón "Guardar"
        # ─────────────────────────────────────────────
        def save_client_selection():
            """
            Guarda el estado de los checkboxes en un archivo JSON.
            """
            new_config = {name: var.get() for name, var in check_vars.items()}
            with open(config_file, "w", encoding="utf-8") as f:
                json.dump(new_config, f, indent=4)
            messagebox.showinfo("Guardado", "Configuración guardada correctamente.")
            window.destroy()

        ttk.Button(
            main_frame,
            text="Guardar",
            command=save_client_selection,
            style='GreenSuccess.TButton'
        ).grid(row=2, column=0, pady=5, sticky="ew")

        # Hacer la ventana modal
        window.transient(self)
        window.grab_set()
        self.wait_window(window)

    def send_mails_filtered(self):
        """
        Envía correos solo a los clientes seleccionados en la configuración.
        """
        config_file = "clientes_config.json"

        if os.path.exists(config_file):
            with open(config_file, "r", encoding="utf-8") as f:
                client_config = json.load(f)
        else:
            messagebox.showerror("Error",
                                 "No se encontró la configuración de clientes. Configura primero la selección.")
            return

        # 📌 Obtener el alias FROM para enviar
        from_email = self.config_manager.get_default_sender()
        if not from_email:
            from_email = self.select_sender_alias()

        if not from_email:
            self.update_progress("Envío cancelado por el usuario.", style='Warning.TLabel')
            return

        # Obtener lista de clientes desde la BD
        clientes_con_correo = [
            (client_id, nombre, correo, correo_cc, correo_bcc)
            for (client_id, nombre, correo, correo_cc, correo_bcc) in self.db.get_all_clients()
            if correo and client_config.get(nombre, False)  # Solo incluir si está seleccionado
        ]

        if not clientes_con_correo:
            messagebox.showwarning("Advertencia", "No hay clientes seleccionados para recibir correos.")
            return

        errores_envio = []
        enviados_exitosamente = []
        template = self.config_manager.get_email_body_template()

        # Obtener fecha actual para reemplazar en la plantilla
        current_date = datetime.now().strftime("%d/%m/%Y")

        for client_id, cliente, correo, correo_cc, correo_bcc in clientes_con_correo:
            archivo_path = self.processor.get_client_file(cliente)
            if archivo_path:
                cuerpo_html = template.replace("{CLIENT_NAME}", cliente)
                cuerpo_html = cuerpo_html.replace("{LAST_UPDATE_DATE}", current_date)

                asunto = f"Reporte para {cliente}"

                # 🔥 Aquí PASAMOS el from_email
                success, msg = self.email_sender.send_email(
                    to_email=correo,
                    subject=asunto,
                    html_content=cuerpo_html,
                    attachment_path=archivo_path,
                    cc_emails=correo_cc,
                    bcc_emails=correo_bcc,
                    from_email=from_email  # 👈 alias personalizado
                )

                if success:
                    enviados_exitosamente.append(cliente)
                else:
                    errores_envio.append(f"Error al enviar a '{cliente}': {msg}")

        if enviados_exitosamente:
            messagebox.showinfo("Éxito", f"Reportes enviados a: {', '.join(enviados_exitosamente)}")

        if errores_envio:
            messagebox.showerror("Errores de Envío", "\n".join(errores_envio))

    def ask_yes_no_custom(self, title: str, message: str) -> bool:
        """
        Muestra un diálogo personalizado con estilo ttk, en lugar de usar messagebox.askyesno.
        Devuelve True si el usuario pulsa 'Sí', False si pulsa 'No' o cierra la ventana.
        """

        dialog = tk.Toplevel(self)
        dialog.title(title)

        # Opcional: centrar la ventana
        width, height = 400, 200
        self.center_window(dialog, width, height)

        # Evitar interacción con la ventana principal mientras está abierto (modal)
        dialog.transient(self)
        dialog.grab_set()

        # Frame principal
        frame = ttk.Frame(dialog, padding=20)
        frame.pack(expand=True, fill='both')

        # Etiqueta con el mensaje
        label_message = ttk.Label(frame, text=message, style='Info.TLabel')
        label_message.pack(pady=(0, 20), anchor='center')

        # Variable para almacenar la respuesta del usuario
        user_answer = tk.BooleanVar(value=False)

        # Funciones internas para los botones
        def on_yes():
            user_answer.set(True)
            dialog.destroy()

        def on_no():
            user_answer.set(False)
            dialog.destroy()

        # Botón "Sí"
        button_yes = ttk.Button(frame, text="Sí", command=on_yes, style='BlueZodiac.TButton')
        button_yes.pack(side='left', expand=True, padx=10)

        # Botón "No"
        button_no = ttk.Button(frame, text="No", command=on_no, style='RedColor.TButton')
        button_no.pack(side='right', expand=True, padx=10)

        # Esperar hasta que la ventana se destruya para continuar
        self.wait_window(dialog)

        return user_answer.get()

    def send_mails(self):
        # Mostrar el GIF de carga y actualizar el mensaje de progreso
        self.update_progress("********************************************************", style='Delete.TLabel')
        self.update_progress("Verificando autenticación...", style='Info.TLabel')

        # Obtener la información del usuario autenticado
        user_info = self.get_current_authenticated_user()

        if not user_info:
            self.update_progress("********************************************************", style='Delete.TLabel')
            self.update_progress("No estás autenticado. Iniciando sesión...", style='Warning.TLabel')
            # Abrir la ventana de configuración para iniciar sesión
            self.open_config_window()
            # Después de cerrar la ventana de configuración, verificar nuevamente la autenticación
            user_info = self.get_current_authenticated_user()
            if not user_info:
                self.update_progress("Autenticación fallida. No se pueden enviar correos.", style='Error.TLabel')
                return
            else:
                self.update_progress("Autenticación exitosa. Preparando para enviar correos...", style='Success.TLabel')

        # Continuar con el envío de correos si está autenticado
        try:
            email = user_info['email']
            if not email:
                self.update_progress("No se pudo obtener el correo electrónico del usuario.", style='Error.TLabel')
                return

            # Seleccionar la lista para enviar correos
            listas = self.db.get_all_lists()
            if not listas:
                messagebox.showwarning("Advertencia", "No hay listas disponibles. Por favor, crea una lista primero.")
                self.update_progress("No hay listas disponibles para enviar correos.", style='Warning.TLabel')
                return

            last_update_date = self.processor.get_last_update_date()
            # Formatear la fecha de última actualización para mostrarla en la ventana de selección de lista
            last_update_date = datetime.strptime(last_update_date, "%Y-%m-%d %H:%M:%S.%f").strftime("%Y-%m-%d %H:%M")


            lista_nombres = [nombre for (id, nombre) in listas]
            list_var = tk.StringVar(value=lista_nombres[0] if lista_nombres else '')

            select_list_window = tk.Toplevel(self)
            select_list_window.title("Seleccionar Lista")

            # Ajustar el tamaño inicial de la ventana según el escalado
            initial_width = int(300 * self.scaling_factor)
            initial_height = int(150 * self.scaling_factor)

            self.center_window(select_list_window, initial_width, initial_height)

            etiqueta_lista = ttk.Label(select_list_window, text="Selecciona la lista para enviar correos:",
                                       style='Custom.TLabel')
            etiqueta_lista.pack(pady=5)

            list_menu = ttk.OptionMenu(select_list_window, list_var, *lista_nombres)
            list_menu.pack(pady=5)

            # Agregar la etiqueta de última actualización
            etiqueta_fecha = ttk.Label(
                select_list_window,
                text=f"Última actualización: {last_update_date}",
                style='Info.TLabel'  # Puedes personalizar el estilo si es necesario
            )
            etiqueta_fecha.pack(pady=5)

            def confirmar_lista():
                seleccion = list_var.get()
                list_id = None
                for id, nombre in listas:
                    if nombre == seleccion:
                        list_id = id
                        break
                if list_id is not None:
                    select_list_window.destroy()
                    self.send_mails_to_list(list_id)
                else:
                    messagebox.showerror("Error", "No se pudo obtener la lista seleccionada.")

            boton_confirmar = ttk.Button(
                select_list_window,
                text="Confirmar",
                command=confirmar_lista,
                style='GreenSuccess.TButton'
            )
            boton_confirmar.pack(pady=10)

            select_list_window.transient(self)
            select_list_window.grab_set()
            self.wait_window(select_list_window)

        except Exception as e:
            self.update_progress(f"Error: {str(e)}", style='Error.TLabel')
        finally:
            pass

    def send_mails_to_list(self, list_id):
        self.update_progress("********************************************************", style='Delete.TLabel')
        self.update_progress("Enviando los reportes...", style='Info.TLabel')
        self.update_idletasks()

        try:
            if not self.processor:
                self.update_progress("", style='Info.TLabel')
                self.update_progress("Primero debes generar los reportes.", style='Warning.TLabel')
                return

            # 📌 Obtener el alias FROM para enviar
            from_email = self.config_manager.get_default_sender()
            if not from_email:
                from_email = self.select_sender_alias()

            if not from_email:
                self.update_progress("Envío cancelado por el usuario.", style='Warning.TLabel')
                return

            # Extrae todos los campos: ID, Nombre, Correo, Correo CC y Correo BCC
            clientes_con_correo = [
                (client_id, nombre, correo, correo_cc, correo_bcc)
                for (client_id, nombre, correo, correo_cc, correo_bcc) in self.db.get_clients_by_list(list_id)
                if correo and correo.lower().strip() != '(correo sin definir)'
            ]

            if not clientes_con_correo:
                self.update_progress("", style='Info.TLabel')
                self.update_progress("No hay clientes con correos asociados para enviar.", style='Warning.TLabel')
                return

            # Obtener la plantilla del correo
            template = self.config_manager.get_email_body_template()

            # Obtener la fecha actual
            current_date = datetime.now().strftime("%d/%m/%Y")

            errores_envio = []
            for client_id, cliente, correo, correo_cc, correo_bcc in clientes_con_correo:
                archivo_path = self.processor.get_client_file(cliente)
                if archivo_path:
                    cuerpo_html = template.replace("{CLIENT_NAME}", cliente)
                    cuerpo_html = cuerpo_html.replace("{LAST_UPDATE_DATE}", current_date)

                    asunto = f"Reporte para {cliente}"

                    # 🔥 Aquí PASAMOS el from_email
                    success, msg = self.email_sender.send_email(
                        to_email=correo,
                        subject=asunto,
                        html_content=cuerpo_html,
                        attachment_path=archivo_path,
                        cc_emails=correo_cc,
                        bcc_emails=correo_bcc,
                        from_email=from_email  # 👈 alias seleccionado
                    )

                    if not success:
                        if "Se requiere reautenticación" in msg:
                            self.update_progress("", style='Info.TLabel')
                            self.update_progress("Reautenticación requerida. Por favor, inicia sesión nuevamente.",
                                                 style='Warning.TLabel')
                            self.open_config_window()
                            errores_envio.append(f"Reautenticación requerida para {cliente} ({correo}).")
                        else:
                            errores_envio.append(f"Error al enviar a {cliente} ({correo}): {msg}")
                else:
                    errores_envio.append(f"No se encontró el archivo para el cliente {cliente}.")

            if errores_envio:
                self.update_progress("Errores durante el envío de correos.", style='Error.TLabel')
                errores_detallados = "\n".join(errores_envio)
                messagebox.showerror("Errores de Envío", errores_detallados)
            else:
                self.update_progress("", style='Info.TLabel')
                self.update_progress("Reportes enviados exitosamente!", style='Success.TLabel')

        except Exception as e:
            self.update_progress(f"Error: {str(e)}", style='Error.TLabel')
        finally:
            pass

    def open_client_manager(self):
        open_client_manager(self)

    def on_tab_change(self, event):
        on_tab_change(self, event)

    def on_list_select_por_lista(self, event):
        on_list_select_por_lista(self, event)

    def load_all_clients(self):
        load_all_clients(self)

    def add_client_to_existing_list(self):
        add_client_to_existing_list(self)

    def on_list_select(self, event):
        on_list_select(self, event)

    def load_clients_in_selected_list(self):
        load_clients_in_selected_list(self)

    def on_client_select(self, event):
        on_client_select(self, event)

    def on_all_client_select(self, event):
        on_all_client_select(self, event)

    def add_client(self):
        add_client(self)

    def edit_client(self):
        edit_client(self)

    def delete_client(self):
        delete_client(self)


    def add_alias(self):
        add_alias(self)

    def delete_alias(self):
        delete_alias(self)

    def on_closing(self):
        on_closing(self)

    def create_new_list(self):
        create_new_list(self)

    def delete_selected_list(self):
        delete_selected_list(self)

    def update_list_combobox_ui(self):
        update_list_combobox_ui(self)

    def add_client_to_selected_list(self):
        add_client_to_selected_list(self)

    def remove_client_from_selected_list(self):
        remove_client_from_selected_list(self)

    def export_clients(self):
        export_clients(self)


    def import_clients(self):
        import_clients(self)
    #VENTANAS

    def center_window(self, window, width, height):
        """
        Centra una ventana en la pantalla.

        :param window: La ventana a centrar (tk.Toplevel o la ventana principal).
        :param width: Ancho de la ventana.
        :param height: Alto de la ventana.
        """
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x = int((screen_width / 2) - (width / 2))
        y = int((screen_height / 2) - (height / 2))
        window.geometry(f"{width}x{height}+{x}+{y}")

    def show_context_menu(self, event):
        show_context_menu(self, event)

    def context_menu_add_to_list(self):
        context_menu_add_to_list(self)

    def show_add_to_list_dialog(self, tree, selected_items, available_lists):
        """
        Muestra un diálogo para seleccionar una lista y asignar todos los clientes seleccionados a esa lista.
        """

        def guardar_asignacion():
            selected_list = combobox_lists.get()
            if not selected_list:
                messagebox.showwarning("Advertencia", "Por favor, selecciona una lista.")
                return
            list_id = self.list_ids.get(selected_list)
            if not list_id:
                messagebox.showerror("Error", "La lista seleccionada no es válida.")
                return

            clientes_ya_en_lista = []
            clientes_agregados = []

            try:
                # Asignar cada cliente seleccionado a la lista
                for item in selected_items:
                    client_id = int(tree.item(item, 'values')[0])
                    # Verificar si el cliente ya está en la lista
                    if not self.db.is_client_in_list(client_id, list_id):
                        self.db.add_client_to_list(client_id, list_id)
                        clientes_agregados.append(tree.item(item, 'values')[1])  # Nombre del cliente
                    else:
                        clientes_ya_en_lista.append(tree.item(item, 'values')[1])  # Nombre del cliente

                mensaje = ""
                if clientes_agregados:
                    mensaje += f"Clientes agregados a la lista '{selected_list}':\n" + "\n".join(
                        clientes_agregados) + "\n\n"
                if clientes_ya_en_lista:
                    mensaje += "Los siguientes clientes ya estaban en la lista y fueron ignorados:\n" + "\n".join(
                        clientes_ya_en_lista)
                messagebox.showinfo("Resultado", mensaje)
                ventana_asignar.destroy()

                # Recargar las tablas para reflejar el cambio
                self.load_all_clients()
                self.load_clients_in_selected_list()
            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error al agregar los clientes a la lista: {str(e)}")

        # Crear la ventana de asignación a lista
        ventana_asignar = ttk.Toplevel(self.client_manager)
        ventana_asignar.title("Añadir Clientes a una Lista")

        # Ajustar el tamaño inicial de la ventana según el escalado
        initial_width = int(400 * self.scaling_factor)
        initial_height = int(200 * self.scaling_factor)
        self.center_window(ventana_asignar, initial_width, initial_height)

        # Configuración del grid
        ventana_asignar.columnconfigure(0, weight=1)
        ventana_asignar.rowconfigure(0, weight=1)

        # Frame principal
        frame_principal = ttk.Frame(ventana_asignar, padding=int(10 * self.scaling_factor))
        frame_principal.grid(row=0, column=0, sticky='nsew')
        frame_principal.columnconfigure(1, weight=1)

        # Etiqueta y Combobox para seleccionar lista
        etiqueta_lista = ttk.Label(
            frame_principal,
            text="Seleccionar Lista:",
            style='Custom.TLabel',
        )
        etiqueta_lista.grid(row=0, column=0, pady=int(5 * self.scaling_factor), sticky='w')

        combobox_lists = ttk.Combobox(
            frame_principal,
            values=list(available_lists.keys()),
            state='readonly',
            width=int(30 * self.scaling_factor),
        )
        combobox_lists.grid(row=0, column=1, pady=int(5 * self.scaling_factor), padx=int(5 * self.scaling_factor),
                            sticky='ew')
        combobox_lists.set(list(available_lists.keys())[0] if available_lists else '')

        # Botón para guardar la asignación
        boton_guardar = ttk.Button(
            frame_principal,
            text="Agregar",
            command=guardar_asignacion,
            style='GreenSuccess.TButton'
        )
        boton_guardar.grid(row=1, column=0, columnspan=2, pady=int(20 * self.scaling_factor),
                           ipadx=int(10 * self.scaling_factor), ipady=int(5 * self.scaling_factor), sticky='ew')

        # Configurar la ventana como modal
        ventana_asignar.transient(self.client_manager)
        ventana_asignar.grab_set()
        self.client_manager.wait_window(ventana_asignar)

    def context_menu_add_to_new_list(self):
        """
        Maneja la acción de agregar los clientes seleccionados a una nueva lista existente.
        """
        # Determinar qué Treeview está activo
        focused_widget = self.client_manager.focus_get()
        if focused_widget == self.tree_clients_por_lista:
            tree = self.tree_clients_por_lista
        elif focused_widget == self.tree_all_clients:
            tree = self.tree_all_clients
        else:
            messagebox.showwarning("Advertencia", "No se pudo determinar la pestaña activa.")
            return

        selected_items = tree.selection()
        if not selected_items:
            messagebox.showwarning("Advertencia", "Por favor, selecciona al menos un cliente.")
            return

        # Abrir el diálogo para crear una nueva lista
        self.show_create_new_list_dialog(tree, selected_items)

    def show_create_new_list_dialog(self, tree, selected_items):
        """
        Muestra un diálogo para crear una nueva lista y asignar los clientes seleccionados a ella.
        """

        def guardar_nueva_lista():
            nombre_lista = entrada_nombre.get().strip()
            if not nombre_lista:
                messagebox.showwarning("Advertencia", "El nombre de la lista no puede estar vacío.")
                return
            try:
                # Crear la nueva lista en la base de datos
                self.db.add_list(nombre_lista)
                # Obtener el ID de la nueva lista
                list_id = self.db.get_list_id_by_name(nombre_lista)
                if not list_id:
                    messagebox.showerror("Error", "No se pudo obtener el ID de la lista recién creada.")
                    return
                # Asignar todos los clientes seleccionados a la nueva lista
                for item in selected_items:
                    client_id = int(tree.item(item, 'values')[0])
                    self.db.add_client_to_list(client_id, list_id)
                # Actualizar la interfaz
                self.update_list_combobox_ui()
                self.load_all_clients()
                self.load_clients_in_selected_list()
                messagebox.showinfo("Éxito", f"Lista '{nombre_lista}' creada y clientes agregados exitosamente.")
                ventana_crear_lista.destroy()
            except sqlite3.IntegrityError:
                messagebox.showerror("Error", f"La lista '{nombre_lista}' ya existe.")
            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error al crear la lista: {str(e)}")

        # Crear la ventana de diálogo
        ventana_crear_lista = ttk.Toplevel(self.client_manager)
        ventana_crear_lista.title("Crear Nueva Lista y Agregar Clientes")

        # Ajustar el tamaño inicial de la ventana según el escalado
        initial_width = int(400 * self.scaling_factor)
        initial_height = int(150 * self.scaling_factor)
        self.center_window(ventana_crear_lista, initial_width, initial_height)

        # Configuración del grid
        ventana_crear_lista.columnconfigure(0, weight=1)
        ventana_crear_lista.rowconfigure(0, weight=1)
        ventana_crear_lista.rowconfigure(1, weight=1)
        ventana_crear_lista.rowconfigure(2, weight=1)

        # Frame principal
        frame_principal = ttk.Frame(ventana_crear_lista, padding=int(10 * self.scaling_factor))
        frame_principal.grid(row=0, column=0, sticky='nsew')
        frame_principal.columnconfigure(1, weight=1)

        # Etiqueta y entrada para el nombre de la lista
        etiqueta_nombre = ttk.Label(
            frame_principal,
            text="Nombre de la Nueva Lista:",
            style='Custom.TLabel',
        )
        etiqueta_nombre.grid(row=0, column=0, pady=int(5 * self.scaling_factor), sticky='w')

        entrada_nombre = ttk.Entry(
            frame_principal,
            width=int(30 * self.scaling_factor),
        )
        entrada_nombre.grid(row=0, column=1, pady=int(5 * self.scaling_factor), padx=int(5 * self.scaling_factor),
                            sticky='ew')

        # Botón para guardar la nueva lista
        boton_guardar = ttk.Button(
            frame_principal,
            text="Crear y Agregar",
            command=guardar_nueva_lista,
            style='GreenSuccess.TButton'
        )
        boton_guardar.grid(row=1, column=0, columnspan=2, pady=int(20 * self.scaling_factor),
                           ipadx=int(10 * self.scaling_factor), ipady=int(5 * self.scaling_factor), sticky='ew')

        # Configurar la ventana como modal
        ventana_crear_lista.transient(self.client_manager)
        ventana_crear_lista.grab_set()
        self.client_manager.wait_window(ventana_crear_lista)

    def delete_client_from_context_menu(self):
        delete_client_from_context_menu(self)

    def edit_alias_inline(self, event):
        edit_alias_inline(self, event)

    def open_sharepoint_file_settings_window(self):
        """
        Abre una ventana para configurar la URL del sitio de SharePoint y la ruta del archivo.
        """
        config_window = tk.Toplevel(self)
        config_window.title("Configurar URL de SharePoint")

        # Ajustar tamaño y centrar la ventana
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        scaling_factor = min(screen_width / 1920, screen_height / 1080)
        window_width = int(500 * scaling_factor)
        window_height = int(250 * scaling_factor)
        self.center_window(config_window, window_width, window_height)

        # Crear un frame principal
        frame = ttk.Frame(config_window, padding=20)
        frame.pack(expand=True, fill="both")

        # Etiqueta + Entry para Site URL
        ttk.Label(frame, text="Site URL de SharePoint:", style='Custom.TLabel').grid(row=0, column=0, sticky='w')
        site_url_var = tk.StringVar(value=self.config_manager.get_sharepoint_site_url())
        site_url_entry = ttk.Entry(frame, textvariable=site_url_var, width=80)
        site_url_entry.grid(row=1, column=0, padx=5, pady=5, sticky='ew')

        # Etiqueta + Entry para File URL
        ttk.Label(frame, text="File URL en SharePoint:", style='Custom.TLabel').grid(row=2, column=0, sticky='w')
        file_url_var = tk.StringVar(value=self.config_manager.get_sharepoint_file_url())
        file_url_entry = ttk.Entry(frame, textvariable=file_url_var, width=80)
        file_url_entry.grid(row=3, column=0, padx=5, pady=5, sticky='ew')

        def save_urls():
            site_url = site_url_var.get().strip()
            file_url = file_url_var.get().strip()
            if not site_url or not file_url:
                messagebox.showwarning("Advertencia", "Por favor, completa ambos campos.")
                return

            # Guardar en la configuración
            self.config_manager.set_sharepoint_site_url(site_url)
            self.config_manager.set_sharepoint_file_url(file_url)

            messagebox.showinfo("Éxito", "URLs de SharePoint guardadas exitosamente.")
            config_window.destroy()

        # Botón para guardar
        save_button = ttk.Button(frame, text="Guardar", command=save_urls, style='GreenSuccess.TButton')
        save_button.grid(row=4, column=0, pady=(10, 0), sticky='ew')

        # Hacer la ventana modal
        config_window.transient(self)
        config_window.grab_set()
        self.wait_window(config_window)

    def undo_action(self, event=None):
        undo_action(self, event)

    def send_reports_to_selected_clients(self):
        send_reports_to_selected_clients(self)

    import os

    def process_sending_reports(self, selected_items, tree):
        self.after(0, lambda: self.update_progress("Enviando reportes...", style='Info.TLabel'))

        errores_envio = []
        enviados_exitosamente = []
        template = self.config_manager.get_email_body_template()

        # 📌 Obtener el alias FROM para enviar
        from_email = self.config_manager.get_default_sender()
        if not from_email:
            from_email = self.select_sender_alias()

        if not from_email:
            self.after(0, lambda: self.update_progress("Envío cancelado por el usuario.", style='Warning.TLabel'))
            return

        # Obtener la fecha del día actual en formato dd/mm/yyyy
        current_date = datetime.now().strftime("%d/%m/%Y")

        for item in selected_items:
            client_values = tree.item(item, 'values')
            client_id = int(client_values[0])
            cliente_nombre = client_values[1]
            cliente_correo = client_values[2]
            cliente_correo_cc = client_values[3] if len(client_values) > 3 else ''
            cliente_correo_bcc = client_values[4] if len(client_values) > 4 else ''

            if not cliente_correo or cliente_correo.lower().strip() == '(correo sin definir)':
                errores_envio.append(f"El cliente '{cliente_nombre}' no tiene un correo válido.")
                continue

            reporte_path = self.processor.get_client_file(cliente_nombre)
            if not reporte_path or not os.path.exists(reporte_path):
                errores_envio.append(
                    f"No se encontró el reporte para el cliente '{cliente_nombre}'. Debe generar primero los reportes.")
                continue

            # Personalizar cuerpo del correo
            cuerpo_html = template.replace("{CLIENT_NAME}", cliente_nombre)
            cuerpo_html = cuerpo_html.replace("{LAST_UPDATE_DATE}", current_date)

            asunto = f"Reporte para {cliente_nombre}"

            # 🔥 Aquí PASAMOS el from_email
            success, msg = self.email_sender.send_email(
                to_email=cliente_correo,
                subject=asunto,
                html_content=cuerpo_html,
                attachment_path=reporte_path,
                cc_emails=cliente_correo_cc,
                bcc_emails=cliente_correo_bcc,
                from_email=from_email  # 👈 alias personalizado
            )

            if success:
                enviados_exitosamente.append(cliente_nombre)
            else:
                errores_envio.append(f"Error al enviar a '{cliente_nombre}': {msg}")

        if enviados_exitosamente:
            self.after(0, lambda: self.update_progress(
                f"Reportes enviados a: {', '.join(enviados_exitosamente)}", style='Success.TLabel')
                       )

        if errores_envio:
            self.after(0, lambda: self.update_progress("Errores durante el envío de reportes.", style='Error.TLabel'))
            errores_detallados = "\n".join(errores_envio)
            self.after(0, lambda: messagebox.showerror("Errores de Envío", errores_detallados))
        else:
            self.after(0, lambda: messagebox.showinfo("Éxito", "Todos los reportes fueron enviados exitosamente."))

    def open_reports_in_excel(self):
        open_reports_in_excel(self)

    def select_all_with_reports_in_all_clients(self):
        select_all_with_reports_in_all_clients(self)

    def filter_all_clients(self):
        filter_all_clients(self)

    def get_user_aliases(self) -> list[str]:
        """
        Consulta Microsoft Graph para obtener todos los alias (proxyAddresses) del usuario autenticado.
        """
        access_token = self.get_email_access_token()
        if not access_token:
            return []

        try:
            headers = {'Authorization': f'Bearer {access_token}'}
            # Cambiamos la URL para pedir solo proxyAddresses
            response = requests.get('https://graph.microsoft.com/v1.0/me?$select=proxyAddresses', headers=headers)
            response.raise_for_status()
            data = response.json()
            proxy_addresses = data.get('proxyAddresses', [])

            # Solo extraer las direcciones válidas de tipo SMTP
            aliases = []
            for address in proxy_addresses:
                if address.lower().startswith("smtp:"):  # SMTP (minúscula) es el alias secundario
                    aliases.append(address[5:])  # Remover el prefijo "smtp:"

            return aliases
        except Exception as e:
            print(f"Error al obtener aliases: {e}")
            return []

    def select_sender_alias(self) -> Optional[str]:
        from tkinter import StringVar, BooleanVar

        def abrir_ventana_alias():
            alias_window = tk.Toplevel(self)
            alias_window.title("Seleccionar Alias de Envío")
            self.center_window(alias_window, 450, 350)

            # Cargar alias
            dynamic_aliases = self.get_user_aliases()
            manual_aliases = self.config_manager.get_manual_aliases()
            all_aliases = list(set(dynamic_aliases + manual_aliases))

            if not all_aliases:
                user_info = self.get_current_authenticated_user()
                if user_info:
                    all_aliases = [user_info['email']]
                else:
                    return None

            selected_alias = StringVar(value=all_aliases[0])
            remember_default = BooleanVar()

            ttk.Label(alias_window, text="Seleccione el alias para enviar los correos:", style='Custom.TLabel').pack(
                pady=10)

            radio_frame = ttk.Frame(alias_window)
            radio_frame.pack(pady=5)

            def eliminar_alias_manual(alias_a_eliminar):
                confirm = messagebox.askyesno("Confirmar", f"¿Eliminar alias '{alias_a_eliminar}'?")
                if confirm:
                    manual_aliases = self.config_manager.get_manual_aliases()
                    if alias_a_eliminar in manual_aliases:
                        manual_aliases.remove(alias_a_eliminar)
                        self.config_manager.config["manual_aliases"] = manual_aliases
                        self.config_manager.save_config()
                        alias_window.destroy()
                        self.after(100, abrir_ventana_alias)  # Recargar ventana

            for alias in all_aliases:
                frame_alias = ttk.Frame(radio_frame)
                frame_alias.pack(anchor='w', pady=2, padx=20, fill='x')

                ttk.Radiobutton(frame_alias, text=alias, variable=selected_alias, value=alias).pack(side='left',
                                                                                                    anchor='w')

                # Solo mostrar botón de borrar para alias manuales
                if alias in manual_aliases:
                    boton_eliminar = ttk.Button(
                        frame_alias,
                        text="X",
                        width=2,
                        style='RedColor.TButton',
                        command=lambda a=alias: eliminar_alias_manual(a)
                    )
                    boton_eliminar.pack(side='left', padx=5)

            # Entrada para agregar nuevo alias
            add_frame = ttk.Frame(alias_window)
            add_frame.pack(pady=10)

            new_alias_entry = ttk.Entry(add_frame, width=30)
            new_alias_entry.grid(row=0, column=0, padx=(5, 5))

            def agregar_alias_manual():
                nuevo_alias = new_alias_entry.get().strip()
                if nuevo_alias and nuevo_alias not in all_aliases:
                    self.config_manager.add_manual_alias(nuevo_alias)
                    alias_window.destroy()
                    self.after(100, abrir_ventana_alias)  # Recargar ventana
                else:
                    messagebox.showwarning("Advertencia", "El alias ya existe o es inválido.")

            ttk.Button(add_frame, text="Agregar Alias", command=agregar_alias_manual, style='BlueZodiac.TButton').grid(
                row=0, column=1)

            ttk.Checkbutton(alias_window, text="Usar como predeterminado", variable=remember_default).pack(pady=10)

            frame_btns = ttk.Frame(alias_window)
            frame_btns.pack(pady=10)

            result = {"alias": None}

            def confirmar():
                if selected_alias.get():
                    result["alias"] = selected_alias.get()
                    if remember_default.get():
                        self.config_manager.set_default_sender(result["alias"])
                    alias_window.destroy()

            def cancelar():
                alias_window.destroy()

            ttk.Button(frame_btns, text="Confirmar", command=confirmar, style='GreenSuccess.TButton').pack(side='left',
                                                                                                           padx=10)
            ttk.Button(frame_btns, text="Cancelar", command=cancelar, style='RedColor.TButton').pack(side='right',
                                                                                                     padx=10)

            alias_window.transient(self)
            alias_window.grab_set()
            self.wait_window(alias_window)

            return result["alias"]

        return abrir_ventana_alias()

    def is_sharepoint_session_active(self):
        """
        Verifica si hay una sesión activa en SharePoint revisando si hay tokens almacenados en el caché.
        """
        token_cache = self.sharepoint_msal_app.token_cache.serialize()

        if token_cache:
            token_data = json.loads(token_cache)
            return any(entry.get("credential_type") == "RefreshToken" for entry in token_data.get("AccessToken", []))

        return False  # Si no hay tokens almacenados, no hay sesión activa

