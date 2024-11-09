# src/gui/app.py
import http.server
import json
import os
import socketserver
import sqlite3
import sys
from urllib.parse import quote
import threading
from urllib.parse import urlparse, parse_qs
import ttkbootstrap as ttk
import tkinter as tk
from tkinter import messagebox, scrolledtext, filedialog
import webbrowser
import requests
import logging
from PIL import Image, ImageTk, ImageSequence
from dotenv import load_dotenv
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from email_service.email_sender import EmailSender
from msal import PublicClientApplication
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
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

from src.manager.client_manager import context_menu_add_to_list, show_context_menu

# Credenciales de Microsoft OAuth2
MICROSOFT_CLIENT_ID = "077e19a4-ab53-4e55-8286-ec02dc3cf36c"
MICROSOFT_CLIENT_SECRET = "pDK8Q~V63IOhtEe6KKrOzz8waIApiaXYrhq1MaFL"
MICROSOFT_TENANT_ID = "common"
REDIRECT_URI = "http://localhost:62190/"

from gui.styles import init_styles
# Importar funciones y clases desde los módulos que creamos
from gui.helpers import (
    check_sharepoint_credentials,
    get_resource_path,
    get_email_google,
    get_microsoft_user_info,
)
import logging

# Configuración básica de logging
logging.basicConfig(
    level=logging.DEBUG,  # Cambia a DEBUG para obtener más detalles
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("app.log"),  # Guarda los logs en un archivo
        logging.StreamHandler()          # También los muestra en la consola
    ]
)
logger = logging.getLogger(__name__)

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
        else:
            html_content = '<html><body><h1>Error en la autenticación.</h1></body></html>'
            self.send_response(400)
            self.send_header('Content-type', 'text/html; charset=utf-8')
            self.end_headers()
            self.wfile.write(html_content.encode('utf-8'))

def open_user_manual_window():
    url = "https://alphak03.github.io/CHC_REPORTES_MANUAL-DE-USUARIO/"  # Reemplaza con el URL del manual de usuario
    webbrowser.open(url)


class Application(ttk.Window):
    def __init__(self):
        # Ajuste de DPI para Windows
        if sys.platform.startswith('win'):
            try:
                from ctypes import windll
                windll.shcore.SetProcessDpiAwareness(1)
            except Exception as e:
                print(f"No se pudo ajustar la conciencia de DPI: {e}")

        super().__init__(themename="flatly")  # Mantener el tema
        self.title("CHC REPORTES")

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
        """Obtiene el token para SharePoint usando el tenant específico."""
        scopes = ['https://chcconsultorescr.sharepoint.com/.default']  # Asegúrate de que este scope es correcto
        accounts = self.sharepoint_msal_app.get_accounts()
        result = None

        if accounts:
            result = self.sharepoint_msal_app.acquire_token_silent(scopes, account=accounts[0])

        if not result or 'access_token' not in result:
            result = self.sharepoint_msal_app.acquire_token_interactive(scopes=scopes, prompt='select_account')
            self.sharepoint_token_cache.save(self.sharepoint_msal_app)  # Guarda el token en caché después de autenticar

        if "access_token" in result:
            return result['access_token']
        else:
            error_message = result.get('error_description', 'No se pudo obtener el token de acceso para SharePoint.')
            messagebox.showerror("Error", f"Error al adquirir el token de acceso para SharePoint: {error_message}")
            return None

    def clear_email_cache(self):
        self.email_token_cache.clear_cache()

    def clear_sharepoint_cache(self):
        self.sharepoint_token_cache.clear_cache()

    def create_widgets(self):
        # Crear la barra de menú
        menu_bar = tk.Menu(self)
        self.config(menu=menu_bar)

        # Menú "Configuración"
        config_menu = tk.Menu(menu_bar, tearoff=0)
        config_menu.add_command(label="Configurar Cuenta", command=self.open_config_window)
        #config_menu.add_command(label="Cambiar correo de SharePoint", command=self.open_sharepoint_credentials_window)
        config_menu.add_command(label="Salir", command=self.on_closing)
        menu_bar.add_cascade(label="Configuración", menu=config_menu)

        # Menú "Gestión"
        gestion_menu = tk.Menu(menu_bar, tearoff=0)
        gestion_menu.add_command(label="Gestor de Clientes", command=self.open_client_manager)
        gestion_menu.add_command(label="Importar Clientes(JSON)", command=self.import_clients)
        gestion_menu.add_command(label="Exportar Clientes(JSON)", command=self.export_clients)
        menu_bar.add_cascade(label="Gestión", menu=gestion_menu)

        # Nuevo Menú "Ayuda"
        ayuda_menu = tk.Menu(menu_bar, tearoff=0)
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
                self.logo_label.image = self.logo_photo  # Mantener la referencia para evitar que se libere la memoria

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

        # Configuración de columnas en el frame de botones
        frame_botones.columnconfigure(0, weight=1)
        frame_botones.columnconfigure(1, weight=1)

        # Botón para generar reportes
        boton_generar = ttk.Button(
            frame_botones,
            text="Generar Reportes",
            command=self.filter_and_generate,
            style='BlueZodiac.TButton'
        )
        boton_generar.grid(row=0, column=0, padx=5, pady=5, sticky='ew')

        # Botón para enviar correos
        boton_enviar = ttk.Button(
            frame_botones,
            text="Enviar Correos",
            command=self.send_mails,
            style='YellowSea.TButton'
        )
        boton_enviar.grid(row=0, column=1, padx=5, pady=5, sticky='ew')

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
                self.loading_frames = [ImageTk.PhotoImage(img) for img in
                                       ImageSequence.Iterator(Image.open(loading_gif_path))]
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

    def start_loading_animation(self):
        if self.loading_frames:
            self.loading_label.grid()  # Mostrar el label
            self.loading_index = 0
            self.animate_loading()

    import logging
    import requests

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
        version = "Versión 2.1.8"
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

    def animate_loading(self):
        frame = self.loading_frames[self.loading_index]
        self.loading_label.configure(image=frame)
        self.loading_index = (self.loading_index + 1) % len(self.loading_frames)
        self.loading_animation = self.loading_label.after(100, self.animate_loading)  # Actualizar cada 100 ms

    def stop_loading_animation(self):
        if self.loading_frames:
            self.loading_label.after_cancel(self.loading_animation)
            self.loading_label.grid_remove()  # Ocultar el label

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
        logger.debug("Iniciando la función open_config_window.")
        try:
            config_window = tk.Toplevel(self)
            logger.debug("Creada la ventana Toplevel para configuración.")

            config_window.title("Configuración de Correo y Autenticación")
            config_window.configure(background='#f0f0f0')  # Fondo suave y profesional
            logger.debug("Configurado el título y el fondo de la ventana de configuración.")

            # Ajustar el tamaño de la ventana según la resolución de pantalla
            screen_width = config_window.winfo_screenwidth()
            screen_height = config_window.winfo_screenheight()
            scaling_factor = min(screen_width / 1920,
                                 screen_height / 1080)  # Factor de escala basado en resolución 1920x1080
            logger.debug(
                f"Resolución de pantalla obtenida: {screen_width}x{screen_height}. Factor de escala: {scaling_factor:.2f}")

            window_width = min(int(screen_width * 0.3), 1200)  # 30% del ancho o máximo 1200 px
            window_height = min(int(screen_height * 0.4), 900)  # 40% de la altura o máximo 900 px
            logger.debug(f"Tamaño de ventana calculado: {window_width}x{window_height}")

            # Centrar la ventana
            self.center_window(config_window, window_width, window_height)
            logger.debug("Ventana de configuración centrada en la pantalla.")

            # Configuración del grid de la ventana
            config_window.columnconfigure(0, weight=1)
            logger.debug("Configurado el grid de la ventana principal.")

            # Frame para la información de autenticación
            auth_frame = ttk.Frame(config_window, padding=(20, 10))
            auth_frame.grid(row=0, column=0, sticky='nsew')
            auth_frame.columnconfigure(0, weight=1)
            logger.debug("Creado el frame para la información de autenticación.")

            # Etiqueta de título
            etiqueta_info = ttk.Label(
                auth_frame,
                text="Información de Autenticación de la Cuenta",
                style='CustomTitle.TLabel'
            )
            etiqueta_info.grid(row=0, column=0, pady=(0, 10), sticky='w')
            logger.debug("Añadida la etiqueta de título al frame de autenticación.")

            # Separador horizontal
            separador = ttk.Separator(auth_frame, orient='horizontal')
            separador.grid(row=1, column=0, sticky='ew', pady=(0, 10))
            logger.debug("Añadido el separador horizontal al frame de autenticación.")

            # Inicializar etiquetas como atributos de la instancia para acceso global
            self.etiqueta_email_actual = ttk.Label(
                auth_frame,
                text="Cargando información...",
                style='CustomValue.TLabel'
            )
            self.etiqueta_email_actual.grid(row=2, column=0, padx=int(5 * scaling_factor), pady=int(5 * scaling_factor),
                                            sticky='w')
            logger.debug("Creada la etiqueta para mostrar el email actual.")

            self.etiqueta_nombre_value = ttk.Label(
                auth_frame,
                text="",
                style='CustomValue.TLabel'
            )
            self.etiqueta_nombre_value.grid(row=3, column=0, padx=int(5 * scaling_factor), pady=int(5 * scaling_factor),
                                            sticky='w')
            logger.debug("Creada la etiqueta para mostrar el nombre de usuario.")

            # Obtener información del usuario autenticado
            logger.debug("Obteniendo información del usuario autenticado.")
            user_info = self.get_current_authenticated_user()

            if user_info:
                email_autenticado = user_info['email']
                nombre_usuario = user_info['name']
                self.etiqueta_email_actual.config(text=f"Cuenta autenticada: {email_autenticado}")
                self.etiqueta_nombre_value.config(text=f"Nombre de Usuario: {nombre_usuario}")
                logger.info(f"Usuario autenticado: {email_autenticado}, Nombre: {nombre_usuario}")
            else:
                self.etiqueta_email_actual.config(text="NO AUTENTICADO")
                self.etiqueta_nombre_value.config(text="")
                logger.warning("No se encontró información del usuario autenticado.")

            # Frame para la selección del proveedor de correo
            provider_frame = ttk.Frame(config_window, padding=(20, 10))
            provider_frame.grid(row=1, column=0, sticky='nsew')
            provider_frame.columnconfigure(0, weight=1)
            logger.debug("Creado el frame para la selección del proveedor de correo.")

            etiqueta_provider = ttk.Label(
                provider_frame,
                text="Proveedor de Correo:",
                style='CustomSubtitle.TLabel'
            )
            etiqueta_provider.grid(row=0, column=0, pady=(0, 5), sticky='w')
            logger.debug("Añadida la etiqueta para el proveedor de correo.")

            # Opciones del proveedor de correo
            proveedor_opciones = ['Google', 'Microsoft']
            current_provider = self.config_manager.get_email_provider()
            provider_var = tk.StringVar(value=current_provider if current_provider else 'Selecciona un proveedor')
            logger.debug(f"Proveedor de correo actual: {current_provider}")

            self.provider_menu = ttk.OptionMenu(provider_frame, provider_var, provider_var.get(), *proveedor_opciones)
            self.provider_menu.config(width=30)  # Ancho adecuado para las opciones
            self.provider_menu.grid(row=1, column=0, pady=(0, 10), sticky='w')
            logger.debug("Añadido el menú desplegable para seleccionar el proveedor de correo.")

            # Frame para los botones
            button_frame = ttk.Frame(config_window, padding=(20, 10))
            button_frame.grid(row=2, column=0, sticky='ew')
            button_frame.columnconfigure(0, weight=1)
            logger.debug("Creado el frame para los botones de acción.")

            # Definir estilos personalizados para una apariencia más profesional
            style = ttk.Style()
            style.configure('CustomTitle.TLabel', font=('Helvetica', 16, 'bold'), foreground='#333333')
            style.configure('CustomSubtitle.TLabel', font=('Helvetica', 12, 'bold'), foreground='#555555')
            style.configure('CustomValue.TLabel', font=('Helvetica', 12), foreground='#000000')
            style.configure('CustomError.TLabel', font=('Helvetica', 12, 'bold'), foreground='red')
            logger.debug("Definidos los estilos personalizados para los widgets.")

            # Definir funciones internas antes de asignarlas a los botones
            def iniciar_oauth2():
                logger.debug("Iniciando el proceso de autenticación OAuth2.")
                provider = provider_var.get()
                logger.info(f"Proveedor seleccionado para autenticación: {provider}")
                if provider not in ['Google', 'Microsoft']:
                    logger.warning("Proveedor de correo no válido seleccionado.")
                    messagebox.showwarning("Advertencia", "Por favor, selecciona un proveedor de correo.")
                    return
                self.config_manager.set_email_provider(provider)
                logger.debug(f"Proveedor de correo establecido en la configuración: {provider}")

                try:
                    if provider == 'Google':
                        logger.debug("Iniciando autenticación con Google.")
                        client_secret_path = get_resource_path(
                            'resources/client_secret_772362460829-n7pc2am5c1tkvp7l9jig619kqmtk2obi.apps.googleusercontent.com.json')
                        flow = InstalledAppFlow.from_client_secrets_file(
                            client_secret_path,
                            scopes=EmailSender.GOOGLE_SCOPES
                        )
                        creds = flow.run_local_server(port=0)
                        self.config_manager.set_google_credentials(creds)
                        logger.info("Credenciales de Google almacenadas correctamente.")
                        email_autenticado = get_email_google(creds)
                        if email_autenticado:
                            self.config_manager.set_email(email_autenticado)
                            self.etiqueta_email_actual.config(text=f"Cuenta autenticada: {email_autenticado}")
                            self.etiqueta_nombre_value.config(text="Nombre de Usuario: (No disponible)")
                            logger.info(f"Autenticación de Google exitosa para el email: {email_autenticado}")
                            messagebox.showinfo("Éxito", "Autenticación de Google exitosa.")
                        else:
                            logger.error("No se pudo obtener el correo electrónico del usuario de Google.")
                            messagebox.showerror("Error", "No se pudo obtener el correo electrónico del usuario.")

                    elif provider == 'Microsoft':
                        logger.debug("Iniciando autenticación con Microsoft.")
                        client_id = MICROSOFT_CLIENT_ID
                        tenant_id = MICROSOFT_TENANT_ID
                        authority = f'https://login.microsoftonline.com/{tenant_id}' if tenant_id else 'https://login.microsoftonline.com/common'
                        redirect_uri = REDIRECT_URI
                        scopes = EmailSender.MS_SCOPES
                        logger.debug(f"Configuración de MSAL para Microsoft con authority: {authority}")

                        app = self.email_msal_app  # Usar la instancia existente
                        result = app.acquire_token_silent(scopes, account=None)
                        logger.debug("Intentando adquirir token silencioso para Microsoft.")

                        if not result:
                            logger.debug("No se encontró token en caché, iniciando flujo de autenticación interactivo.")
                            auth_url = app.get_authorization_request_url(
                                scopes=scopes,
                                redirect_uri=redirect_uri,
                                response_mode='query',
                                prompt='select_account'
                            )
                            webbrowser.open(auth_url)
                            auth_code = []
                            logger.debug("Abierta la URL de autenticación en el navegador.")

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
                                        logger.debug("Código de autorización recibido de Microsoft.")
                                    else:
                                        html_content = '<html><body><h1>Error en la autenticación.</h1></body></html>'
                                        self.send_response(400)
                                        self.send_header('Content-type', 'text/html; charset=utf-8')
                                        self.end_headers()
                                        self.wfile.write(html_content.encode('utf-8'))
                                        logger.error("Error en la autenticación de Microsoft: Código no encontrado.")

                            # Iniciar el servidor y esperar la respuesta
                            with socketserver.TCPServer(("", 62190), lambda *args, **kwargs: AuthHandler(
                                    *args,
                                    auth_code_container=auth_code,
                                    **kwargs)) as httpd:
                                logger.debug(
                                    "Servidor HTTP iniciado para manejar el callback de autenticación de Microsoft.")
                                server_thread = threading.Thread(target=httpd.handle_request)
                                server_thread.start()
                                server_thread.join()
                                logger.debug("Servidor HTTP cerrado después de recibir la respuesta.")

                            if not auth_code:
                                logger.error("No se recibió el código de autorización de Microsoft.")
                                messagebox.showerror("Error", "No se recibió el código de autorización de Microsoft.")
                                return

                            result = app.acquire_token_by_authorization_code(auth_code[0], scopes=scopes,
                                                                             redirect_uri=redirect_uri)
                            self.email_token_cache.save(app)
                            logger.debug("Token de Microsoft adquirido y cacheado correctamente.")

                        if "access_token" in result:
                            user_info = self.get_current_authenticated_user()
                            if user_info:
                                email_autenticado = user_info['email']
                                nombre_usuario = user_info['name']
                                self.config_manager.set_email(email_autenticado)
                                self.etiqueta_email_actual.config(text=f"Cuenta autenticada: {email_autenticado}")
                                self.etiqueta_nombre_value.config(text=f"Nombre de Usuario: {nombre_usuario}")
                                logger.info(f"Autenticación de Microsoft exitosa para el email: {email_autenticado}")
                                messagebox.showinfo("Éxito", "Autenticación de Microsoft exitosa.")
                            else:
                                logger.error("No se pudo obtener la información del usuario de Microsoft.")
                                messagebox.showerror("Error", "No se pudo obtener la información del usuario.")
                        else:
                            error_description = result.get('error_description',
                                                           'No se pudo obtener el token de acceso.')
                            logger.error(f"Error al adquirir el token de acceso de Microsoft: {error_description}")
                            messagebox.showerror("Error", f"Error al adquirir el token de acceso: {error_description}")

                except Exception as e:
                    logger.exception("Error durante la autenticación OAuth2.")
                    messagebox.showerror("Error", f"Error durante la autenticación: {str(e)}")

            def cerrar_sesion():
                logger.debug("Iniciando el proceso de cierre de sesión.")
                try:
                    # Limpiar credenciales almacenadas
                    self.config_manager.clear_credentials()
                    self.email_token_cache.clear_cache()
                    self.sharepoint_token_cache.clear_cache()
                    self.etiqueta_email_actual.config(text="NO AUTENTICADO")
                    self.etiqueta_nombre_value.config(text="")
                    logger.info("Credenciales y cachés limpiados correctamente. Usuario desautenticado.")

                    messagebox.showinfo("Cierre de sesión", "Sesión cerrada. Por favor, vuelve a iniciar sesión.")
                    config_window.destroy()
                    logger.debug("Ventana de configuración cerrada después del cierre de sesión.")
                    self.open_config_window()
                    logger.debug("Reabierta la ventana de configuración para reflejar el estado desautenticado.")

                except Exception as e:
                    logger.exception("No se pudo cerrar sesión correctamente.")
                    messagebox.showerror("Error", f"No se pudo cerrar sesión: {str(e)}")

            # Botón para iniciar la autenticación OAuth2
            self.boton_autenticar = ttk.Button(
                button_frame,
                text="Iniciar Sesión con OAuth2",
                command=iniciar_oauth2,
                style='Success.TButton',
                width=25
            )
            self.boton_autenticar.grid(row=0, column=0, pady=(0, 10), sticky='ew')
            logger.debug("Añadido el botón para iniciar la autenticación OAuth2.")

            # Botón para cerrar sesión y reautenticar
            self.boton_cerrar_sesion = ttk.Button(
                button_frame,
                text="Cerrar Sesión",
                command=cerrar_sesion,
                style='Danger.TButton',
                width=25
            )
            self.boton_cerrar_sesion.grid(row=1, column=0, pady=(0, 10), sticky='ew')
            logger.debug("Añadido el botón para cerrar sesión.")

            # Asignar funciones a los botones
            # Ya se asignaron directamente al crear los botones, por lo que esta línea puede omitirse
            # self.boton_autenticar.config(command=iniciar_oauth2)
            # self.boton_cerrar_sesion.config(command=cerrar_sesion)
            logger.debug("Funciones asignadas a los botones de autenticación y cierre de sesión.")

            # Configurar la ventana para que sea modal
            config_window.transient(self)
            config_window.grab_set()
            self.wait_window(config_window)
            logger.debug("Configuración de la ventana como modal completada. Esperando a que se cierre.")

        except Exception as e:
            logger.exception("Error al abrir la ventana de configuración.")
            messagebox.showerror("Error", f"No se pudo abrir la ventana de configuración: {str(e)}")

    def filter_and_generate(self):
        site_url = 'https://chcconsultorescr.sharepoint.com/sites/CONSULTORAS'
        sharepoint_file_url = '/sites/CONSULTORAS/Shared Documents/Pruebas ASW/Proyecto Informe de Cargas en Transito/Informe de cargas Power BI.xlsm'

        # Obtener el token de acceso para SharePoint
        access_token = self.get_sharepoint_access_token()
        if not access_token:
            self.update_progress("No se pudo obtener un token de acceso para SharePoint.", style='Error.TLabel')
            return

        try:
            # Crear la instancia de FileProcessor pasando el token de acceso
            self.processor = FileProcessor(site_url, sharepoint_file_url, access_token)
            self.update_progress("", style='Info.TLabel')
            self.update_progress("********************************************************", style='Delete.TLabel')

            # Descargar el archivo
            self.update_progress("Descargando el informe...", style='Info.TLabel')
            self.update_idletasks()
            self.processor.download_file()

            self.update_progress("********************************************************", style='Delete.TLabel')

            # Ejecutar la macro
            self.update_progress("Actualizando el informe...", style='Info.TLabel')
            self.update_idletasks()
            self.processor.run_macro()

            self.update_progress("********************************************************", style='Delete.TLabel')

            # Generar reportes
            self.update_progress("Generando reportes...", style='Info.TLabel')
            self.update_idletasks()
            self.processor.process_file()

            self.update_progress("********************************************************", style='Delete.TLabel')

            # Mensaje de éxito
            self.update_progress("Reportes generados!!", style='Success.TLabel')
            self.stop_loading_animation()

        except Exception as e:
            self.update_progress(f"Error: {str(e)}", style='Error.TLabel')
        finally:
            # Asegurarse de cerrar el archivo si la instancia de processor existe
            if self.processor and hasattr(self.processor, 'close') and callable(self.processor.close):
                self.processor.close()
            self.stop_loading_animation()

    def send_mails(self):
        # Mostrar el GIF de carga y actualizar el mensaje de progreso
        self.start_loading_animation()
        self.update_progress("Verificando autenticación...", style='Info.TLabel')

        # Obtener la información del usuario autenticado
        user_info = self.get_current_authenticated_user()

        if not user_info:
            self.update_progress("No estás autenticado. Iniciando sesión...", style='Warning.TLabel')
            # Abrir la ventana de configuración para iniciar sesión
            self.open_config_window()
            # Después de cerrar la ventana de configuración, verificar nuevamente la autenticación
            user_info = self.get_current_authenticated_user()
            if not user_info:
                self.update_progress("Autenticación fallida. No se pueden enviar correos.", style='Error.TLabel')
                self.stop_loading_animation()
                return
            else:
                self.update_progress("Autenticación exitosa. Preparando para enviar correos...", style='Success.TLabel')

        # Continuar con el envío de correos si está autenticado
        try:
            email = user_info['email']
            if not email:
                self.update_progress("No se pudo obtener el correo electrónico del usuario.", style='Error.TLabel')
                self.stop_loading_animation()
                return

            # Seleccionar la lista para enviar correos
            listas = self.db.get_all_lists()
            if not listas:
                messagebox.showwarning("Advertencia", "No hay listas disponibles. Por favor, crea una lista primero.")
                self.update_progress("No hay listas disponibles para enviar correos.", style='Warning.TLabel')
                self.stop_loading_animation()
                return

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
            self.stop_loading_animation()

    def send_mails_to_list(self, list_id):
        self.start_loading_animation()
        self.update_progress("********************************************************", style='Delete.TLabel')

        self.update_progress("Enviando los reportes...", style='Info.TLabel')
        self.update_idletasks()

        try:
            if not self.processor:
                self.update_progress("", style='Info.TLabel')
                self.update_progress("Primero debes generar los reportes.", style='Warning.TLabel')
                self.stop_loading_animation()

                return

            clientes_con_correo = [
                (client_id, nombre, correo)
                for (client_id, nombre, correo) in self.db.get_clients_by_list(list_id)
                if correo and correo.lower().strip() != '(correo sin definir)'
            ]

            if not clientes_con_correo:
                self.update_progress("", style='Info.TLabel')
                self.update_progress("No hay clientes con correos asociados para enviar.", style='Warning.TLabel')

                self.stop_loading_animation()
                return

            sender = EmailSender(self.email_msal_app)
            errores_envio = []
            for client_id, cliente, correo in clientes_con_correo:
                archivo_path = self.processor.get_client_file(cliente)
                if archivo_path:
                    asunto = f"Reporte para {cliente}"
                    cuerpo_html = f"<h1>Estimado {cliente}</h1><p>Adjunto encontrará su reporte.</p>"
                    success, msg = self.email_sender.send_email(correo, asunto, cuerpo_html, archivo_path)
                    if not success:
                        if "Se requiere reautenticación" in msg:
                            # Solicitar reautenticación
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
                # Programar la limpieza del mensaje de progreso después de 3 segundos
                #self.after(3000, lambda: self.update_progress("", style='Info.TLabel'))
        except Exception as e:
            self.update_progress(f"Error: {str(e)}", style='Error.TLabel')
        finally:
            self.stop_loading_animation()

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

