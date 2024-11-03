# src/gui/app.py
import http.server
import os
import socketserver
import sqlite3
import sys

import threading
from urllib.parse import urlparse, parse_qs

import ttkbootstrap as ttk
import tkinter as tk
from tkinter import messagebox, scrolledtext
import webbrowser
import requests
import logging
from PIL import Image, ImageTk, ImageSequence  # Importar para manejar imágenes y GIFs
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
from utils.token_cache import FileTokenCache


# Credenciales de Microsoft OAuth2
MICROSOFT_CLIENT_ID = "077e19a4-ab53-4e55-8286-ec02dc3cf36c"
MICROSOFT_CLIENT_SECRET = "pDK8Q~V63IOhtEe6KKrOzz8waIApiaXYrhq1MaFL"
MICROSOFT_TENANT_ID = "common"
REDIRECT_URI = "http://localhost:62190/"

def check_sharepoint_credentials(username, password):
    try:
        # URL base de tu sitio de SharePoint
        site_url = 'https://chcconsultorescr.sharepoint.com/sites/CONSULTORAS'

        ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))

        # Intentar obtener información del sitio para validar las credenciales
        web = ctx.web
        ctx.load(web)
        ctx.execute_query()

        # Si la consulta se ejecuta correctamente, las credenciales son válidas
        return True
    except Exception as e:
        logging.error(f"Error de autenticación en SharePoint: {e}")
        return False


def init_styles(scaling_factor):
    style = ttk.Style()

    # Tamaños de fuente ajustados al escalado
    base_font_size = int(12 * scaling_factor)
    small_font_size = int(10 * scaling_factor)
    large_font_size = int(17 * scaling_factor)

    # Estilo para etiquetas
    style.configure(
        'Custom.TLabel',
        background='white',
        foreground='#0c1834',  # Blue Zodiac
        font=("Helvetica", base_font_size)
    )

    # Estilo para etiquetas de éxito
    style.configure(
        'Success.TLabel',
        background='white',
        foreground='#28a745',  # Color verde
        font=("Helvetica", base_font_size)
    )

    # Estilo para botones 'Blue Zodiac' con bordes curvos
    style.configure(
        'BlueZodiac.TButton',
        font=("Helvetica", small_font_size, "bold"),
        foreground='white',
        background='#0c1834',
        borderwidth=1,
        relief='solid'
    )
    style.map(
        'BlueZodiac.TButton',
        foreground=[('active', 'white')],
        background=[('active', '#0a1430')]
    )

    # Repite este ajuste para los demás estilos de botones
    # Estilo para botones 'Yellow Sea'
    style.configure(
        'YellowSea.TButton',
        font=("Helvetica", small_font_size, "bold"),
        foreground='white',
        background='#ffb000',
        borderwidth=1,
        relief='solid'
    )
    style.map(
        'YellowSea.TButton',
        foreground=[('active', 'white')],
        background=[('active', '#e09a00')]
    )

    # Estilo para botones 'Red Color'
    style.configure(
        'RedColor.TButton',
        font=("Helvetica", small_font_size, "bold"),
        foreground='white',
        background='#f70103',
        borderwidth=1,
        relief='solid'
    )
    style.map(
        'RedColor.TButton',
        foreground=[('active', 'white')],
        background=[('active', '#d60002')]
    )

    # Estilo para botones 'Green Success'
    style.configure(
        'GreenSuccess.TButton',
        font=("Helvetica", small_font_size, "bold"),
        foreground='white',
        background='#28a745',
        borderwidth=1,
        relief='solid'
    )
    style.map(
        'GreenSuccess.TButton',
        foreground=[('active', 'white')],
        background=[('active', '#218838')]
    )

    # Estilo para botones 'Warning'
    style.configure(
        'Warning.TButton',
        font=("Helvetica", small_font_size, "bold"),
        foreground='black',
        background='#ffc107',
        borderwidth=1,
        relief='solid'
    )
    style.map(
        'Warning.TButton',
        foreground=[('active', 'black')],
        background=[('active', '#e0a800')]
    )

    # Estilo para etiquetas informativas
    style.configure(
        'Info.TLabel',
        background='white',
        foreground='#0c1834',  # Blue Zodiac
        font=("Helvetica", base_font_size)
    )

def get_resource_path(relative_path):
    """Obtiene la ruta del recurso empaquetado o del entorno de desarrollo."""
    base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
    return os.path.join(base_path, relative_path)


def get_email_google(creds):
    try:
        service = build('people', 'v1', credentials=creds)
        profile = service.people().get(resourceName='people/me', personFields='emailAddresses').execute()
        email_addresses = profile.get('emailAddresses', [])
        if email_addresses:
            return email_addresses[0]['value']
        else:
            return None
    except Exception as e:
        logging.error(f"Error al obtener el correo electrónico: {str(e)}")
        return None


def get_microsoft_user_info(access_token):
    """
    Obtiene la información del usuario autenticado en Microsoft Graph.
    """
    try:
        headers = {'Authorization': f'Bearer {access_token}'}
        response = requests.get('https://graph.microsoft.com/v1.0/me', headers=headers)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        logging.error(f"Error al obtener información del usuario de Microsoft: {str(e)}")
        return {}


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

        # Ajuste del factor de escalado
        self.scaling_factor = self.get_scaling_factor()
        self.tk.call('tk', 'scaling', self.scaling_factor)

        # Ajustar el tamaño inicial de la ventana según el escalado
        initial_width = int(500 * self.scaling_factor)
        initial_height = int(400 * self.scaling_factor)
        self.geometry(f"{initial_width}x{initial_height}")

        self.processor = None
        self.db = ClientDatabase()
        self.config_manager = ConfigManager()

        self.configure(background='white')  # Fondo blanco para la ventana

        # Inicialización del token_cache antes de msal_app y EmailSender
        self.token_cache = FileTokenCache()  # Inicialización del token_cache

        # Inicializar PublicClientApplication una sola vez
        self.msal_app = PublicClientApplication(
            client_id=MICROSOFT_CLIENT_ID,
            authority=f'https://login.microsoftonline.com/{MICROSOFT_TENANT_ID}',
            token_cache=self.token_cache.cache
        )

        # Inicializar EmailSender pasando la instancia de msal_app
        self.email_sender = EmailSender(self.msal_app)

        init_styles(self.scaling_factor)  # Inicializar estilos personalizados
        self.create_widgets()  # Crear widgets
        self.check_credentials()  # Verificar credenciales


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


    def create_widgets(self):
        # Crear la barra de menú
        menu_bar = tk.Menu(self)
        self.config(menu=menu_bar)

        # Menú "Configuración"
        config_menu = tk.Menu(menu_bar, tearoff=0)
        config_menu.add_command(label="Cambiar correo de origen", command=self.open_config_window)
        config_menu.add_command(label="Cambiar correo de SharePoint", command=self.open_sharepoint_credentials_window)
        config_menu.add_command(label="Salir", command=self.on_closing)
        menu_bar.add_cascade(label="Configuración", menu=config_menu)

        # Menú "Gestión"
        gestion_menu = tk.Menu(menu_bar, tearoff=0)
        gestion_menu.add_command(label="Gestor de Clientes", command=self.open_client_manager)
        menu_bar.add_cascade(label="Gestión", menu=gestion_menu)

        # **Nuevo Menú "Ayuda"**
        ayuda_menu = tk.Menu(menu_bar, tearoff=0)
        ayuda_menu.add_command(label="Acerca de", command=self.open_about_window)
        menu_bar.add_cascade(label="Ayuda", menu=ayuda_menu)

        # Configuración del grid de la ventana principal
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self.rowconfigure(1, weight=0)
        self.rowconfigure(2, weight=1)

        # Ruta de la imagen
        logo_path = get_resource_path('resources/Logo_CHC.png')

        # Cargar la imagen original
        self.original_logo_image = None
        if os.path.exists(logo_path):
            try:
                self.original_logo_image = Image.open(logo_path)
                self.logo_photo = ImageTk.PhotoImage(self.original_logo_image)

                self.logo_label = tk.Label(self, image=self.logo_photo, bg='white')
                self.logo_label.grid(row=0, column=0, padx=10, pady=10, sticky='n')
            except Exception as e:
                print(f"Error al cargar la imagen: {e}")
        else:
            print(f"Imagen no encontrada: {logo_path}")

        # Título principal
        titulo_principal = ttk.Label(
            self,
            text="Gestor de Reportes",
            style='Custom.TLabel',
            font=("Helvetica", 17, "bold")
        )
        titulo_principal.grid(row=1, column=0, padx=10, pady=10, sticky='n')

        # Frame de los botones
        frame_botones = tk.Frame(self, bg='white')
        frame_botones.grid(row=2, column=0, padx=10, pady=10, sticky='ew')

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

        # **Nuevo código agregado aquí**

        # Label para mostrar mensajes de progreso
        self.progress_label = ttk.Label(
            self,
            text="",
            style='Custom.TLabel',
            font=("Helvetica", 12)
        )
        self.progress_label.grid(row=3, column=0, padx=10, pady=5, sticky='n')



        # Cargar el GIF de carga
        loading_gif_path = get_resource_path('resources/loading.gif')
        if os.path.exists(loading_gif_path):
            self.loading_frames = [ImageTk.PhotoImage(img) for img in
                                   ImageSequence.Iterator(Image.open(loading_gif_path))]
            self.loading_label = tk.Label(self, bg='white')
            self.loading_label.grid(row=4, column=0, padx=10, pady=5, sticky='n')
            self.loading_label.grid_remove()  # Ocultar el label inicialmente
        else:
            self.loading_frames = None
            self.loading_label = None

        # Pie de página
        footer = ttk.Label(
            self,
            text="© 2024 CHC-Consultores. Todos los derechos reservados.",
            style='Custom.TLabel',
            font=("Helvetica", 10)
        )
        footer.grid(row=5, column=0, padx=10, pady=10, sticky='s')

        # Configurar el evento de redimensionamiento
        self.bind("<Configure>", self.on_resize)

    def start_loading_animation(self):
        if self.loading_frames:
            self.loading_label.grid()  # Mostrar el label
            self.loading_index = 0
            self.animate_loading()

    from tkinter import scrolledtext
    import requests

    def open_about_window(self):
        about_window = ttk.Toplevel(self)
        about_window.title("Acerca de")

        # Ajuste del tamaño según el escalado
        scaling_factor = self.scaling_factor
        width = int(400 * scaling_factor)
        height = int(300 * scaling_factor)
        about_window.geometry(f"{width}x{height}")
        about_window.configure(background='white')

        # Información de la aplicación
        app_name = "Gestor de Reportes"
        version = "Versión 2.4"
        company = "CHC-Consultores"

        # Etiqueta de nombre de la aplicación
        label_app_name = ttk.Label(
            about_window,
            text=app_name,
            style='Custom.TLabel',
            font=("Helvetica", int(14 * scaling_factor), "bold")
        )
        label_app_name.pack(pady=int(10 * scaling_factor))

        # Etiqueta de versión
        label_version = ttk.Label(
            about_window,
            text=version,
            style='Custom.TLabel',
            font=("Helvetica", int(12 * scaling_factor))
        )
        label_version.pack(pady=int(5 * scaling_factor))

        # Etiqueta de empresa
        label_company = ttk.Label(
            about_window,
            text=f"Desarrollado por {company}",
            style='Custom.TLabel',
            font=("Helvetica", int(12 * scaling_factor))
        )
        label_company.pack(pady=int(5 * scaling_factor))

        # Botón para abrir las políticas y términos en una nueva ventana
        boton_privacidad = ttk.Button(
            about_window,
            text="Política de Privacidad y Términos de Uso",
            command=self.open_privacy_terms_window,
            style='BlueZodiac.TButton'
        )
        boton_privacidad.pack(pady=int(20 * scaling_factor))

        # Hacer que la ventana sea modal
        about_window.transient(self)
        about_window.grab_set()
        self.wait_window(about_window)

    def open_privacy_terms_window(self):
        privacy_window = tk.Toplevel(self)
        privacy_window.geometry("800x800")
        privacy_window.configure(background='white')

        # Título centrado
        title_label = ttk.Label(
            privacy_window,
            font=("Helvetica", 16, "bold"),
            anchor="center"
        )
        title_label.pack(pady=10)

        # Obtener HTML de la página de GitHub
        try:
            response = requests.get("https://alphak03.github.io/CHC-REPORTES---POLITICAS-DE-PRIVACIDAD/")
            response.raise_for_status()
            content_html = response.text
        except requests.RequestException as e:
            content_html = f"<p>Error al cargar las políticas y términos: {e}</p>"

        # Crear un HTMLLabel para mostrar el contenido HTML
        html_label = HTMLLabel(privacy_window, html=content_html)
        html_label.pack(fill="both", expand=True, padx=10, pady=10)

        # Botón de cierre
        close_button = ttk.Button(
            privacy_window,
            text="Cerrar",
            command=privacy_window.destroy,
            style='RedColor.TButton'
        )
        close_button.pack(pady=10)

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
            Style (str): Estilo de la etiqueta. Valores posibles: 'Info.TLabel', 'Success.TLabel', etc.
        """
        self.progress_label.config(style=style, text=message)
        self.update_idletasks()  # Actualiza la interfaz para reflejar el cambio inmediatamente

    def resize_logo(self, max_width, max_height):
        """Redimensiona el logo manteniendo la proporción."""
        if self.original_logo_image:
            original_width, original_height = self.original_logo_image.size
            width_ratio = max_width / original_width
            height_ratio = max_height / original_height
            scaling_factor = min(width_ratio, height_ratio)

            new_width = int(original_width * scaling_factor)
            new_height = int(original_height * scaling_factor)

            return self.original_logo_image.resize((new_width, new_height), Image.Resampling.LANCZOS)
        return None

    def on_resize(self, event):
        """Ajusta el tamaño del logo cuando se redimensiona la ventana."""
        if event.widget == self:
            min_width, min_height = 140, 116  # Tamaño mínimo del logo
            max_width = event.width // 5  # Ajustar el ancho al 1/5 del ancho de la ventana
            max_height = event.height // 5  # Ajustar la altura al 1/5 de la altura de la ventana

            # Ajustar el tamaño del logo manteniendo el mínimo
            new_logo = self.resize_logo(max(max_width, min_width), max(max_height, min_height))
            if new_logo:
                self.logo_photo = ImageTk.PhotoImage(new_logo)
                self.logo_label.config(image=self.logo_photo)
                self.logo_label.image = self.logo_photo  # Actualizar la imagen en el widget

    def check_credentials(self):
        email = self.config_manager.get_email()
        app_password = self.config_manager.get_app_password()
        if not email or not app_password:
            self.open_config_window()

    def open_config_window(self):
        config_window = ttk.Toplevel(self)
        config_window.title("Configuración de Correo y Autenticación")
        config_window.geometry("500x300")
        config_window.configure(background='white')

        # Configuración del grid de la ventana
        config_window.columnconfigure(0, weight=1)

        # Etiqueta de información general
        etiqueta_info = ttk.Label(
            config_window,
            text="Información de autenticación de la cuenta:",
            style='Custom.TLabel'
        )
        etiqueta_info.grid(row=0, column=0, pady=10, sticky='w')

        # Mostrar el correo autenticado actual en texto plano, si existe
        email_actual = self.config_manager.get_email()
        etiqueta_email_actual = ttk.Label(
            config_window,
            text=f"Cuenta autenticada: {email_actual}" if email_actual else "No autenticado",
            style='Custom.TLabel',
            font=("Helvetica", 10)
        )
        etiqueta_email_actual.grid(row=1, column=0, padx=5, pady=5, sticky='w')

        # Selección del proveedor de correo
        current_provider = self.config_manager.get_email_provider()
        provider_var = tk.StringVar(value=current_provider if current_provider else 'Seleccione un proveedor')

        etiqueta_provider = ttk.Label(config_window, text="Proveedor de Correo:", style='Custom.TLabel')
        etiqueta_provider.grid(row=2, column=0, pady=5, sticky='w')

        proveedor_opciones = ['Google', 'Microsoft']
        provider_menu = ttk.OptionMenu(config_window, provider_var, provider_var.get(), *proveedor_opciones)
        provider_menu.grid(row=3, column=0, padx=5, pady=5, sticky='ew')

        # Autenticación OAuth2
        load_dotenv()

        # Botón para iniciar la autenticación OAuth2
        def iniciar_oauth2():
            provider = provider_var.get()
            if provider not in ['Google', 'Microsoft']:
                messagebox.showwarning("Advertencia", "Por favor, selecciona un proveedor de correo.")
                return
            self.config_manager.set_email_provider(provider)

            try:
                if provider == 'Google':
                    client_secret_path = get_resource_path(
                        'resources/client_secret_772362460829-n7pc2am5c1tkvp7l9jig619kqmtk2obi.apps.googleusercontent.com.json')
                    flow = InstalledAppFlow.from_client_secrets_file(
                        client_secret_path,
                        scopes=EmailSender.GOOGLE_SCOPES
                    )
                    creds = flow.run_local_server(port=0)
                    self.config_manager.set_google_credentials(creds)
                    email_autenticado = get_email_google(creds)
                    if email_autenticado:
                        self.config_manager.set_email(email_autenticado)
                        etiqueta_email_actual.config(text=f"Cuenta autenticada: {email_autenticado}")
                        messagebox.showinfo("Éxito", "Autenticación de Google exitosa.")
                    else:
                        messagebox.showerror("Error", "No se pudo obtener el correo electrónico del usuario.")

                elif provider == 'Microsoft':
                    client_id = MICROSOFT_CLIENT_ID
                    tenant_id = MICROSOFT_TENANT_ID
                    authority = f'https://login.microsoftonline.com/{tenant_id}' if tenant_id else 'https://login.microsoftonline.com/common'
                    redirect_uri = REDIRECT_URI
                    scopes = EmailSender.MS_SCOPES

                    app = self.msal_app  # Usar la instancia existente
                    result = app.acquire_token_silent(scopes, account=None)

                    if not result:
                        auth_url = app.get_authorization_request_url(scopes=scopes, redirect_uri=redirect_uri,
                                                                     response_mode='query', prompt='select_account')
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
                                else:
                                    html_content = '<html><body><h1>Error en la autenticación.</h1></body></html>'
                                    self.send_response(400)
                                    self.send_header('Content-type', 'text/html; charset=utf-8')
                                    self.end_headers()
                                    self.wfile.write(html_content.encode('utf-8'))

                        # Iniciar el servidor y esperar la respuesta
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
                        self.token_cache.save(app)

                    if "access_token" in result:
                        user_info = get_microsoft_user_info(result['access_token'])
                        email_autenticado = user_info.get('mail') or user_info.get('userPrincipalName')
                        if email_autenticado:
                            self.config_manager.set_email(email_autenticado)
                            etiqueta_email_actual.config(text=f"Cuenta autenticada: {email_autenticado}")
                            messagebox.showinfo("Éxito", "Autenticación de Microsoft exitosa.")
                        else:
                            messagebox.showerror("Error", "No se pudo obtener el correo electrónico del usuario.")
                    else:
                        error_description = result.get('error_description', 'No se pudo obtener el token de acceso.')
                        messagebox.showerror("Error", f"Error al adquirir el token de acceso: {error_description}")

            except Exception as e:
                logging.error(f"Error durante la autenticación: {e}")
                messagebox.showerror("Error", f"Error durante la autenticación: {str(e)}")

        boton_autenticar = ttk.Button(
            config_window,
            text="Iniciar Sesión con OAuth2",
            command=iniciar_oauth2,
            style='GreenSuccess.TButton'
        )
        boton_autenticar.grid(row=4, column=0, padx=5, pady=10, sticky='ew')

        # Botón para cerrar sesión y reautenticar
        def cerrar_sesion():
            try:
                # Limpiar credenciales almacenadas
                self.config_manager.clear_credentials()
                self.token_cache.clear_cache()
                etiqueta_email_actual.config(text="No autenticado")

                messagebox.showinfo("Cierre de sesión", "Sesión cerrada. Por favor, vuelve a iniciar sesión.")
                config_window.destroy()
                self.open_config_window()

            except Exception as e:
                logging.error(f"No se pudo cerrar sesión: {e}")
                messagebox.showerror("Error", f"No se pudo cerrar sesión: {str(e)}")

        boton_cerrar_sesion = ttk.Button(
            config_window,
            text="Cerrar Sesión",
            command=cerrar_sesion,
            style='RedColor.TButton'
        )
        boton_cerrar_sesion.grid(row=5, column=0, padx=5, pady=10, sticky='ew')

        config_window.columnconfigure(0, weight=1)
        config_window.transient(self)
        config_window.grab_set()
        self.wait_window(config_window)

    def open_sharepoint_credentials_window(self):
        config_window = ttk.Toplevel(self)
        config_window.title("Configurar Credenciales de SharePoint")
        config_window.geometry("500x300")
        config_window.configure(background='white')

        # Configuración del grid
        config_window.columnconfigure(0, weight=1)

        etiqueta_info = ttk.Label(
            config_window,
            text="Ingresa tus credenciales de SharePoint:",
            style='Custom.TLabel'
        )
        etiqueta_info.grid(row=0, column=0, pady=10, sticky='w')

        etiqueta_username = ttk.Label(config_window, text="Usuario de SharePoint:", style='Custom.TLabel')
        etiqueta_username.grid(row=1, column=0, pady=5, sticky='w')

        entrada_username = ttk.Entry(config_window)
        entrada_username.insert(0, self.config_manager.get_sharepoint_username())
        entrada_username.grid(row=2, column=0, padx=5, pady=5, sticky='ew')

        etiqueta_password = ttk.Label(config_window, text="Contraseña de SharePoint:", style='Custom.TLabel')
        etiqueta_password.grid(row=3, column=0, pady=5, sticky='w')

        # Configuración de la entrada de contraseña con alternancia de visibilidad
        frame_password = ttk.Frame(config_window)
        frame_password.grid(row=4, column=0, pady=5, sticky='ew')

        entrada_password = ttk.Entry(frame_password, show='*')
        entrada_password.grid(row=0, column=0, sticky='ew', padx=5, pady=5)

        # Cargar imágenes de los íconos de ojo usando rutas seguras
        try:
            eye_open_path = get_resource_path('resources/eye_open.png')
            eye_closed_path = get_resource_path('resources/eye_closed.png')

            # Utilizar el factor de escalado de config_window en lugar de crear una nueva instancia de ttk.Window()
            scaling_factor = config_window.tk.call('tk', 'scaling')
            icon_size = int(20 * scaling_factor)

            eye_open_img = Image.open(eye_open_path)
            eye_closed_img = Image.open(eye_closed_path)
            eye_open_photo = ImageTk.PhotoImage(eye_open_img.resize((icon_size, icon_size)))
            eye_closed_photo = ImageTk.PhotoImage(eye_closed_img.resize((icon_size, icon_size)))

        except Exception as e:
            messagebox.showerror("Error", f"Error al cargar iconos de visibilidad de contraseña: {e}")
            eye_open_photo = eye_closed_photo = None

        password_visible = False

        def toggle_password_visibility():
            nonlocal password_visible
            if password_visible:
                entrada_password.config(show='*')
                boton_toggle.config(image=eye_closed_photo)
                password_visible = False
            else:
                entrada_password.config(show='')
                boton_toggle.config(image=eye_open_photo)
                password_visible = True

        boton_toggle = ttk.Button(
            frame_password,
            image=eye_closed_photo,
            command=toggle_password_visibility,
            style='TButton'
        )
        boton_toggle.grid(row=0, column=1, padx=5, sticky='e')
        frame_password.columnconfigure(0, weight=1)

        # Función para guardar las credenciales
        def guardar_sharepoint_credentials():
            username = entrada_username.get().strip()
            password = entrada_password.get().strip()

            if not username or not password:
                messagebox.showwarning("Advertencia", "Por favor, completa todos los campos.")
                return

            try:
                if check_sharepoint_credentials(username, password):
                    self.config_manager.set_sharepoint_username(username)
                    self.config_manager.set_sharepoint_password(password)
                    config_window.destroy()
                    messagebox.showinfo("Éxito", "Credenciales de SharePoint guardadas exitosamente.")
                else:
                    messagebox.showerror("Error",
                                         "Las credenciales de SharePoint son incorrectas. Por favor, intenta nuevamente.")
            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error al verificar las credenciales: {e}")

        save_button = ttk.Button(config_window, text="Guardar", command=guardar_sharepoint_credentials,
                                 style='GreenSuccess.TButton')
        save_button.grid(row=5, column=0, padx=5, pady=15, sticky='ew')

        config_window.columnconfigure(0, weight=1)
        config_window.transient(self)
        config_window.grab_set()
        self.wait_window(config_window)

    def filter_and_generate(self):
        #self.update_progress("Iniciando proceso...")
        #self.start_loading_animation()

        # URLs fijas dentro del código
        site_url = 'https://chcconsultorescr.sharepoint.com/sites/CONSULTORAS'
        file_url = '/sites/CONSULTORAS/Shared Documents/Pruebas ASW/Proyecto Informe de Cargas en Transito/Informe de cargas en tránsito_Prueba conexión Power BI.xlsm'
        username = self.config_manager.get_sharepoint_username()
        password = self.config_manager.get_sharepoint_password()

        if not username or not password:
            self.update_progress("Por favor, configura tu usuario y contraseña de SharePoint.")
            self.stop_loading_animation()
            self.open_sharepoint_credentials_window()
            return

        try:
            # Crear la instancia de FileProcessor
            self.processor = FileProcessor(site_url, file_url)
            # Descargando el archivo}
            self.update_progress("Descargando informe...", style='Info.TLabel')
            self.update_idletasks()
            self.processor.download_file(username, password)

            # Ejecutar la macro
            self.update_progress("Actualizando informe...", style='Info.TLabel')
            self.update_idletasks()
            self.processor.run_macro()

            # Generar reportes
            self.update_progress("Generando reportes...", style='Info.TLabel')
            self.update_idletasks()
            self.processor.process_file()

            # Mensaje de éxito
            self.update_progress("Reportes generados!!", style='Success.TLabel')
            self.stop_loading_animation()


        except Exception as e:
            self.update_progress(f"Error: {str(e)}")
        finally:
            # Asegurarse de cerrar el archivo si la instancia de processor existe
            if self.processor and hasattr(self.processor, 'close') and callable(self.processor.close):
                self.processor.close()
            self.stop_loading_animation()

    def send_mails(self):
        email = self.config_manager.get_email()
        if not email:
            self.update_progress("Por favor, configura tu correo.")
            self.open_config_window()
            return

        # Seleccionar la lista para enviar correos
        listas = self.db.get_all_lists()
        if not listas:
            messagebox.showwarning("Advertencia", "No hay listas disponibles. Por favor, crea una lista primero.")
            return

        lista_nombres = [nombre for (id, nombre) in listas]
        list_var = tk.StringVar(value=lista_nombres[0] if lista_nombres else '')

        select_list_window = tk.Toplevel(self)
        select_list_window.title("Seleccionar Lista")
        select_list_window.geometry("300x150")
        select_list_window.configure(background='white')

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

    def send_mails_to_list(self, list_id):
        self.start_loading_animation()
        self.update_progress("Enviando los reportes...", style='Info.TLabel')
        self.update_idletasks()

        try:
            if not self.processor:
                self.update_progress("Primero debes generar los reportes.")
                self.stop_loading_animation()
                return

            clientes_con_correo = [
                (client_id, nombre, correo)
                for (client_id, nombre, correo) in self.db.get_clients_by_list(list_id)
                if correo and correo.lower().strip() != '(correo sin definir)'
            ]

            if not clientes_con_correo:
                self.update_progress("No hay clientes con correos asociados para enviar.")
                self.stop_loading_animation()
                return

            sender = EmailSender(self.token_cache)
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
                            self.update_progress("Reautenticación requerida. Por favor, inicia sesión nuevamente.",
                                                 style='Warning.TLabel')
                            self.open_config_window()
                            errores_envio.append(f"Reautenticación requerida para {cliente} ({correo}).")
                        else:
                            errores_envio.append(f"Error al enviar a {cliente} ({correo}): {msg}")
                else:
                    errores_envio.append(f"No se encontró el archivo para el cliente {cliente}.")

            if errores_envio:
                self.update_progress("Errores durante el envío de correos.")
                errores_detallados = "\n".join(errores_envio)
                messagebox.showerror("Errores de Envío", errores_detallados)
            else:
                self.update_progress("Reportes enviados exitosamente!", style='Success.TLabel')
        except Exception as e:
            self.update_progress(f"Error: {str(e)}")
        finally:
            self.stop_loading_animation()

    def open_client_manager(self):
        self.client_manager = ttk.Toplevel(self)
        self.client_manager.title("Gestor de Clientes")
        self.client_manager.geometry("1200x800")  # Tamaño adecuado
        self.client_manager.configure(background='white')

        # Configuración del grid principal
        self.client_manager.columnconfigure(0, weight=1)
        self.client_manager.rowconfigure(0, weight=1)

        # Crear el Notebook
        notebook = ttk.Notebook(self.client_manager)
        notebook.grid(row=0, column=0, sticky='nsew')

        # Pestaña 1: Por Lista
        tab_por_lista = ttk.Frame(notebook, padding=10)
        notebook.add(tab_por_lista, text="Por Lista")
        tab_por_lista.columnconfigure(0, weight=3)  # Clientes
        tab_por_lista.columnconfigure(1, weight=1)  # Aliases
        tab_por_lista.rowconfigure(3, weight=1)

        # Frame para la selección de listas
        frame_selection = ttk.Frame(tab_por_lista, padding=10)
        frame_selection.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky='ew')
        frame_selection.columnconfigure(1, weight=1)

        # Etiqueta para seleccionar lista
        label_select_lista = ttk.Label(
            frame_selection,
            text="Seleccionar Lista:",
            style='Custom.TLabel'
        )
        label_select_lista.grid(row=0, column=0, pady=5, sticky='w')

        # Obtener todas las listas
        self.listas = self.db.get_all_lists()
        self.list_names = [nombre for (id, nombre) in self.listas]
        self.list_ids = {nombre: id for (id, nombre) in self.listas}

        # Variable para la lista seleccionada
        self.selected_lista = tk.StringVar()
        if self.list_names:
            self.selected_lista.set(self.list_names[0])
        else:
            self.selected_lista.set('')

        # Combobox para seleccionar lista
        self.combobox_listas = ttk.Combobox(
            frame_selection,
            textvariable=self.selected_lista,
            values=self.list_names,
            state='readonly'
        )
        self.combobox_listas.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        self.combobox_listas.bind("<<ComboboxSelected>>", self.on_list_select_por_lista)

        # Botones para crear y eliminar listas
        boton_crear_lista = ttk.Button(
            frame_selection,
            text="Crear Lista",
            command=self.create_new_list,
            style='BlueZodiac.TButton'
        )
        boton_crear_lista.grid(row=1, column=0, padx=5, pady=5, sticky='ew')

        boton_eliminar_lista = ttk.Button(
            frame_selection,
            text="Eliminar Lista",
            command=self.delete_selected_list,
            style='RedColor.TButton'
        )
        boton_eliminar_lista.grid(row=1, column=1, padx=5, pady=5, sticky='ew')

        # Frame para la tabla de clientes por lista
        frame_clientes = ttk.Frame(tab_por_lista, padding=10)
        frame_clientes.grid(row=2, column=0, padx=10, pady=10, sticky='nsew')
        frame_clientes.columnconfigure(0, weight=1)
        frame_clientes.rowconfigure(0, weight=1)

        # Tabla para mostrar los clientes por lista
        columns = ('ID', 'Nombre', 'Correo')
        self.tree_clients_por_lista = ttk.Treeview(
            frame_clientes,
            columns=columns,
            show='headings',
            selectmode='browse'
        )
        self.tree_clients_por_lista.heading('ID', text='ID')
        self.tree_clients_por_lista.heading('Nombre', text='Nombre')
        self.tree_clients_por_lista.heading('Correo', text='Correo')
        self.tree_clients_por_lista.column('ID', width=50, anchor='center')
        self.tree_clients_por_lista.column('Nombre', width=200, anchor='w')
        self.tree_clients_por_lista.column('Correo', width=250, anchor='w')
        self.tree_clients_por_lista.grid(row=0, column=0, sticky='nsew')

        # Vincular evento para la selección de cliente en "Por Lista"
        self.tree_clients_por_lista.bind('<<TreeviewSelect>>', self.on_client_select)

        # Scrollbars para la tabla de clientes por lista
        scrollbar_vertical = ttk.Scrollbar(
            frame_clientes,
            orient=tk.VERTICAL,
            command=self.tree_clients_por_lista.yview
        )
        scrollbar_vertical.grid(row=0, column=1, sticky='ns')
        self.tree_clients_por_lista.configure(yscrollcommand=scrollbar_vertical.set)

        scrollbar_horizontal = ttk.Scrollbar(
            frame_clientes,
            orient=tk.HORIZONTAL,
            command=self.tree_clients_por_lista.xview
        )
        scrollbar_horizontal.grid(row=1, column=0, sticky='ew')
        self.tree_clients_por_lista.configure(xscrollcommand=scrollbar_horizontal.set)

        # Panel de botones de gestión de clientes por lista
        frame_botones_clientes = ttk.Frame(frame_clientes, padding=10)
        frame_botones_clientes.grid(row=2, column=0, padx=10, pady=5, sticky='ew')
        frame_botones_clientes.columnconfigure(0, weight=1)
        frame_botones_clientes.columnconfigure(1, weight=1)
        frame_botones_clientes.columnconfigure(2, weight=1)

        boton_anadir_cliente = ttk.Button(
            frame_botones_clientes,
            text="Añadir Cliente",
            command=self.add_client,
            style='BlueZodiac.TButton'
        )
        boton_anadir_cliente.grid(row=0, column=0, padx=5, pady=5, sticky='ew')

        boton_modificar_cliente = ttk.Button(
            frame_botones_clientes,
            text="Modificar Cliente",
            command=self.edit_client,
            style='Warning.TButton'
        )
        boton_modificar_cliente.grid(row=0, column=1, padx=5, pady=5, sticky='ew')

        boton_borrar_cliente = ttk.Button(
            frame_botones_clientes,
            text="Borrar Cliente",
            command=self.delete_client,
            style='RedColor.TButton'
        )
        boton_borrar_cliente.grid(row=0, column=2, padx=5, pady=5, sticky='ew')

        # Lista para mostrar los aliases
        frame_aliases = ttk.Frame(tab_por_lista, padding=10)
        frame_aliases.grid(row=3, column=0, padx=10, pady=10, sticky='nsew')
        frame_aliases.columnconfigure(0, weight=1)
        frame_aliases.rowconfigure(1, weight=1)

        label_aliases = ttk.Label(
            frame_aliases,
            text="Aliases del Cliente Seleccionado:",
            style='Custom.TLabel'
        )
        label_aliases.grid(row=0, column=0, pady=5, sticky='w')

        self.listbox_aliases_por_lista = tk.Listbox(frame_aliases)
        self.listbox_aliases_por_lista.grid(row=1, column=0, sticky='nsew')

        # Scrollbar para la lista de aliases por lista
        scrollbar_aliases = ttk.Scrollbar(
            frame_aliases,
            orient=tk.VERTICAL,
            command=self.listbox_aliases_por_lista.yview
        )
        scrollbar_aliases.grid(row=1, column=1, sticky='ns')
        self.listbox_aliases_por_lista.configure(yscrollcommand=scrollbar_aliases.set)

        # Panel de botones para gestión de aliases por lista
        frame_botones_aliases = ttk.Frame(frame_aliases, padding=10)
        frame_botones_aliases.grid(row=2, column=0, pady=5, sticky='ew')
        frame_botones_aliases.columnconfigure(0, weight=1)
        frame_botones_aliases.columnconfigure(1, weight=1)

        boton_anadir_alias = ttk.Button(
            frame_botones_aliases,
            text="Añadir Alias",
            command=self.add_alias,
            style='BlueZodiac.TButton'
        )
        boton_anadir_alias.grid(row=0, column=0, padx=5, pady=5, sticky='ew')

        boton_borrar_alias = ttk.Button(
            frame_botones_aliases,
            text="Borrar Alias",
            command=self.delete_alias,
            style='RedColor.TButton'
        )
        boton_borrar_alias.grid(row=0, column=1, padx=5, pady=5, sticky='ew')

        # Pestaña 2: "Todos los Clientes"
        tab_todos_clientes = ttk.Frame(notebook, padding=10)
        notebook.add(tab_todos_clientes, text="Todos los Clientes")
        tab_todos_clientes.columnconfigure(0, weight=1)
        tab_todos_clientes.rowconfigure(0, weight=1)

        # Frame para la lista de todos los clientes
        frame_all_clients = ttk.Frame(tab_todos_clientes, padding=10)
        frame_all_clients.grid(row=0, column=0, padx=10, pady=10, sticky='nsew')
        frame_all_clients.columnconfigure(0, weight=1)
        frame_all_clients.rowconfigure(0, weight=1)

        # Tabla para mostrar todos los clientes
        columns_all = ('ID', 'Nombre', 'Correo')
        self.tree_all_clients = ttk.Treeview(
            frame_all_clients,
            columns=columns_all,
            show='headings',
            selectmode='browse'
        )
        self.tree_all_clients.heading('ID', text='ID')
        self.tree_all_clients.heading('Nombre', text='Nombre')
        self.tree_all_clients.heading('Correo', text='Correo')
        self.tree_all_clients.column('ID', width=50, anchor='center')
        self.tree_all_clients.column('Nombre', width=200, anchor='w')
        self.tree_all_clients.column('Correo', width=250, anchor='w')
        self.tree_all_clients.grid(row=0, column=0, sticky='nsew')

        # Vincular evento para la selección de cliente en "Todos los Clientes"
        self.tree_all_clients.bind('<<TreeviewSelect>>', self.on_all_client_select)

        # Scrollbars para la tabla de todos los clientes
        scrollbar_vertical_all = ttk.Scrollbar(
            frame_all_clients,
            orient=tk.VERTICAL,
            command=self.tree_all_clients.yview
        )
        scrollbar_vertical_all.grid(row=0, column=1, sticky='ns')
        self.tree_all_clients.configure(yscrollcommand=scrollbar_vertical_all.set)

        scrollbar_horizontal_all = ttk.Scrollbar(
            frame_all_clients,
            orient=tk.HORIZONTAL,
            command=self.tree_all_clients.xview
        )
        scrollbar_horizontal_all.grid(row=1, column=0, sticky='ew')
        self.tree_all_clients.configure(xscrollcommand=scrollbar_horizontal_all.set)

        # Panel de botones para gestión de todos los clientes
        frame_botones_all_clients = ttk.Frame(frame_all_clients, padding=10)
        frame_botones_all_clients.grid(row=2, column=0, padx=10, pady=5, sticky='ew')
        frame_botones_all_clients.columnconfigure(0, weight=1)
        frame_botones_all_clients.columnconfigure(1, weight=1)
        frame_botones_all_clients.columnconfigure(2, weight=1)

        boton_anadir_cliente_all = ttk.Button(
            frame_botones_all_clients,
            text="Añadir Cliente",
            command=self.add_client,
            style='BlueZodiac.TButton'
        )
        boton_anadir_cliente_all.grid(row=0, column=0, padx=5, pady=5, sticky='ew')

        boton_modificar_cliente_all = ttk.Button(
            frame_botones_all_clients,
            text="Modificar Cliente",
            command=self.edit_client,
            style='Warning.TButton'
        )
        boton_modificar_cliente_all.grid(row=0, column=1, padx=5, pady=5, sticky='ew')

        boton_borrar_cliente_all = ttk.Button(
            frame_botones_all_clients,
            text="Borrar Cliente",
            command=self.delete_client,
            style='RedColor.TButton'
        )
        boton_borrar_cliente_all.grid(row=0, column=2, padx=5, pady=5, sticky='ew')

        # Lista para mostrar los aliases de todos los clientes
        self.listbox_aliases_all_clients = tk.Listbox(frame_all_clients)
        self.listbox_aliases_all_clients.grid(row=3, column=0, columnspan=2, sticky='nsew', padx=(0, 10))

        # Scrollbar para la lista de aliases de todos los clientes
        scrollbar_aliases_all_clients = ttk.Scrollbar(
            frame_all_clients,
            orient=tk.VERTICAL,
            command=self.listbox_aliases_all_clients.yview
        )
        scrollbar_aliases_all_clients.grid(row=3, column=1, sticky='ns', padx=(0, 10))
        self.listbox_aliases_all_clients.configure(yscrollcommand=scrollbar_aliases_all_clients.set)

        # Panel de botones para gestión de aliases de todos los clientes
        frame_botones_aliases_all_clients = ttk.Frame(frame_all_clients, padding=10)
        frame_botones_aliases_all_clients.grid(row=4, column=0, columnspan=3, pady=5, sticky='ew')
        frame_botones_aliases_all_clients.columnconfigure(0, weight=1)
        frame_botones_aliases_all_clients.columnconfigure(1, weight=1)

        boton_anadir_alias_all = ttk.Button(
            frame_botones_aliases_all_clients,
            text="Añadir Alias",
            command=self.add_alias,
            style='BlueZodiac.TButton'
        )
        boton_anadir_alias_all.grid(row=0, column=0, padx=5, pady=5, sticky='ew')

        boton_borrar_alias_all = ttk.Button(
            frame_botones_aliases_all_clients,
            text="Borrar Alias",
            command=self.delete_alias,
            style='RedColor.TButton'
        )
        boton_borrar_alias_all.grid(row=0, column=1, padx=5, pady=5, sticky='ew')

        # Cargar todos los clientes
        self.load_all_clients()

    def on_list_select(self, event):
        """Evento llamado al seleccionar una lista. Muestra y carga los detalles."""
        selected_list_name = self.selected_lista.get()
        if selected_list_name:
            self.frame_details.grid()  # Mostrar el frame_details
            self.load_clients_in_selected_list()
        else:
            self.frame_details.grid_remove()  # Ocultar el frame_details si no hay selección

    def load_clients_in_selected_list(self):
        """Carga los clientes asociados a la lista seleccionada en la tabla."""
        # Limpiar la tabla de clientes por lista
        for item in self.tree_clients_por_lista.get_children():
            self.tree_clients_por_lista.delete(item)

        try:
            selected_list_name = self.selected_lista.get()
            if selected_list_name and selected_list_name in self.list_ids:
                list_id = self.list_ids[selected_list_name]
                clientes = self.db.get_clients_by_list(list_id)
            else:
                clientes = []

            for count, (client_id, nombre, correo) in enumerate(clientes):
                if count % 2 == 0:
                    self.tree_clients_por_lista.insert(
                        '',
                        tk.END,
                        iid=client_id,
                        values=(client_id, nombre, correo),
                        tags=('evenrow',)
                    )
                else:
                    self.tree_clients_por_lista.insert(
                        '',
                        tk.END,
                        iid=client_id,
                        values=(client_id, nombre, correo),
                        tags=('oddrow',)
                    )
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error al cargar los clientes: {str(e)}")

        # Limpiar la lista de aliases por lista
        self.listbox_aliases_por_lista.delete(0, tk.END)

    def on_client_select(self, event):
        # Determinar la pestaña activa
        notebook = self.client_manager.nametowidget("!notebook")
        current_tab = notebook.tab(notebook.select(), "text")

        if current_tab == "Por Lista":
            tree = self.tree_clients_por_lista
            listbox = self.listbox_aliases_por_lista
        elif current_tab == "Todos los Clientes":
            tree = self.tree_all_clients
            listbox = self.listbox_aliases_all_clients
        else:
            messagebox.showwarning("Advertencia", "No se pudo determinar la pestaña activa.")
            return

        selected_item = tree.selection()
        if selected_item:
            item = selected_item[0]
            client_id = int(tree.item(item, 'values')[0])
            aliases = self.db.get_aliases_for_client(client_id)
            # Limpiar la lista de aliases
            listbox.delete(0, tk.END)
            # Agregar los aliases a la lista
            for alias_id, alias in aliases:
                listbox.insert(tk.END, f"{alias_id}:{alias}")

    def on_all_client_select(self, event):
        self.on_client_select(event)

    def add_client(self):
        def guardar_cliente():
            nombre = entrada_nombre.get().strip()
            correo = entrada_correo.get().strip()
            if not nombre:
                messagebox.showwarning("Advertencia", "El nombre del cliente no puede estar vacío.")
                return
            try:
                self.db.add_client(nombre, correo)
                client_id = self.db.get_client_id_by_name(nombre)
                if client_id:
                    selected_list_name = self.selected_lista.get()
                    if selected_list_name and selected_list_name in self.list_ids:
                        list_id = self.list_ids[selected_list_name]
                        self.db.add_client_to_list(client_id, list_id)
                # Determinar la pestaña activa para recargar la tabla correspondiente
                notebook = self.client_manager.nametowidget("!notebook")
                current_tab = notebook.tab(notebook.select(), "text")
                if current_tab == "Por Lista":
                    self.load_clients_in_selected_list()
                elif current_tab == "Todos los Clientes":
                    self.load_all_clients()
                ventana_agregar.destroy()
                messagebox.showinfo("Éxito", "Cliente agregado exitosamente.")
            except sqlite3.IntegrityError:
                messagebox.showerror("Error", f"El cliente '{nombre}' ya existe.")
            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error al agregar el cliente: {str(e)}")

        ventana_agregar = ttk.Toplevel(self.client_manager)
        ventana_agregar.title("Agregar Cliente")

        # Ajuste del tamaño según el escalado
        scaling_factor = self.scaling_factor
        width = int(400 * scaling_factor)
        height = int(250 * scaling_factor)
        ventana_agregar.geometry(f"{width}x{height}")
        ventana_agregar.configure(background='white')

        etiqueta_nombre = ttk.Label(ventana_agregar, text="Nombre del Cliente:", style='Custom.TLabel')
        etiqueta_nombre.pack(pady=int(5 * scaling_factor))
        entrada_nombre = ttk.Entry(ventana_agregar, width=50, font=("Helvetica", int(10 * scaling_factor)))
        entrada_nombre.pack(pady=5, padx=10, fill='x')

        etiqueta_correo = ttk.Label(ventana_agregar, text="Correo Electrónico:", style='Custom.TLabel')
        etiqueta_correo.pack(pady=int(5 * scaling_factor))
        entrada_correo = ttk.Entry(ventana_agregar, width=50, font=("Helvetica", int(10 * scaling_factor)))
        entrada_correo.pack(pady=5, padx=10, fill='x')

        boton_guardar = ttk.Button(
            ventana_agregar,
            text="Guardar",
            command=guardar_cliente,
            style='GreenSuccess.TButton'
        )
        boton_guardar.pack(pady=int(20 * scaling_factor), ipadx=int(10 * scaling_factor), ipady=int(5 * scaling_factor))

        ventana_agregar.transient(self.client_manager)
        ventana_agregar.grab_set()
        self.client_manager.wait_window(ventana_agregar)

    def edit_client(self):
        # Determinar la pestaña activa
        notebook = self.client_manager.nametowidget("!notebook")
        current_tab = notebook.tab(notebook.select(), "text")

        if current_tab == "Por Lista":
            tree = self.tree_clients_por_lista
            listbox = self.listbox_aliases_por_lista
        elif current_tab == "Todos los Clientes":
            tree = self.tree_all_clients
            listbox = self.listbox_aliases_all_clients
        else:
            messagebox.showwarning("Advertencia", "No se pudo determinar la pestaña activa.")
            return

        selected_item = tree.selection()
        if not selected_item:
            messagebox.showwarning("Advertencia", "Por favor, selecciona un cliente de la lista.")
            return

        item = selected_item[0]
        cliente_nombre = tree.item(item, 'values')[1]
        correo_actual = tree.item(item, 'values')[2]

        def guardar_cambio():
            nuevo_nombre = entrada_nombre.get().strip()
            nuevo_correo = entrada_correo.get().strip()

            if not nuevo_nombre:
                messagebox.showwarning("Advertencia", "El nombre del cliente no puede estar vacío.")
                return

            try:
                client_id = int(tree.item(item, 'values')[0])
                self.db.update_client_name_and_email(client_id, nuevo_nombre, nuevo_correo)

                # Recargar las tablas después de modificar el cliente
                if current_tab == "Por Lista":
                    self.load_clients_in_selected_list()
                elif current_tab == "Todos los Clientes":
                    self.load_all_clients()

                ventana_editar.destroy()
                messagebox.showinfo("Éxito", "Cliente actualizado exitosamente.")
            except sqlite3.IntegrityError:
                messagebox.showerror("Error", f"El cliente '{nuevo_nombre}' ya existe.")
            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error al actualizar el cliente: {str(e)}")

        # Ventana de edición de cliente
        ventana_editar = ttk.Toplevel(self.client_manager)
        ventana_editar.title(f"Editar Cliente - {cliente_nombre}")

        # Ajuste del tamaño según el escalado
        scaling_factor = self.scaling_factor
        width = int(400 * scaling_factor)
        height = int(250 * scaling_factor)
        ventana_editar.geometry(f"{width}x{height}")
        ventana_editar.configure(background='white')

        # Configuración del grid de la ventana
        ventana_editar.columnconfigure(0, weight=1)

        # Etiqueta y entrada para el nombre del cliente
        etiqueta_nombre = ttk.Label(
            ventana_editar,
            text="Nombre del Cliente:",
            style='Custom.TLabel'
        )
        etiqueta_nombre.grid(row=0, column=0, padx=10, pady=int(5 * scaling_factor), sticky='w')

        entrada_nombre = ttk.Entry(ventana_editar)
        entrada_nombre.insert(0, cliente_nombre)
        entrada_nombre.grid(row=1, column=0, padx=10, pady=int(5 * scaling_factor), sticky='ew')

        # Etiqueta y entrada para el correo del cliente
        etiqueta_correo = ttk.Label(
            ventana_editar,
            text="Correo Electrónico:",
            style='Custom.TLabel'
        )
        etiqueta_correo.grid(row=2, column=0, padx=10, pady=int(5 * scaling_factor), sticky='w')

        entrada_correo = ttk.Entry(ventana_editar)
        entrada_correo.insert(0, correo_actual)
        entrada_correo.grid(row=3, column=0, padx=10, pady=int(5 * scaling_factor), sticky='ew')

        # Botón para guardar los cambios
        boton_guardar = ttk.Button(
            ventana_editar,
            text="Guardar",
            command=guardar_cambio,
            style='GreenSuccess.TButton'
        )
        boton_guardar.grid(row=4, column=0, padx=10, pady=int(15 * scaling_factor), sticky='ew')

        # Configuración del grid para adaptabilidad
        ventana_editar.columnconfigure(0, weight=1)

        ventana_editar.transient(self.client_manager)
        ventana_editar.grab_set()
        self.client_manager.wait_window(ventana_editar)

    def delete_client(self):
        # Determinar la pestaña activa
        notebook = self.client_manager.nametowidget("!notebook")
        current_tab = notebook.tab(notebook.select(), "text")

        if current_tab == "Por Lista":
            tree = self.tree_clients_por_lista
            listbox = self.listbox_aliases_por_lista
        elif current_tab == "Todos los Clientes":
            tree = self.tree_all_clients
            listbox = self.listbox_aliases_all_clients
        else:
            messagebox.showwarning("Advertencia", "No se pudo determinar la pestaña activa.")
            return

        selected_item = tree.selection()
        if not selected_item:
            messagebox.showwarning("Advertencia", "Por favor, selecciona un cliente de la lista.")
            return

        item = selected_item[0]
        client_id = int(tree.item(item, 'values')[0])
        cliente = tree.item(item, 'values')[1]  # Suponiendo que 'Nombre' está en el índice 1

        respuesta = messagebox.askyesno("Confirmar",
                                        f"¿Estás seguro de eliminar el cliente '{cliente}' y todos sus aliases?")
        if respuesta:
            try:
                self.db.delete_client(client_id)

                # Recargar las tablas después de eliminar el cliente
                if current_tab == "Por Lista":
                    self.load_clients_in_selected_list()
                elif current_tab == "Todos los Clientes":
                    self.load_all_clients()

                # Limpiar la lista de aliases
                listbox.delete(0, tk.END)

                messagebox.showinfo("Éxito", "Cliente eliminado exitosamente.")
            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error al eliminar el cliente: {str(e)}")

    def add_alias(self):
        # Determinar la pestaña activa
        notebook = self.client_manager.nametowidget("!notebook")
        current_tab = notebook.tab(notebook.select(), "text")

        if current_tab == "Por Lista":
            tree = self.tree_clients_por_lista
            listbox = self.listbox_aliases_por_lista
        elif current_tab == "Todos los Clientes":
            tree = self.tree_all_clients
            listbox = self.listbox_aliases_all_clients
        else:
            messagebox.showwarning("Advertencia", "No se pudo determinar la pestaña activa.")
            return

        selected_item = tree.selection()
        if not selected_item:
            messagebox.showwarning("Advertencia", "Por favor, selecciona un cliente de la lista.")
            return

        item = selected_item[0]
        client_id = int(tree.item(item, 'values')[0])

        def guardar_alias():
            nuevo_alias = entrada_alias.get().strip()
            if not nuevo_alias:
                messagebox.showwarning("Advertencia", "El alias no puede estar vacío.")
                return
            try:
                self.db.add_alias(client_id, nuevo_alias)
                # Actualizar la lista de aliases
                aliases = self.db.get_aliases_for_client(client_id)
                listbox.delete(0, tk.END)
                for alias_id, alias in aliases:
                    listbox.insert(tk.END, f"{alias_id}:{alias}")
                ventana_alias.destroy()
                messagebox.showinfo("Éxito", "Alias agregado exitosamente.")
            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error al agregar el alias: {str(e)}")

        ventana_alias = ttk.Toplevel(self.client_manager)
        ventana_alias.title("Agregar Alias")
        ventana_alias.geometry("400x200")
        ventana_alias.configure(background='white')

        etiqueta_alias = ttk.Label(ventana_alias, text="Nuevo Alias:", style='Custom.TLabel')
        etiqueta_alias.pack(pady=5)
        entrada_alias = ttk.Entry(ventana_alias, width=50)
        entrada_alias.pack()

        boton_guardar = ttk.Button(ventana_alias, text="Guardar", command=guardar_alias,
                                   style='GreenSuccess.TButton')
        boton_guardar.pack(pady=20, ipadx=10, ipady=5)

        ventana_alias.transient(self.client_manager)
        ventana_alias.grab_set()
        self.client_manager.wait_window(ventana_alias)

    def delete_alias(self):
        # Determinar la pestaña activa
        notebook = self.client_manager.nametowidget("!notebook")
        current_tab = notebook.tab(notebook.select(), "text")

        if current_tab == "Por Lista":
            tree = self.tree_clients_por_lista
            listbox = self.listbox_aliases_por_lista
        elif current_tab == "Todos los Clientes":
            tree = self.tree_all_clients
            listbox = self.listbox_aliases_all_clients
        else:
            messagebox.showwarning("Advertencia", "No se pudo determinar la pestaña activa.")
            return

        selected_alias = listbox.curselection()
        if not selected_alias:
            messagebox.showwarning("Advertencia", "Por favor, selecciona un alias de la lista.")
            return

        alias_item = listbox.get(selected_alias[0])
        try:
            alias_id = int(alias_item.split(':')[0])
        except (IndexError, ValueError):
            messagebox.showerror("Error", "Formato de alias inválido.")
            return

        respuesta = messagebox.askyesno("Confirmar",
                                        f"¿Estás seguro de eliminar el alias '{alias_item.split(':')[1]}'?")
        if respuesta:
            try:
                self.db.remove_alias(alias_id)
                # Actualizar la lista de aliases
                selected_item = tree.selection()
                if selected_item:
                    client_id = int(tree.item(selected_item[0], 'values')[0])
                    aliases = self.db.get_aliases_for_client(client_id)
                    listbox.delete(0, tk.END)
                    for alias_id, alias in aliases:
                        listbox.insert(tk.END, f"{alias_id}:{alias}")
                messagebox.showinfo("Éxito", "Alias eliminado exitosamente.")
            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error al eliminar el alias: {str(e)}")

    def on_closing(self):
        # Cerrar la conexión a la base de datos al cerrar la aplicación
        if self.processor:
            self.processor.close()
        if self.db:
            self.db.close()
        self.destroy()

    def create_new_list(self):
        def guardar_lista():
            nombre_lista = entrada_nombre.get().strip()
            if not nombre_lista:
                messagebox.showwarning("Advertencia", "El nombre de la lista no puede estar vacío.")
                return
            try:
                self.db.add_list(nombre_lista)
                self.update_list_combobox()
                ventana_nueva_lista.destroy()
                messagebox.showinfo("Éxito", f"Lista '{nombre_lista}' creada exitosamente.")
            except sqlite3.IntegrityError:
                messagebox.showerror("Error", f"La lista '{nombre_lista}' ya existe.")
            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error al crear la lista: {str(e)}")

        ventana_nueva_lista = ttk.Toplevel(self.client_manager)
        ventana_nueva_lista.title("Crear Nueva Lista")
        ventana_nueva_lista.geometry("300x150")
        ventana_nueva_lista.configure(background='white')

        etiqueta_nombre = ttk.Label(ventana_nueva_lista, text="Nombre de la Lista:", style='Custom.TLabel')
        etiqueta_nombre.pack(pady=5)
        entrada_nombre = ttk.Entry(ventana_nueva_lista)
        entrada_nombre.pack(pady=5, padx=10, fill='x')

        boton_guardar = ttk.Button(
            ventana_nueva_lista,
            text="Guardar",
            command=guardar_lista,
            style='GreenSuccess.TButton'
        )
        boton_guardar.pack(pady=10)

        ventana_nueva_lista.transient(self.client_manager)
        ventana_nueva_lista.grab_set()
        self.client_manager.wait_window(ventana_nueva_lista)

    def delete_selected_list(self):
        selected_lista = self.selected_lista.get()
        if not selected_lista:
            messagebox.showwarning("Advertencia", "Por favor, selecciona una lista.")
            return
        if selected_lista == 'Predeterminada':
            messagebox.showwarning("Advertencia", "La lista 'Predeterminada' no puede ser eliminada.")
            return
        respuesta = messagebox.askyesno("Confirmar", f"¿Estás seguro de eliminar la lista '{selected_lista}'?")
        if respuesta:
            try:
                list_id = self.list_ids[selected_lista]
                self.db.delete_list(list_id)
                self.update_list_combobox()
                messagebox.showinfo("Éxito", f"Lista '{selected_lista}' eliminada exitosamente.")
            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error al eliminar la lista: {str(e)}")

    def update_list_combobox(self):
        self.listas = self.db.get_all_lists()
        self.list_names = [nombre for (id, nombre) in self.listas]
        self.list_ids = {nombre: id for (id, nombre) in self.listas}
        self.combobox_listas['values'] = self.list_names
        if self.list_names:
            self.selected_lista.set(self.list_names[0])
        else:
            self.selected_lista.set('')
        self.load_clients_in_selected_list()
        self.load_all_clients()

    def add_client_to_selected_list(self):
        selected_list_name = self.selected_lista.get()
        if not selected_list_name:
            messagebox.showwarning("Advertencia", "Por favor, selecciona una lista.")
            return

        selected_item = self.tree_all_clients.selection()
        if not selected_item:
            messagebox.showwarning("Advertencia", "Por favor, selecciona un cliente para agregar.")
            return

        client_id = int(selected_item[0])
        try:
            list_id = self.list_ids[selected_list_name]
            self.db.add_client_to_list(client_id, list_id)
            self.load_clients()
            self.load_clients_in_selected_list()
            messagebox.showinfo("Éxito", "Cliente agregado a la lista exitosamente.")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error al agregar el cliente a la lista: {str(e)}")

    def load_all_clients(self):
        """Carga todos los clientes en la tabla 'Todos los Clientes'."""
        # Limpiar la tabla de todos los clientes
        for item in self.tree_all_clients.get_children():
            self.tree_all_clients.delete(item)

        try:
            clientes = self.db.get_all_clients()
            for count, (client_id, nombre, correo) in enumerate(clientes):
                if count % 2 == 0:
                    self.tree_all_clients.insert(
                        '',
                        tk.END,
                        iid=client_id,
                        values=(client_id, nombre, correo),
                        tags=('evenrow',)
                    )
                else:
                    self.tree_all_clients.insert(
                        '',
                        tk.END,
                        iid=client_id,
                        values=(client_id, nombre, correo),
                        tags=('oddrow',)
                    )
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error al cargar todos los clientes: {str(e)}")

        # Limpiar la lista de aliases de todos los clientes
        if hasattr(self, 'listbox_aliases_all_clients'):
            self.listbox_aliases_all_clients.delete(0, tk.END)
        else:
            messagebox.showerror("Error", "No se encontró 'listbox_aliases_all_clients'.")

    def on_list_select_por_lista(self, event):
        # Limpiar la tabla de clientes por lista
        for item in self.tree_clients_por_lista.get_children():
            self.tree_clients_por_lista.delete(item)

        # Obtener la lista seleccionada y cargar los clientes asociados
        lista_nombre = self.selected_lista.get()
        if lista_nombre:
            lista_id = self.list_ids.get(lista_nombre)
            if lista_id:
                clientes = self.db.get_clients_by_list(lista_id)
                for client_id, nombre, correo in clientes:
                    self.tree_clients_por_lista.insert('', 'end', values=(client_id, nombre, correo))

