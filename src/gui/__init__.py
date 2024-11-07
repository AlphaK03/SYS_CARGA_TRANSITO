import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import msal
import requests
import os
from PIL import Image, ImageTk

# Configuración de Azure AD
MICROSOFT_CLIENT_ID = "077e19a4-ab53-4e55-8286-ec02dc3cf36c"
MICROSOFT_TENANT_ID = 'https://login.microsoftonline.com/f7015bc9-c229-413a-ab08-e1e5a76a09f1'  # Reemplaza <TU_TENANT_ID> con tu ID de tenant
REDIRECT_URI = 'http://localhost'  # URI de redirección permitido
SCOPE = ['https://chcconsultorescr.sharepoint.com/.default']  # Alcance para SharePoint REST API
CACHE_FILE = os.path.join(os.path.expanduser('~'), '.sharepoint_token_cache.json')

class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CHC REPORTES")

        # Configuración del tamaño y escalado
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        scaling_factor = min(screen_width / 1920, screen_height / 1080)
        window_width = min(int(screen_width * 0.3 * scaling_factor), 1200)
        window_height = min(int(screen_height * 0.4 * scaling_factor), 900)
        self.center_window(self, window_width, window_height)

        # Crear la barra de menú
        menu_bar = tk.Menu(self)
        self.config(menu=menu_bar)

        # Menú "Configuración"
        config_menu = tk.Menu(menu_bar, tearoff=0)
        config_menu.add_command(label="Iniciar sesión en SharePoint", command=self.open_sharepoint_credentials_window)
        config_menu.add_command(label="Cerrar Sesión de SharePoint", command=self.logout_sharepoint)
        config_menu.add_command(label="Salir", command=self.on_closing)
        menu_bar.add_cascade(label="Configuración", menu=config_menu)

        # Agregar otros elementos de la interfaz según tus necesidades
        self.create_main_interface()

    def create_main_interface(self):
        # Ejemplo de botón para descargar un archivo de SharePoint
        download_button = ttk.Button(self, text="Descargar Archivo de SharePoint", command=self.download_sharepoint_file)
        download_button.pack(pady=20)

    def open_sharepoint_credentials_window(self):
        # Crear una instancia de PublicClientApplication
        app = msal.PublicClientApplication(
            MICROSOFT_CLIENT_ID,
            authority=MICROSOFT_TENANT_ID,
            token_cache=msal.SerializableTokenCache()
        )

        # Cargar el caché de tokens si existe
        if os.path.exists(CACHE_FILE):
            with open(CACHE_FILE, 'r') as f:
                cache_data = f.read()
                app.token_cache.deserialize(cache_data)

        # Verificar si ya hay una cuenta autenticada
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(SCOPE, account=accounts[0])
            if result and 'access_token' in result:
                messagebox.showinfo("Autenticación", "Ya estás autenticado en SharePoint.")
                return

        # Iniciar el flujo de autenticación interactivo, abrirá automáticamente el navegador
        result = app.acquire_token_interactive(scopes=SCOPE)

        if 'access_token' in result:
            # Guardar el caché de tokens
            with open(CACHE_FILE, 'w') as f:
                f.write(app.token_cache.serialize())

            messagebox.showinfo("Éxito", "Autenticación exitosa en SharePoint.")
        else:
            messagebox.showerror("Error", f"No se pudo autenticar en SharePoint: {result.get('error_description')}")

    def logout_sharepoint(self):
        if os.path.exists(CACHE_FILE):
            os.remove(CACHE_FILE)
            messagebox.showinfo("Cerrar Sesión", "Se ha cerrado la sesión de SharePoint correctamente.")
        else:
            messagebox.showinfo("Cerrar Sesión", "No hay una sesión activa de SharePoint.")

    def get_sharepoint_access_token(self):
        # Borra el archivo de caché para forzar una nueva autenticación
        if os.path.exists(CACHE_FILE):
            os.remove(CACHE_FILE)

        app = msal.PublicClientApplication(
            MICROSOFT_CLIENT_ID,
            authority=MICROSOFT_TENANT_ID,
            token_cache=msal.SerializableTokenCache()
        )

        # Inicia la autenticación interactiva para asegurar un token actualizado
        result = app.acquire_token_interactive(scopes=SCOPE)

        if "access_token" in result:
            # Guarda el nuevo token en el caché
            with open(CACHE_FILE, 'w') as f:
                f.write(app.token_cache.serialize())
        else:
            messagebox.showerror("Error", "No se pudo obtener un token de acceso para SharePoint.")
            return None

        return result['access_token']

    def download_sharepoint_file(self):
        site_url = 'https://chcconsultorescr.sharepoint.com/sites/CONSULTORAS'
        file_relative_url = '/sites/CONSULTORAS/Shared Documents/Pruebas ASW/Proyecto Informe de Cargas en Transito/Informe de cargas en tránsito_Prueba conexión Power BI.xlsm'

        save_path = os.path.join('data', 'Informe.xlsm')

        access_token = self.get_sharepoint_access_token()
        if not access_token:
            return

        headers = {
            'Authorization': f'Bearer {access_token}',
            'Accept': 'application/json;odata=verbose'
        }

        # Construir la URL completa del archivo
        download_url = f"{site_url}/_api/web/GetFileByServerRelativeUrl('{file_relative_url}')/$value"

        try:
            response = requests.get(download_url, headers=headers)
            response.raise_for_status()
            os.makedirs(os.path.dirname(save_path), exist_ok=True)
            with open(save_path, 'wb') as f:
                f.write(response.content)
            messagebox.showinfo("Éxito", f"Archivo descargado exitosamente en {save_path}.")
        except requests.exceptions.RequestException as e:
            messagebox.showerror("Error", f"No se pudo descargar el archivo: {e}")

    def on_closing(self):
        self.destroy()

    def center_window(self, window, width, height):
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        window.geometry(f'{width}x{height}+{x}+{y}')

if __name__ == "__main__":
    app = Application()
    app.mainloop()
