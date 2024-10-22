# src/gui/app.py
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import webbrowser
import os
import re

from email_service.email_sender import EmailSender
from processing.file_processor import FileProcessor
from database.client_database import ClientDatabase
from utils.config_manager import ConfigManager


class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Gestor de Reportes y Correos a Clientes")
        self.geometry("600x500")
        self.configure(bg="#f0f4f7")  # Fondo de color pastel
        self.file_path = ''
        self.processor = None
        self.db = ClientDatabase()
        self.config_manager = ConfigManager()
        self.create_widgets()
        self.load_client_data()  # Cargar datos de clientes al iniciar
        self.check_credentials()  # Verificar si las credenciales están configuradas

    def create_widgets(self):
        # Crear la barra de menú
        menu_bar = tk.Menu(self)
        self.config(menu=menu_bar)

        # Menú "Archivo"
        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Cambiar correo de origen", command=self.open_config_window)
        file_menu.add_command(label="Salir", command=self.on_closing)
        menu_bar.add_cascade(label="Configuración", menu=file_menu)

        # Menú "Gestión"
        gestion_menu = tk.Menu(menu_bar, tearoff=0)
        gestion_menu.add_command(label="Gestor de Clientes", command=self.open_client_manager)
        menu_bar.add_cascade(label="Gestión", menu=gestion_menu)

        # Frame para seleccionar archivo
        frame_archivo = ttk.Frame(self)
        frame_archivo.pack(pady=10)

        etiqueta_archivo = ttk.Label(frame_archivo, text="Selecciona el archivo XLSX:")
        etiqueta_archivo.pack(side=tk.LEFT, padx=5)

        self.entrada_archivo = ttk.Entry(frame_archivo, width=40)
        self.entrada_archivo.pack(side=tk.LEFT, padx=5)

        boton_seleccionar = ttk.Button(frame_archivo, text="Seleccionar archivo", command=self.seleccionar_archivo)
        boton_seleccionar.pack(side=tk.LEFT, padx=5)

        # Botón para filtrar y generar reportes
        boton_generar = ttk.Button(self, text="Generar Reportes y Enviar Correos", command=self.filtrar_generar_enviar)
        boton_generar.pack(pady=10)

        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def check_credentials(self):
        email = self.config_manager.get_email()
        app_password = self.config_manager.get_app_password()
        if not email or not app_password:
            self.open_config_window()

    def open_config_window(self):
        # Ventana para configurar el correo y la contraseña
        config_window = tk.Toplevel(self)
        config_window.title("Configurar Correo y Contraseña de Aplicación")
        config_window.geometry("400x250")
        config_window.configure(bg="#f0f4f7")

        etiqueta_info = ttk.Label(config_window, text="Ingresa tu correo de Gmail y contraseña de aplicación:")
        etiqueta_info.pack(pady=10)

        # Campo para el correo electrónico
        etiqueta_email = ttk.Label(config_window, text="Correo Electrónico:")
        etiqueta_email.pack(pady=5)
        entrada_email = ttk.Entry(config_window, width=50)
        entrada_email.insert(0, self.config_manager.get_email())
        entrada_email.pack()

        # Campo para la contraseña de aplicación
        etiqueta_password = ttk.Label(config_window, text="Contraseña de Aplicación:")
        etiqueta_password.pack(pady=5)
        entrada_password = ttk.Entry(config_window, show='*', width=50)
        entrada_password.pack()

        # Enlace para obtener la contraseña de aplicación
        def abrir_link(event):
            webbrowser.open("https://support.google.com/accounts/answer/185833?hl=es-419")

        etiqueta_link = tk.Label(config_window, text="Obtener clave de aplicación con Google", fg="blue", cursor="hand2", bg="#f0f4f7")
        etiqueta_link.pack(pady=5)
        etiqueta_link.bind("<Button-1>", abrir_link)

        # Botón para guardar
        def guardar_config():
            email = entrada_email.get().strip()
            password = entrada_password.get().strip()
            if not email or not password:
                messagebox.showwarning("Advertencia", "Por favor, ingresa el correo y la contraseña.")
                return
            # Validar el formato del correo
            email_pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'
            if not re.match(email_pattern, email):
                messagebox.showwarning("Advertencia", "Por favor, ingresa un correo electrónico válido.")
                return
            self.config_manager.set_email(email)
            self.config_manager.set_app_password(password)
            config_window.destroy()
            messagebox.showinfo("Éxito", "Configuración guardada exitosamente.")

        boton_guardar = ttk.Button(config_window, text="Guardar", command=guardar_config)
        boton_guardar.pack(pady=10)

        config_window.transient(self)
        config_window.grab_set()
        self.wait_window(config_window)

    def seleccionar_archivo(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xlsm")])
        if self.file_path:
            self.entrada_archivo.delete(0, tk.END)
            self.entrada_archivo.insert(0, self.file_path)

    def filtrar_generar_enviar(self):
        if not self.file_path:
            messagebox.showwarning("Advertencia", "Por favor, selecciona un archivo.")
            return

        # Verificar credenciales
        email = self.config_manager.get_email()
        app_password = self.config_manager.get_app_password()
        if not email or not app_password:
            messagebox.showwarning("Advertencia", "Por favor, configura tu correo y contraseña de aplicación.")
            self.open_config_window()
            return

        try:
            self.processor = FileProcessor(self.file_path)
            clientes_unicos = self.processor.process_file()

            # Cargar los clientes en el gestor si está abierto
            if hasattr(self, 'client_manager') and self.client_manager.winfo_exists():
                self.load_clients()

            # Obtener clientes sin correo
            clientes_sin_correo = self.processor.get_clients_without_email()

            if clientes_sin_correo:
                messagebox.showwarning(
                    "Clientes sin correo",
                    f"Los siguientes clientes no tienen correo asociado y se omitieron:\n{', '.join(clientes_sin_correo)}"
                )

            # Obtener clientes con correo
            clientes_con_correo = self.processor.get_clients_with_email()

            if not clientes_con_correo:
                messagebox.showinfo("Información", "No hay clientes con correos asociados para enviar.")
                return

            sender = EmailSender()
            errores_envio = []
            for client_id, cliente, correo in clientes_con_correo:
                archivo_path = self.processor.get_client_file(cliente)
                if archivo_path:
                    asunto = f"Reporte para {cliente}"
                    cuerpo_html = f"<h1>Estimado {cliente}</h1><p>Adjunto encontrará su reporte.</p>"
                    success, msg = sender.send_email(
                        correo,
                        asunto,
                        cuerpo_html,
                        archivo_path
                    )
                    if not success:
                        errores_envio.append(f"Error al enviar a {cliente} ({correo}): {msg}")
                else:
                    errores_envio.append(f"No se encontró el archivo para el cliente {cliente}.")

            if errores_envio:
                messagebox.showerror("Errores de envío", "\n".join(errores_envio))
            else:
                messagebox.showinfo("Éxito", "Se enviaron los correos exitosamente.")

            self.processor.close()

        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error: {str(e)}")

    def open_client_manager(self):
        # [Código existente para el gestor de clientes]
        # No es necesario cambiar esta parte
        pass

    def load_client_data(self):
        # [Código existente para cargar los clientes]
        pass

    def on_closing(self):
        # Cerrar la conexión a la base de datos al cerrar la aplicación
        if self.processor:
            self.processor.close()
        if self.db:
            self.db.close()
        self.destroy()
