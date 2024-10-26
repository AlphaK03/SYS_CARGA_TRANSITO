# src/gui/app.py

import os

import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import tkinter as tk
from tkinter import messagebox
import webbrowser
import re
from PIL import Image, ImageTk  # Importar para manejar imágenes

from email_service.email_sender import EmailSender
from processing.file_processor import FileProcessor
from database.client_database import ClientDatabase
from utils.config_manager import ConfigManager


class Application(ttk.Window):
    def __init__(self):
        super().__init__(themename="flatly")
        self.title("Proyecto ASW")
        self.geometry("550x350")
        self.processor = None
        self.db = ClientDatabase()
        self.config_manager = ConfigManager()

        self.configure(background='white')  # Fondo blanco para la ventana

        self.init_styles()  # Inicializar estilos personalizados
        self.create_widgets()
        self.check_credentials()

    def init_styles(self):
        style = ttk.Style()

        # Estilo para etiquetas
        style.configure(
            'Custom.TLabel',
            background='white',
            foreground='#0c1834',  # Blue Zodiac
            font=("Helvetica", 12)
        )

        # Estilo para botones 'Blue Zodiac'
        style.configure(
            'BlueZodiac.TButton',
            background='#0c1834',
            foreground='white',
            bordercolor='#0c1834',
            focusthickness=3,
            focuscolor='#0c1834',
            font=("Helvetica", 10, "bold")
        )

        # Estilo para botones 'Yellow Sea'
        style.configure(
            'YellowSea.TButton',
            background='#ffb000',
            foreground='white',
            bordercolor='#ffb000',
            focusthickness=3,
            focuscolor='#ffb000',
            font=("Helvetica", 10, "bold")
        )

        # Estilo para botones 'Red Color'
        style.configure(
            'RedColor.TButton',
            background='#f70103',
            foreground='white',
            bordercolor='#f70103',
            focusthickness=3,
            focuscolor='#f70103',
            font=("Helvetica", 10, "bold")
        )

        # Estilo para botones 'Green Success'
        style.configure(
            'GreenSuccess.TButton',
            background='#28a745',
            foreground='white',
            bordercolor='#28a745',
            focusthickness=3,
            focuscolor='#28a745',
            font=("Helvetica", 10, "bold")
        )

        # Estilo para botones 'Warning'
        style.configure(
            'Warning.TButton',
            background='#ffc107',
            foreground='black',
            bordercolor='#ffc107',
            focusthickness=3,
            focuscolor='#ffc107',
            font=("Helvetica", 10, "bold")
        )

    def create_widgets(self):
        # Crear la barra de menú
        menu_bar = tk.Menu(self)
        self.config(menu=menu_bar)

        # Menú "Configuración"
        config_menu = tk.Menu(menu_bar, tearoff=0)
        config_menu.add_command(label="Cambiar correo de origen", command=self.open_config_window)
        config_menu.add_command(label="Configurar SharePoint", command=self.open_sharepoint_config_window)
        config_menu.add_command(label="Salir", command=self.on_closing)
        menu_bar.add_cascade(label="Configuración", menu=config_menu)

        # Menú "Gestión"
        gestion_menu = tk.Menu(menu_bar, tearoff=0)
        gestion_menu.add_command(label="Gestor de Clientes", command=self.open_client_manager)
        menu_bar.add_cascade(label="Gestión", menu=gestion_menu)

        # Ruta de la imagen
        script_dir = os.path.dirname(os.path.abspath(__file__))
        logo_path = os.path.join(script_dir, '..', '..', 'resources', 'Logo_CHC.png')

        if os.path.exists(logo_path):
            try:
                logo_image = Image.open(logo_path)
                logo_image = logo_image.resize((140, 116), Image.Resampling.LANCZOS)  # Usar LANCZOS en lugar de ANTIALIAS
                logo_photo = ImageTk.PhotoImage(logo_image)
                logo_label = tk.Label(self, image=logo_photo, bg='white')
                logo_label.image = logo_photo
                logo_label.pack(pady=10)
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
        titulo_principal.pack(pady=10)

        # Botones para filtrar/generar y enviar correos
        frame_botones = tk.Frame(self, bg='white')
        frame_botones.pack(pady=20)

        boton_generar = ttk.Button(
            frame_botones,
            text="Generar Reportes",
            command=self.filtrar_generar,
            style='BlueZodiac.TButton'
        )
        boton_generar.pack(side=tk.LEFT, padx=20, ipadx=10, ipady=10)

        boton_enviar = ttk.Button(
            frame_botones,
            text="Enviar Correos",
            command=self.enviar_correos,
            style='YellowSea.TButton'
        )
        boton_enviar.pack(side=tk.LEFT, padx=20, ipadx=10, ipady=10)

        # Pie de página
        footer = ttk.Label(
            self,
            text="© 2024 CHC-Consultores. Todos los derechos reservados.",
            style='Custom.TLabel',
            font=("Helvetica", 10)
        )
        footer.pack(side=tk.BOTTOM, pady=10)

        self.protocol("WM_DELETE_WINDOW", self.on_closing)


    def check_credentials(self):
        email = self.config_manager.get_email()
        app_password = self.config_manager.get_app_password()
        if not email or not app_password:
            self.open_config_window()

    def open_config_window(self):
        # Ventana para configurar el correo y la contraseña
        config_window = ttk.Toplevel(self)
        config_window.title("Configurar Correo y Contraseña de Aplicación")
        config_window.geometry("400x400")
        config_window.configure(background='white')

        etiqueta_info = ttk.Label(config_window, text="Ingresa tu correo y contraseña de aplicación:",
                                  style='Custom.TLabel')
        etiqueta_info.pack(pady=10)

        # Selección del proveedor de correo
        etiqueta_provider = ttk.Label(config_window, text="Proveedor de Correo:", style='Custom.TLabel')
        etiqueta_provider.pack(pady=5)
        provider_var = tk.StringVar(value=self.config_manager.get_email_provider())
        proveedor_opciones = ['Gmail', 'Microsoft']
        provider_menu = ttk.OptionMenu(config_window, provider_var, provider_var.get(), *proveedor_opciones)
        provider_menu.pack()

        # Campo para el correo electrónico
        etiqueta_email = ttk.Label(config_window, text="Correo Electrónico:", style='Custom.TLabel')
        etiqueta_email.pack(pady=5)
        entrada_email = ttk.Entry(config_window, width=50)
        entrada_email.insert(0, self.config_manager.get_email())
        entrada_email.pack()

        # Campo para la contraseña de aplicación
        etiqueta_password = ttk.Label(config_window, text="Contraseña de Aplicación:", style='Custom.TLabel')
        etiqueta_password.pack(pady=5)
        entrada_password = ttk.Entry(config_window, show='*', width=50)
        entrada_password.pack()

        # Enlace para obtener la contraseña de aplicación
        def abrir_link(event):
            if provider_var.get() == 'Gmail':
                webbrowser.open("https://support.google.com/accounts/answer/185833?hl=es-419")
            elif provider_var.get() == 'Microsoft':
                webbrowser.open(
                    "https://support.microsoft.com/es-es/account-billing/crear-una-contrasena-de-aplicacion-3ea7a3a3-1899-4d5c-9f95-09770406a3d2")

        etiqueta_link = ttk.Label(config_window, text="Obtener clave de aplicación",
                                  foreground="blue", cursor="hand2", background='white')
        etiqueta_link.pack(pady=5)
        etiqueta_link.bind("<Button-1>", abrir_link)

        # Función para guardar la configuración
        def guardar_config():
            email = entrada_email.get().strip()
            password = entrada_password.get().strip()
            provider = provider_var.get()
            if not email or not password:
                messagebox.showwarning("Advertencia", "Por favor, ingresa el correo y la contraseña.")
                return
            # Validar el formato del correo
            email_pattern = r'^[\w\.-]+@[\w\.-]+\.\w+$'
            if not re.match(email_pattern, email):
                messagebox.showwarning("Advertencia", "Por favor, ingresa un correo electrónico válido.")
                return
            self.config_manager.set_email_provider(provider)
            self.config_manager.set_email(email)
            self.config_manager.set_app_password(password)
            config_window.destroy()
            messagebox.showinfo("Éxito", "Configuración guardada exitosamente.")

        # Botón para guardar
        boton_guardar = ttk.Button(config_window, text="Guardar", command=guardar_config, style='GreenSuccess.TButton')
        boton_guardar.pack(pady=20, ipadx=10, ipady=5)

        config_window.transient(self)
        config_window.grab_set()
        self.wait_window(config_window)

    def open_sharepoint_config_window(self):
        config_window = ttk.Toplevel(self)
        config_window.title("Configurar SharePoint")
        config_window.geometry("400x400")
        config_window.configure(background='white')

        etiqueta_info = ttk.Label(config_window, text="Ingresa los detalles de SharePoint:", style='Custom.TLabel')
        etiqueta_info.pack(pady=10)

        # Campo para el sitio de SharePoint
        etiqueta_site_url = ttk.Label(config_window, text="URL del sitio de SharePoint:", style='Custom.TLabel')
        etiqueta_site_url.pack(pady=5)
        entrada_site_url = ttk.Entry(config_window, width=50)
        entrada_site_url.insert(0, self.config_manager.get_sharepoint_site_url())
        entrada_site_url.pack()

        # Campo para el usuario de SharePoint
        etiqueta_username = ttk.Label(config_window, text="Usuario de SharePoint:", style='Custom.TLabel')
        etiqueta_username.pack(pady=5)
        entrada_username = ttk.Entry(config_window, width=50)
        entrada_username.insert(0, self.config_manager.get_sharepoint_username())
        entrada_username.pack()

        # Campo para la contraseña de SharePoint
        etiqueta_password = ttk.Label(config_window, text="Contraseña de SharePoint:", style='Custom.TLabel')
        etiqueta_password.pack(pady=5)
        entrada_password = ttk.Entry(config_window, show='*', width=50)
        entrada_password.pack()

        # Campo para el archivo en SharePoint
        etiqueta_file_url = ttk.Label(config_window, text="URL relativa del archivo en SharePoint:",
                                      style='Custom.TLabel')
        etiqueta_file_url.pack(pady=5)
        entrada_file_url = ttk.Entry(config_window, width=50)
        entrada_file_url.insert(0, self.config_manager.get_sharepoint_file_url())
        entrada_file_url.pack()

        # Función para guardar la configuración
        def guardar_sharepoint_config():
            site_url = entrada_site_url.get().strip()
            username = entrada_username.get().strip()
            password = entrada_password.get().strip()
            file_url = entrada_file_url.get().strip()
            if not site_url or not username or not password or not file_url:
                messagebox.showwarning("Advertencia", "Por favor, completa todos los campos.")
                return
            self.config_manager.set_sharepoint_site_url(site_url)
            self.config_manager.set_sharepoint_username(username)
            self.config_manager.set_sharepoint_password(password)
            self.config_manager.set_sharepoint_file_url(file_url)
            config_window.destroy()
            messagebox.showinfo("Éxito", "Configuración de SharePoint guardada exitosamente.")

        # Botón para guardar
        boton_guardar = ttk.Button(config_window, text="Guardar", command=guardar_sharepoint_config,
                                   style='GreenSuccess.TButton')
        boton_guardar.pack(pady=20, ipadx=10, ipady=5)

        config_window.transient(self)
        config_window.grab_set()
        self.wait_window(config_window)

    def filtrar_generar(self):
        site_url = self.config_manager.get_sharepoint_site_url()
        username = self.config_manager.get_sharepoint_username()
        password = self.config_manager.get_sharepoint_password()
        file_url = self.config_manager.get_sharepoint_file_url()

        if not site_url or not username or not password or not file_url:
            messagebox.showwarning("Advertencia", "Por favor, configura los detalles de SharePoint.")
            self.open_sharepoint_config_window()
            return

        try:
            self.processor = FileProcessor(file_url)
            # Descargar el archivo
            self.processor.download_file()

            # Ejecutar la macro 'Name'
            self.processor.run_macro()
            messagebox.showinfo("Éxito", "La macro se ejecutó correctamente.")

            # Después de ejecutar la macro, continuar con la generación de reportes
            clientes_procesados = self.processor.process_file()
            messagebox.showinfo("Éxito", "Se generaron los reportes exitosamente.")

        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error: {str(e)}")


    def enviar_correos(self):
        email = self.config_manager.get_email()
        app_password = self.config_manager.get_app_password()
        if not email or not app_password:
            messagebox.showwarning("Advertencia", "Por favor, configura tu correo y contraseña de aplicación.")
            self.open_config_window()
            return

        try:
            if not self.processor:
                messagebox.showwarning("Advertencia", "Primero debes generar los reportes.")
                return

            clientes_con_correo = self.db.get_clients_with_email()

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
                    success, msg = sender.send_email(correo, asunto, cuerpo_html, archivo_path)
                    if not success:
                        errores_envio.append(f"Error al enviar a {cliente} ({correo}): {msg}")
                else:
                    errores_envio.append(f"No se encontró el archivo para el cliente {cliente}.")

            if errores_envio:
                messagebox.showerror("Errores de envío", "\n".join(errores_envio))
            else:
                messagebox.showinfo("Éxito", "Se enviaron los correos exitosamente.")

        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error: {str(e)}")

    def open_client_manager(self):
        # Ventana del gestor de clientes
        self.client_manager = ttk.Toplevel(self)
        self.client_manager.title("Gestor de Clientes")
        self.client_manager.geometry("1000x600")
        self.client_manager.configure(background='white')

        # Configuración del grid para la ventana principal
        self.client_manager.columnconfigure(0, weight=1)
        self.client_manager.columnconfigure(1, weight=0)
        self.client_manager.columnconfigure(2, weight=1)
        self.client_manager.rowconfigure(0, weight=1)

        # ----------- Frame de la tabla de clientes -----------
        frame_clientes = ttk.Frame(self.client_manager)
        frame_clientes.grid(row=0, column=0, padx=10, pady=10, sticky='nsew')

        # Tabla para mostrar los clientes
        columns = ('ID', 'Nombre', 'Correo')
        self.tree_clients = ttk.Treeview(
            frame_clientes,
            columns=columns,
            show='headings',
            selectmode='browse'
        )
        self.tree_clients.heading('ID', text='ID')
        self.tree_clients.heading('Nombre', text='Nombre')
        self.tree_clients.heading('Correo', text='Correo')
        self.tree_clients.column('ID', width=50, anchor='center')
        self.tree_clients.column('Nombre', width=200, anchor='w')
        self.tree_clients.column('Correo', width=250, anchor='w')

        # Estilo para filas alternas
        style = ttk.Style()
        style.configure(
            "Treeview",
            background="white",
            foreground="black",
            rowheight=25,
            fieldbackground="white"
        )
        style.map(
            'Treeview',
            background=[('selected', '#0c1834')],
            foreground=[('selected', 'white')]
        )
        self.tree_clients.tag_configure('oddrow', background='#f0f0f0')
        self.tree_clients.tag_configure('evenrow', background='white')

        # Scrollbars para la tabla
        scrollbar_vertical = ttk.Scrollbar(
            frame_clientes,
            orient=tk.VERTICAL,
            command=self.tree_clients.yview
        )
        scrollbar_vertical.pack(side=tk.RIGHT, fill='y')
        self.tree_clients.configure(yscrollcommand=scrollbar_vertical.set)

        scrollbar_horizontal = ttk.Scrollbar(
            frame_clientes,
            orient=tk.HORIZONTAL,
            command=self.tree_clients.xview
        )
        scrollbar_horizontal.pack(side=tk.BOTTOM, fill='x')
        self.tree_clients.configure(xscrollcommand=scrollbar_horizontal.set)

        self.tree_clients.pack(fill='both', expand=True)

        # ----------- Frame de botones para clientes -----------
        frame_botones_clientes = ttk.Frame(self.client_manager)
        frame_botones_clientes.grid(row=0, column=1, padx=10, pady=10, sticky='ns')

        # Botones para clientes
        boton_anadir_cliente = ttk.Button(
            frame_botones_clientes,
            text="Añadir Cliente",
            command=self.agregar_cliente,
            style='BlueZodiac.TButton',
            width=20
        )
        boton_anadir_cliente.pack(pady=10)

        boton_modificar_cliente = ttk.Button(
            frame_botones_clientes,
            text="Modificar Cliente",
            command=self.editar_cliente,
            style='Warning.TButton',
            width=20
        )
        boton_modificar_cliente.pack(pady=10)

        boton_borrar_cliente = ttk.Button(
            frame_botones_clientes,
            text="Borrar Cliente",
            command=self.eliminar_cliente,
            style='RedColor.TButton',
            width=20
        )
        boton_borrar_cliente.pack(pady=10)

        # ----------- Frame de la lista de aliases -----------
        frame_aliases = ttk.Frame(self.client_manager)
        frame_aliases.grid(row=0, column=2, padx=10, pady=10, sticky='nsew')

        # Etiqueta para aliases
        self.label_aliases = ttk.Label(
            frame_aliases,
            text="Aliases del Cliente Seleccionado:",
            style='Custom.TLabel'
        )
        self.label_aliases.pack(pady=5)

        # Lista para mostrar los aliases
        self.listbox_aliases = tk.Listbox(frame_aliases, height=15, width=30)
        self.listbox_aliases.pack(fill='both', expand=True)

        # Frame para los botones de aliases
        frame_botones_aliases = ttk.Frame(frame_aliases)
        frame_botones_aliases.pack(pady=10)

        boton_anadir_alias = ttk.Button(
            frame_botones_aliases,
            text="Añadir Alias",
            command=self.agregar_alias,
            style='BlueZodiac.TButton',
            width=20
        )
        boton_anadir_alias.pack(pady=5)

        boton_borrar_alias = ttk.Button(
            frame_botones_aliases,
            text="Borrar Alias",
            command=self.eliminar_alias,
            style='RedColor.TButton',
            width=20
        )
        boton_borrar_alias.pack(pady=5)

        # Cargar los clientes en la tabla
        self.load_clients()

        # Evento para cuando se selecciona un cliente
        self.tree_clients.bind('<<TreeviewSelect>>', self.on_client_select)

    def load_clients(self):
        # Limpiar la tabla de clientes
        for item in self.tree_clients.get_children():
            self.tree_clients.delete(item)

        try:
            clientes = self.db.get_all_clients()
            for count, (client_id, nombre, correo) in enumerate(clientes):
                if count % 2 == 0:
                    self.tree_clients.insert(
                        '',
                        tk.END,
                        iid=client_id,
                        values=(client_id, nombre, correo),
                        tags=('evenrow',)
                    )
                else:
                    self.tree_clients.insert(
                        '',
                        tk.END,
                        iid=client_id,
                        values=(client_id, nombre, correo),
                        tags=('oddrow',)
                    )
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error al cargar los clientes: {str(e)}")

    def on_client_select(self, event):
        selected_item = self.tree_clients.selection()
        if selected_item:
            item = selected_item[0]
            client_id = int(item)
            aliases = self.db.get_aliases_for_client(client_id)
            # Limpiar la lista de aliases
            self.listbox_aliases.delete(0, tk.END)
            # Agregar los aliases a la lista
            for alias_id, alias in aliases:
                self.listbox_aliases.insert(tk.END, f"{alias_id}:{alias}")

    def agregar_cliente(self):
        def guardar_cliente():
            nombre = entrada_nombre.get().strip()
            correo = entrada_correo.get().strip()
            if not nombre:
                messagebox.showwarning("Advertencia", "El nombre del cliente no puede estar vacío.")
                return
            try:
                self.db.add_client(nombre, correo)
                self.load_clients()
                ventana_agregar.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error al agregar el cliente: {str(e)}")

        ventana_agregar = ttk.Toplevel(self.client_manager)
        ventana_agregar.title("Agregar Cliente")
        ventana_agregar.geometry("400x250")
        ventana_agregar.configure(background='white')

        etiqueta_nombre = ttk.Label(ventana_agregar, text="Nombre del Cliente:", style='Custom.TLabel')
        etiqueta_nombre.pack(pady=5)
        entrada_nombre = ttk.Entry(ventana_agregar, width=50)
        entrada_nombre.pack()

        etiqueta_correo = ttk.Label(ventana_agregar, text="Correo Electrónico:", style='Custom.TLabel')
        etiqueta_correo.pack(pady=5)
        entrada_correo = ttk.Entry(ventana_agregar, width=50)
        entrada_correo.pack()

        boton_guardar = ttk.Button(ventana_agregar, text="Guardar", command=guardar_cliente,
                                   style='GreenSuccess.TButton')
        boton_guardar.pack(pady=20, ipadx=10, ipady=5)

        ventana_agregar.transient(self.client_manager)
        ventana_agregar.grab_set()
        self.client_manager.wait_window(ventana_agregar)

    def editar_cliente(self):
        selected_item = self.tree_clients.selection()
        if not selected_item:
            messagebox.showwarning("Advertencia", "Por favor, selecciona un cliente de la lista.")
            return

        item = selected_item[0]
        client_id = int(item)
        cliente = self.tree_clients.item(item, 'values')[0]
        correo_actual = self.tree_clients.item(item, 'values')[1]

        def guardar_cambio():
            nuevo_nombre = entrada_nombre.get().strip()
            nuevo_correo = entrada_correo.get().strip()
            if not nuevo_nombre:
                messagebox.showwarning("Advertencia", "El nombre del cliente no puede estar vacío.")
                return
            try:
                self.db.update_client_name_and_email(client_id, nuevo_nombre, nuevo_correo)
                self.load_clients()
                ventana_editar.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error al actualizar el cliente: {str(e)}")

        ventana_editar = ttk.Toplevel(self.client_manager)
        ventana_editar.title(f"Editar Cliente - {cliente}")
        ventana_editar.geometry("400x250")
        ventana_editar.configure(background='white')

        etiqueta_nombre = ttk.Label(ventana_editar, text="Nombre del Cliente:", style='Custom.TLabel')
        etiqueta_nombre.pack(pady=5)
        entrada_nombre = ttk.Entry(ventana_editar, width=50)
        entrada_nombre.insert(0, cliente)
        entrada_nombre.pack()

        etiqueta_correo = ttk.Label(ventana_editar, text="Correo Electrónico:", style='Custom.TLabel')
        etiqueta_correo.pack(pady=5)
        entrada_correo = ttk.Entry(ventana_editar, width=50)
        entrada_correo.insert(0, correo_actual)
        entrada_correo.pack()

        boton_guardar = ttk.Button(ventana_editar, text="Guardar", command=guardar_cambio,
                                   style='GreenSuccess.TButton')
        boton_guardar.pack(pady=20, ipadx=10, ipady=5)

        ventana_editar.transient(self.client_manager)
        ventana_editar.grab_set()
        self.client_manager.wait_window(ventana_editar)

    def eliminar_cliente(self):
        selected_item = self.tree_clients.selection()
        if not selected_item:
            messagebox.showwarning("Advertencia", "Por favor, selecciona un cliente de la lista.")
            return

        item = selected_item[0]
        client_id = int(item)
        cliente = self.tree_clients.item(item, 'values')[0]

        respuesta = messagebox.askyesno("Confirmar",
                                        f"¿Estás seguro de eliminar el cliente '{cliente}' y todos sus aliases?")
        if respuesta:
            try:
                self.db.delete_client(client_id)
                self.load_clients()
                self.listbox_aliases.delete(0, tk.END)
            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error al eliminar el cliente: {str(e)}")

    def agregar_alias(self):
        selected_item = self.tree_clients.selection()
        if not selected_item:
            messagebox.showwarning("Advertencia", "Por favor, selecciona un cliente de la lista.")
            return

        item = selected_item[0]
        client_id = int(item)

        def guardar_alias():
            nuevo_alias = entrada_alias.get().strip()
            if not nuevo_alias:
                messagebox.showwarning("Advertencia", "El alias no puede estar vacío.")
                return
            try:
                self.db.add_alias(client_id, nuevo_alias)
                self.on_client_select(None)
                ventana_alias.destroy()
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

    def eliminar_alias(self):
        selected_alias = self.listbox_aliases.curselection()
        if not selected_alias:
            messagebox.showwarning("Advertencia", "Por favor, selecciona un alias de la lista.")
            return

        alias_item = self.listbox_aliases.get(selected_alias[0])
        alias_id = int(alias_item.split(':')[0])

        respuesta = messagebox.askyesno("Confirmar",
                                        f"¿Estás seguro de eliminar el alias '{alias_item.split(':')[1]}'?")
        if respuesta:
            try:
                self.db.remove_alias(alias_id)
                self.on_client_select(None)
            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error al eliminar el alias: {str(e)}")

    def on_closing(self):
        # Cerrar la conexión a la base de datos al cerrar la aplicación
        if self.processor:
            self.processor.close()
        if self.db:
            self.db.close()
        self.destroy()

