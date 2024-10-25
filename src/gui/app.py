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
        self.geometry("800x600")
        self.configure(bg="#f0f4f7")  # Fondo de color pastel
        self.file_path = ''
        self.processor = None
        self.db = ClientDatabase()
        self.config_manager = ConfigManager()
        self.create_widgets()
        self.check_credentials()  # Verificar si las credenciales están configuradas



    def create_widgets(self):
        # Crear la barra de menú
        menu_bar = tk.Menu(self)
        self.config(menu=menu_bar)

        # Menú "Configuración"
        config_menu = tk.Menu(menu_bar, tearoff=0)
        config_menu.add_command(label="Cambiar correo de origen", command=self.open_config_window)
        config_menu.add_command(label="Salir", command=self.on_closing)
        menu_bar.add_cascade(label="Configuración", menu=config_menu)

        # Menú "Gestión"
        gestion_menu = tk.Menu(menu_bar, tearoff=0)
        gestion_menu.add_command(label="Gestor de Clientes", command=self.open_client_manager)
        menu_bar.add_cascade(label="Gestión", menu=gestion_menu)

        # Frame para seleccionar archivo
        frame_archivo = ttk.Frame(self)
        frame_archivo.pack(pady=10)

        etiqueta_archivo = ttk.Label(frame_archivo, text="Selecciona el archivo XLSX:")
        etiqueta_archivo.pack(side=tk.LEFT, padx=5)

        self.entrada_archivo = ttk.Entry(frame_archivo, width=50)
        self.entrada_archivo.pack(side=tk.LEFT, padx=5)

        boton_seleccionar = ttk.Button(frame_archivo, text="Seleccionar archivo", command=self.seleccionar_archivo)
        boton_seleccionar.pack(side=tk.LEFT, padx=5)

        # Botones para filtrar/generar y enviar correos
        frame_botones = ttk.Frame(self)
        frame_botones.pack(pady=10)

        boton_generar = ttk.Button(frame_botones, text="Generar Reportes", command=self.filtrar_generar)
        boton_generar.pack(side=tk.LEFT, padx=5)

        boton_enviar = ttk.Button(frame_botones, text="Enviar Correos", command=self.enviar_correos)
        boton_enviar.pack(side=tk.LEFT, padx=5)

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

    def filtrar_generar(self):
        if not self.file_path:
            messagebox.showwarning("Advertencia", "Por favor, selecciona un archivo.")
            return

        try:
            self.processor = FileProcessor(self.file_path)
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
        self.client_manager = tk.Toplevel(self)
        self.client_manager.title("Gestor de Clientes")
        self.client_manager.geometry("800x500")
        self.client_manager.configure(bg="#f0f4f7")

        # Frame para la lista de clientes
        frame_clientes = ttk.Frame(self.client_manager)
        frame_clientes.pack(side=tk.LEFT, fill='both', expand=True, padx=10, pady=10)

        # Tabla para mostrar los clientes
        self.tree_clients = ttk.Treeview(frame_clientes, columns=('Cliente', 'Correo'), show='headings', selectmode='browse')
        self.tree_clients.heading('Cliente', text='Cliente')
        self.tree_clients.heading('Correo', text='Correo')
        self.tree_clients.column('Cliente', width=200)
        self.tree_clients.column('Correo', width=250)
        self.tree_clients.pack(fill='both', expand=True)

        # Scrollbar para la tabla de clientes
        scrollbar_clients = ttk.Scrollbar(frame_clientes, orient=tk.VERTICAL, command=self.tree_clients.yview)
        self.tree_clients.configure(yscroll=scrollbar_clients.set)
        scrollbar_clients.pack(side=tk.RIGHT, fill='y')

        # Frame para los botones de clientes
        frame_botones_clientes = ttk.Frame(self.client_manager)
        frame_botones_clientes.pack(pady=5)

        boton_agregar_cliente = ttk.Button(frame_botones_clientes, text="Agregar Cliente", command=self.agregar_cliente)
        boton_agregar_cliente.pack(side=tk.LEFT, padx=5)

        boton_editar_cliente = ttk.Button(frame_botones_clientes, text="Editar Cliente", command=self.editar_cliente)
        boton_editar_cliente.pack(side=tk.LEFT, padx=5)

        boton_eliminar_cliente = ttk.Button(frame_botones_clientes, text="Eliminar Cliente", command=self.eliminar_cliente)
        boton_eliminar_cliente.pack(side=tk.LEFT, padx=5)

        # Frame para la lista de aliases
        frame_aliases = ttk.Frame(self.client_manager)
        frame_aliases.pack(side=tk.RIGHT, fill='both', expand=True, padx=10, pady=10)

        # Etiqueta para aliases
        self.label_aliases = ttk.Label(frame_aliases, text="Aliases del Cliente Seleccionado:")
        self.label_aliases.pack()

        # Lista para mostrar los aliases
        self.listbox_aliases = tk.Listbox(frame_aliases)
        self.listbox_aliases.pack(fill='both', expand=True)

        # Frame para los botones de aliases
        frame_botones_aliases = ttk.Frame(self.client_manager)
        frame_botones_aliases.pack(pady=5)

        boton_agregar_alias = ttk.Button(frame_botones_aliases, text="Agregar Alias", command=self.agregar_alias)
        boton_agregar_alias.pack(side=tk.LEFT, padx=5)

        boton_eliminar_alias = ttk.Button(frame_botones_aliases, text="Eliminar Alias", command=self.eliminar_alias)
        boton_eliminar_alias.pack(side=tk.LEFT, padx=5)

        # Cargar los clientes
        self.load_clients()

        # Evento para cuando se selecciona un cliente
        self.tree_clients.bind('<<TreeviewSelect>>', self.on_client_select)

    def load_clients(self):
        # Limpiar la tabla de clientes
        for item in self.tree_clients.get_children():
            self.tree_clients.delete(item)

        try:
            clientes = self.db.get_all_clients()
            for client_id, nombre, correo in clientes:
                self.tree_clients.insert('', tk.END, iid=client_id, values=(nombre, correo))
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

        ventana_agregar = tk.Toplevel(self.client_manager)
        ventana_agregar.title("Agregar Cliente")
        ventana_agregar.geometry("400x200")
        ventana_agregar.configure(bg="#f0f4f7")

        etiqueta_nombre = ttk.Label(ventana_agregar, text="Nombre del Cliente:")
        etiqueta_nombre.pack(pady=5)
        entrada_nombre = ttk.Entry(ventana_agregar, width=50)
        entrada_nombre.pack()

        etiqueta_correo = ttk.Label(ventana_agregar, text="Correo Electrónico:")
        etiqueta_correo.pack(pady=5)
        entrada_correo = ttk.Entry(ventana_agregar, width=50)
        entrada_correo.pack()

        boton_guardar = ttk.Button(ventana_agregar, text="Guardar", command=guardar_cliente)
        boton_guardar.pack(pady=10)

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

        ventana_editar = tk.Toplevel(self.client_manager)
        ventana_editar.title(f"Editar Cliente - {cliente}")
        ventana_editar.geometry("400x200")
        ventana_editar.configure(bg="#f0f4f7")

        etiqueta_nombre = ttk.Label(ventana_editar, text="Nombre del Cliente:")
        etiqueta_nombre.pack(pady=5)
        entrada_nombre = ttk.Entry(ventana_editar, width=50)
        entrada_nombre.insert(0, cliente)
        entrada_nombre.pack()

        etiqueta_correo = ttk.Label(ventana_editar, text="Correo Electrónico:")
        etiqueta_correo.pack(pady=5)
        entrada_correo = ttk.Entry(ventana_editar, width=50)
        entrada_correo.insert(0, correo_actual)
        entrada_correo.pack()

        boton_guardar = ttk.Button(ventana_editar, text="Guardar", command=guardar_cambio)
        boton_guardar.pack(pady=10)

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

        respuesta = messagebox.askyesno("Confirmar", f"¿Estás seguro de eliminar el cliente '{cliente}' y todos sus aliases?")
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

        ventana_alias = tk.Toplevel(self.client_manager)
        ventana_alias.title("Agregar Alias")
        ventana_alias.geometry("400x150")
        ventana_alias.configure(bg="#f0f4f7")

        etiqueta_alias = ttk.Label(ventana_alias, text="Nuevo Alias:")
        etiqueta_alias.pack(pady=5)
        entrada_alias = ttk.Entry(ventana_alias, width=50)
        entrada_alias.pack()

        boton_guardar = ttk.Button(ventana_alias, text="Guardar", command=guardar_alias)
        boton_guardar.pack(pady=10)

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

        respuesta = messagebox.askyesno("Confirmar", f"¿Estás seguro de eliminar el alias '{alias_item.split(':')[1]}'?")
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
