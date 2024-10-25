import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
import logging
import os
import threading
import time

# Configuración del logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# URL de tu sitio de SharePoint
site_url = "https://chcconsultorescr.sharepoint.com/sites/CONSULTORAS"


def iniciar_sesion():
    username = entry_usuario.get()
    password = entry_contrasena.get()

    logging.debug(f"Intentando autenticación para el usuario: {username}")

    # Autenticación
    try:
        ctx_auth = AuthenticationContext(site_url)
        if ctx_auth.acquire_token_for_user(username, password):
            ctx = ClientContext(site_url, ctx_auth)
            logging.debug("Autenticación exitosa.")
            messagebox.showinfo("Éxito", "Autenticación exitosa.")
            ventana.destroy()  # Cerrar la ventana después de iniciar sesión
            # Llamar a la función principal
            acceder_sharepoint(ctx)
        else:
            logging.error("Error de autenticación. Por favor, verifica tus credenciales.")
            messagebox.showerror("Error", "Error de autenticación. Por favor, verifica tus credenciales.")
    except Exception as e:
        logging.error(f"Ocurrió un error: {e}")
        messagebox.showerror("Error", f"Ocurrió un error: {e}")


def acceder_sharepoint(ctx):
    # Crear la ventana del explorador
    ventana_explorador = tk.Tk()
    ventana_explorador.title("Explorador de SharePoint - Pruebas ASW")

    # Crear el Treeview
    tree = ttk.Treeview(ventana_explorador, columns=('type', 'url'))
    tree.heading('#0', text='Archivos y Carpetas', anchor='w')
    tree.column('type', width=0, stretch=False)
    tree.column('url', width=0, stretch=False)

    # Añadir un scrollbar
    scrollbar = ttk.Scrollbar(ventana_explorador, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)

    # Empaquetar el Treeview y el scrollbar
    tree.pack(side='left', fill='both', expand=True)
    scrollbar.pack(side='right', fill='y')

    # Especificar la ruta de la carpeta 'Pruebas ASW'
    folder_url = "/sites/CONSULTORAS/Shared Documents/Pruebas ASW"
    logging.debug(f"Analizando la carpeta: {folder_url}")

    # Obtener la carpeta raíz
    folder = ctx.web.get_folder_by_server_relative_url(folder_url)
    ctx.load(folder)
    ctx.execute_query()

    # Cargar los elementos en el Treeview
    cargar_treeview(ctx, folder, tree)

    # Añadir un botón de refrescar
    boton_refrescar = tk.Button(ventana_explorador, text="Refrescar",
                                command=lambda: refrescar_treeview(ctx, folder, tree))
    boton_refrescar.pack(side='bottom', fill='x')

    # Asociar evento de doble clic
    tree.bind("<Double-1>", lambda event: item_doble_click(event, ctx, tree))

    # Iniciar el ChangeDetector
    detector = ChangeDetector(ctx, folder_url, tree, interval=5)  # Intervalo de 60 segundos
    detector.start()

    # Detener el ChangeDetector al cerrar la ventana
    def on_closing():
        detector.stop()
        ventana_explorador.destroy()

    ventana_explorador.protocol("WM_DELETE_WINDOW", on_closing)
    ventana_explorador.mainloop()


def cargar_treeview(ctx, folder, tree, parent=""):
    try:
        ctx.load(folder, ["Name", "ServerRelativeUrl", "UniqueId"])
        ctx.execute_query()

        folder_id = folder.properties['UniqueId']
        folder_name = folder.properties['Name']

        # Insertar la carpeta en el Treeview
        node = tree.insert(parent, 'end', iid=folder_id, text=folder_name, open=False,
                           values=('folder', folder.properties['ServerRelativeUrl']))

        # Listar subcarpetas
        sub_folders = folder.folders
        ctx.load(sub_folders)
        ctx.execute_query()
        for sub_folder in sub_folders:
            cargar_treeview(ctx, sub_folder, tree, node)

        # Listar archivos en la carpeta actual
        files = folder.files
        ctx.load(files)
        ctx.execute_query()
        for file in files:
            file_id = file.properties['UniqueId']
            file_name = file.properties['Name']
            # Insertar el archivo en el Treeview
            tree.insert(node, 'end', iid=file_id, text=file_name, open=False,
                        values=('file', file.properties['ServerRelativeUrl']))
    except Exception as e:
        logging.error(f"Error cargando el Treeview: {e}")


def refrescar_treeview(ctx, folder, tree):
    # Limpiar el Treeview
    tree.delete(*tree.get_children())
    # Recargar el Treeview
    cargar_treeview(ctx, folder, tree)


def item_doble_click(event, ctx, tree):
    item_id = tree.focus()
    item = tree.item(item_id)
    item_type = tree.set(item_id, 'type')
    if item_type == 'file':
        file_url = tree.set(item_id, 'url')
        file_name = item['text']
        descargar_archivo(ctx, file_url, file_name)


def descargar_archivo(ctx, file_url, file_name, show_message=True):
    try:
        # Crear carpeta local si no existe
        os.makedirs('descargas', exist_ok=True)
        download_path = os.path.join('descargas', file_name)

        with open(download_path, "wb") as local_file:
            file = ctx.web.get_file_by_server_relative_url(file_url)
            file.download(local_file)
            ctx.execute_query()
            if show_message:
                messagebox.showinfo("Descarga",
                                    f"Archivo '{file_name}' descargado exitosamente en la carpeta 'descargas'.")
            else:
                logging.debug(f"Archivo '{file_name}' descargado automáticamente.")
    except Exception as e:
        logging.error(f"Error descargando el archivo '{file_name}': {e}")
        if show_message:
            messagebox.showerror("Error", f"Error descargando el archivo '{file_name}': {e}")


class ChangeDetector(threading.Thread):
    def __init__(self, ctx, folder_url, tree, interval=5):
        super().__init__()
        self.ctx = ctx
        self.folder_url = folder_url
        self.tree = tree
        self.interval = interval
        self.stop_event = threading.Event()
        self.previous_state = {}

    def run(self):
        while not self.stop_event.is_set():
            logging.debug("Verificando cambios en SharePoint...")
            self.check_for_changes()
            time.sleep(self.interval)

    def stop(self):
        self.stop_event.set()

    def check_for_changes(self):
        # Obtener el estado actual de los archivos y carpetas
        current_state = {}
        folder = self.ctx.web.get_folder_by_server_relative_url(self.folder_url)
        self.collect_items(self.ctx, folder, current_state)

        # Comparar con el estado anterior
        if self.previous_state:
            changes = self.detect_changes(self.previous_state, current_state)
            if changes:
                logging.debug(f"Se detectaron cambios: {changes}")
                # Actualizar la interfaz gráfica
                self.tree.delete(*self.tree.get_children())
                cargar_treeview(self.ctx, folder, self.tree)
                # Descargar archivos actualizados
                for file_url in changes.get('updated_files', []):
                    file_name = file_url.split('/')[-1]
                    descargar_archivo(self.ctx, file_url, file_name, show_message=False)
        else:
            # Primera ejecución, no hay estado anterior
            pass

        # Actualizar el estado anterior
        self.previous_state = current_state

    def collect_items(self, ctx, folder, state):
        try:
            ctx.load(folder, ["Name", "ServerRelativeUrl", "TimeLastModified"])
            ctx.execute_query()
            folder_url = folder.properties['ServerRelativeUrl']
            state[folder_url] = folder.properties['TimeLastModified']

            # Archivos
            files = folder.files
            ctx.load(files)
            ctx.execute_query()
            for file in files:
                file_url = file.properties['ServerRelativeUrl']
                state[file_url] = file.properties['TimeLastModified']

            # Subcarpetas
            sub_folders = folder.folders
            ctx.load(sub_folders)
            ctx.execute_query()
            for sub_folder in sub_folders:
                self.collect_items(ctx, sub_folder, state)
        except Exception as e:
            logging.error(f"Error al recopilar elementos: {e}")

    def detect_changes(self, old_state, new_state):
        changes = {'updated_files': []}
        for item_url, modified_time in new_state.items():
            if item_url not in old_state:
                # Nuevo elemento agregado
                changes['updated_files'].append(item_url)
            elif old_state[item_url] != modified_time:
                # Elemento modificado
                changes['updated_files'].append(item_url)
        return changes


# Crear la ventana principal
ventana = tk.Tk()
ventana.title("Inicio de Sesión SharePoint")

# Etiquetas y campos de entrada
label_usuario = tk.Label(ventana, text="Correo electrónico:")
label_usuario.grid(row=0, column=0, padx=10, pady=10)

entry_usuario = tk.Entry(ventana)
entry_usuario.grid(row=0, column=1, padx=10, pady=10)

label_contrasena = tk.Label(ventana, text="Contraseña:")
label_contrasena.grid(row=1, column=0, padx=10, pady=10)

entry_contrasena = tk.Entry(ventana, show="*")
entry_contrasena.grid(row=1, column=1, padx=10, pady=10)

# Botón de inicio de sesión
boton_iniciar = tk.Button(ventana, text="Iniciar Sesión", command=iniciar_sesion)
boton_iniciar.grid(row=2, column=0, columnspan=2, padx=10, pady=10)

ventana.mainloop()
