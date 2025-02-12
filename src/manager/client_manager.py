# src/gui/app.py
import json
import sqlite3
import threading
import ttkbootstrap as ttk
import tkinter as tk
from tkinter import messagebox, filedialog
from datetime import datetime

def on_tab_change(self, event):
    notebook = event.widget
    selected_tab = notebook.tab('current')['text']

    if selected_tab == "Por Lista":
        selected_list = self.selected_lista.get()
        if selected_list:
            self.on_list_select_por_lista(None)
    elif selected_tab == "Todos los Clientes":
        self.load_all_clients()

def on_list_select_por_lista(self, event):
    selected_list_name = self.selected_lista.get()
    if not selected_list_name:
        return

    list_id = self.list_ids.get(selected_list_name)
    if not list_id:
        return

    # Limpiar los clientes existentes en el Treeview
    for item in self.tree_clients_por_lista.get_children():
        self.tree_clients_por_lista.delete(item)

    # Obtener los clientes de la lista seleccionada (la tupla ahora tiene 5 elementos)
    clients = self.db.get_clients_by_list(list_id)
    for client in clients:
        # Insertar todos los campos: (id, nombre, correo, correo_cc, correo_bcc)
        self.tree_clients_por_lista.insert('', 'end', values=client)

    # Limpiar el Listbox de aliases
    self.listbox_aliases_por_lista.delete(0, tk.END)

def load_all_clients(self):
    # Limpiar los clientes existentes
    for item in self.tree_all_clients.get_children():
        self.tree_all_clients.delete(item)

    # Obtener todos los clientes
    # Se espera que self.db.get_all_clients() retorne tuplas con:
    # (id, nombre, correo, correo_cc, correo_bcc)
    clients = self.db.get_all_clients()

    for client in clients:
        client_id, client_name, client_email, client_cc, client_bcc = client
        # Verificar si existe reporte para este cliente
        tag = 'has_report' if self.processor and self.processor.get_client_file(client_name) else 'no_report'
        # Insertar todos los 5 campos en el Treeview
        self.tree_all_clients.insert('', 'end',
                                     values=(client_id, client_name, client_email, client_cc, client_bcc),
                                     tags=(tag,))
    # Limpiar la lista de aliases
    self.listbox_aliases_all_clients.delete(0, tk.END)

def add_client_to_existing_list(self):
    # Determinar la pestaña activa
    notebook = self.client_manager.nametowidget("!notebook")
    current_tab = notebook.tab(notebook.select(), "text")

    if current_tab == "Por Lista":
        tree = self.tree_clients_por_lista
    elif current_tab == "Todos los Clientes":
        tree = self.tree_all_clients
    else:
        messagebox.showwarning("Advertencia", "No se pudo determinar la pestaña activa.")
        return

    selected_items = tree.selection()
    if not selected_items:
        messagebox.showwarning("Advertencia", "Por favor, selecciona al menos un cliente de la lista.")
        return

    # Obtener las listas a las que los clientes seleccionados NO están agregados
    try:
        all_lists = self.db.get_all_lists()  # Retorna una lista de tuplas (id, nombre)
        available_lists = {nombre: id for (id, nombre) in all_lists}

        # Verificar para cada cliente seleccionado
        for item in selected_items:
            client_id = int(tree.item(item, 'values')[0])
            client_lists = self.db.get_client_lists(client_id)  # Listas en las que está el cliente
            for (id, nombre) in all_lists:
                if id in client_lists:
                    available_lists.pop(nombre, None)

        if not available_lists:
            messagebox.showinfo("Información",
                                "Los clientes seleccionados ya están en todas las listas existentes.")
            return
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al obtener las listas: {str(e)}")
        return

    def guardar_asignacion():
        selected_list = combobox_lists.get()
        if not selected_list:
            messagebox.showwarning("Advertencia", "Por favor, selecciona una lista.")
            return
        list_id = self.list_ids.get(selected_list)
        if not list_id:
            messagebox.showerror("Error", "La lista seleccionada no es válida.")
            return

        try:
            # Asignar cada cliente seleccionado a la lista
            for item in selected_items:
                client_id = int(tree.item(item, 'values')[0])
                self.db.add_client_to_list(client_id, list_id)
            messagebox.showinfo("Éxito", f"Clientes agregados a la lista '{selected_list}' exitosamente.")
            ventana_asignar.destroy()

            # Recargar las tablas para reflejar el cambio
            if current_tab == "Por Lista":
                self.load_clients_in_selected_list()
            elif current_tab == "Todos los Clientes":
                self.load_all_clients()
        except sqlite3.IntegrityError:
            messagebox.showerror("Error", f"Uno o más clientes ya están en la lista '{selected_list}'.")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error al agregar los clientes a la lista: {str(e)}")

    # Crear la ventana de asignación a lista
    ventana_asignar = ttk.Toplevel(self.client_manager)
    ventana_asignar.title("Añadir Clientes a una Lista")

    # Ajustar el tamaño inicial de la ventana según el escalado
    initial_width = int(600 * self.scaling_factor)
    initial_height = int(250 * self.scaling_factor)
    self.center_window(ventana_asignar, initial_width, initial_height)

    # Configuración del grid
    ventana_asignar.columnconfigure(0, weight=1)
    ventana_asignar.rowconfigure(0, weight=1)
    ventana_asignar.rowconfigure(1, weight=1)
    ventana_asignar.rowconfigure(2, weight=1)

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
        width=int(20 * self.scaling_factor),
    )
    combobox_lists.grid(row=0, column=1, pady=int(5 * self.scaling_factor), padx=int(5 * self.scaling_factor),
                        sticky='ew')
    combobox_lists.set(list(available_lists.keys())[0] if available_lists else '')

    # Botón para guardar la asignación
    boton_guardar = ttk.Button(
        frame_principal,
        text="Guardar",
        command=guardar_asignacion,
        style='GreenSuccess.TButton'
    )
    boton_guardar.grid(row=1, column=0, columnspan=2, pady=int(20 * self.scaling_factor),
                       ipadx=int(10 * self.scaling_factor), ipady=int(5 * self.scaling_factor), sticky='ew')

    # Configurar la ventana como modal
    ventana_asignar.transient(self.client_manager)
    ventana_asignar.grab_set()
    self.client_manager.wait_window(ventana_asignar)

def on_list_select(self, event):
    """Evento llamado al seleccionar una lista. Muestra y carga los detalles."""
    selected_list_name = self.selected_lista.get()
    if selected_list_name:
        self.frame_details.grid()  # Mostrar el frame_details
        self.load_clients_in_selected_list()
    else:
        self.frame_details.grid_remove()  # Ocultar el frame_details si no hay selección

def load_clients_in_selected_list(self):
    """Carga los clientes asociados a la lista seleccionada en la tabla (Por Lista)."""
    # Limpiar la tabla de clientes por lista
    for item in self.tree_clients_por_lista.get_children():
        self.tree_clients_por_lista.delete(item)

    try:
        selected_list_name = self.selected_lista.get()
        if selected_list_name and selected_list_name in self.list_ids:
            list_id = self.list_ids[selected_list_name]
            # Se espera que get_clients_by_list retorne tuplas de 5 elementos:
            # (id, nombre, correo, correo_cc, correo_bcc)
            clientes = self.db.get_clients_by_list(list_id)
        else:
            clientes = []

        for count, client in enumerate(clientes):
            # client es una tupla con 5 valores.
            tag = 'evenrow' if count % 2 == 0 else 'oddrow'
            self.tree_clients_por_lista.insert('', tk.END, iid=client[0],
                                                values=client, tags=(tag,))
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
        correo_cc = entrada_cc.get().strip()
        correo_bcc = entrada_bcc.get().strip()
        if not nombre:
            messagebox.showwarning("Advertencia", "El nombre del cliente no puede estar vacío.")
            return
        try:
            # Se pasan también correo_cc y correo_bcc
            self.db.add_client(nombre, correo, correo_cc, correo_bcc)
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
        except sqlite3.IntegrityError:
            messagebox.showerror("Error", f"El cliente '{nombre}' ya existe.")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error al agregar el cliente: {str(e)}")

    # Crear la ventana de agregar cliente
    ventana_agregar = ttk.Toplevel(self.client_manager)
    ventana_agregar.title("Agregar Cliente")

    # Ajustar el tamaño inicial de la ventana según el escalado
    initial_width = int(650 * self.scaling_factor)  # Ajustado para mayor legibilidad
    initial_height = int(350 * self.scaling_factor)  # Más espacio para los campos

    self.center_window(ventana_agregar, initial_width, initial_height)

    # Configurar el grid para una mejor disposición de los widgets
    ventana_agregar.columnconfigure(0, weight=1)
    for i in range(4):
        ventana_agregar.rowconfigure(i, weight=1)

    # Definir fuente para los labels y entradas
    font_labels = ("Helvetica", 14, "bold")  # Etiquetas en negrita
    font_entries = ("Helvetica", 14)  # Entradas de texto

    # Frame principal dentro de la ventana
    frame_principal = ttk.Frame(ventana_agregar, padding=int(10 * self.scaling_factor))
    frame_principal.grid(row=0, column=0, sticky='nsew')
    frame_principal.columnconfigure(1, weight=1)

    # Etiqueta y entrada para el nombre del cliente
    etiqueta_nombre = ttk.Label(frame_principal, text="Nombre del Cliente:", font=font_labels)
    etiqueta_nombre.grid(row=0, column=0, pady=int(5 * self.scaling_factor), sticky='w')
    entrada_nombre = ttk.Entry(frame_principal, font=font_entries, width=int(30 * self.scaling_factor))
    entrada_nombre.grid(row=0, column=1, pady=int(5 * self.scaling_factor), padx=int(5 * self.scaling_factor), sticky='ew')

    # Etiqueta y entrada para el correo electrónico
    etiqueta_correo = ttk.Label(frame_principal, text="Correo Electrónico:", font=font_labels)
    etiqueta_correo.grid(row=1, column=0, pady=int(5 * self.scaling_factor), sticky='w')
    entrada_correo = ttk.Entry(frame_principal, font=font_entries, width=int(30 * self.scaling_factor))
    entrada_correo.grid(row=1, column=1, pady=int(5 * self.scaling_factor), padx=int(5 * self.scaling_factor), sticky='ew')

    # Etiqueta y entrada para Correo CC
    etiqueta_cc = ttk.Label(frame_principal, text="Correo CC:", font=font_labels)
    etiqueta_cc.grid(row=2, column=0, pady=int(5 * self.scaling_factor), sticky='w')
    entrada_cc = ttk.Entry(frame_principal, font=font_entries, width=int(30 * self.scaling_factor))
    entrada_cc.grid(row=2, column=1, pady=int(5 * self.scaling_factor), padx=int(5 * self.scaling_factor), sticky='ew')

    # Etiqueta y entrada para Correo BCC (CCO)
    etiqueta_bcc = ttk.Label(frame_principal, text="Correo BCC:", font=font_labels)
    etiqueta_bcc.grid(row=3, column=0, pady=int(5 * self.scaling_factor), sticky='w')
    entrada_bcc = ttk.Entry(frame_principal, font=font_entries, width=int(30 * self.scaling_factor))
    entrada_bcc.grid(row=3, column=1, pady=int(5 * self.scaling_factor), padx=int(5 * self.scaling_factor), sticky='ew')

    # Botón para guardar el cliente
    boton_guardar = ttk.Button(frame_principal, text="Guardar", command=guardar_cliente, style='GreenSuccess.TButton')
    boton_guardar.grid(row=4, column=0, columnspan=2, pady=int(20 * self.scaling_factor),
                       ipadx=int(10 * self.scaling_factor), ipady=int(5 * self.scaling_factor))

    # Configurar la ventana como modal
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

    # Se obtiene el ID del cliente desde el Treeview y se consulta la información completa en la base de datos
    item = selected_item[0]
    client_id = int(tree.item(item, 'values')[0])
    client_info = self.db.get_client_by_id(client_id)
    if not client_info:
        messagebox.showerror("Error", "No se encontró información del cliente.")
        return
    # client_info es (id, nombre, correo, correo_cc, correo_bcc)
    cliente_nombre = client_info[1]
    correo_actual = client_info[2]
    correo_cc_actual = client_info[3]
    correo_bcc_actual = client_info[4]

    def guardar_cambio():
        nuevo_nombre = entrada_nombre.get().strip()
        nuevo_correo = entrada_correo.get().strip()
        nuevo_cc = entrada_cc.get().strip()
        nuevo_bcc = entrada_bcc.get().strip()

        if not nuevo_nombre:
            messagebox.showwarning("Advertencia", "El nombre del cliente no puede estar vacío.")
            return

        try:
            self.db.update_client_name_and_email(client_id, nuevo_nombre, nuevo_correo, nuevo_cc, nuevo_bcc)

            # Recargar las tablas después de modificar el cliente
            if current_tab == "Por Lista":
                self.load_clients_in_selected_list()
            elif current_tab == "Todos los Clientes":
                self.load_all_clients()

            ventana_editar.destroy()
        except sqlite3.IntegrityError:
            messagebox.showerror("Error", f"El cliente '{nuevo_nombre}' ya existe.")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error al actualizar el cliente: {str(e)}")

    # Ventana de edición de cliente
    ventana_editar = ttk.Toplevel(self.client_manager)
    ventana_editar.title(f"Editar Cliente - {cliente_nombre}")

    # Ajustar el tamaño inicial de la ventana según el escalado
    initial_width = int(650 * self.scaling_factor)
    initial_height = int(350 * self.scaling_factor)  # Se aumenta la altura para incluir CC y BCC

    self.center_window(ventana_editar, initial_width, initial_height)

    # Configuración del grid de la ventana (5 filas: para nombre, correo, CC, BCC y el botón)
    for i in range(5):
        ventana_editar.rowconfigure(i, weight=1)
    ventana_editar.columnconfigure(0, weight=1)

    # Definir fuente para los labels y entradas
    font_labels = ("Helvetica", 14, "bold")  # Etiquetas en negrita
    font_entries = ("Helvetica", 14)  # Entradas de texto

    # Frame principal dentro de la ventana de edición
    frame_principal = ttk.Frame(ventana_editar, padding=int(10 * self.scaling_factor))
    frame_principal.grid(row=0, column=0, sticky='nsew')
    frame_principal.columnconfigure(1, weight=1)

    # Etiqueta y entrada para el nombre del cliente
    etiqueta_nombre = ttk.Label(frame_principal, text="Nombre del Cliente:", font=font_labels)
    etiqueta_nombre.grid(row=0, column=0, pady=int(5 * self.scaling_factor), sticky='w')
    entrada_nombre = ttk.Entry(frame_principal, font=font_entries, width=int(30 * self.scaling_factor))
    entrada_nombre.insert(0, cliente_nombre)
    entrada_nombre.grid(row=0, column=1, pady=int(5 * self.scaling_factor), padx=int(5 * self.scaling_factor), sticky='ew')

    # Etiqueta y entrada para el correo electrónico
    etiqueta_correo = ttk.Label(frame_principal, text="Correo Electrónico:", font=font_labels)
    etiqueta_correo.grid(row=1, column=0, pady=int(5 * self.scaling_factor), sticky='w')
    entrada_correo = ttk.Entry(frame_principal, font=font_entries, width=int(30 * self.scaling_factor))
    entrada_correo.insert(0, correo_actual)
    entrada_correo.grid(row=1, column=1, pady=int(5 * self.scaling_factor), padx=int(5 * self.scaling_factor), sticky='ew')

    # Etiqueta y entrada para Correo CC
    etiqueta_cc = ttk.Label(frame_principal, text="Correo CC:", font=font_labels)
    etiqueta_cc.grid(row=2, column=0, pady=int(5 * self.scaling_factor), sticky='w')
    entrada_cc = ttk.Entry(frame_principal, font=font_entries, width=int(30 * self.scaling_factor))
    entrada_cc.insert(0, correo_cc_actual)
    entrada_cc.grid(row=2, column=1, pady=int(5 * self.scaling_factor), padx=int(5 * self.scaling_factor), sticky='ew')

    # Etiqueta y entrada para Correo BCC
    etiqueta_bcc = ttk.Label(frame_principal, text="Correo BCC:", font=font_labels)
    etiqueta_bcc.grid(row=3, column=0, pady=int(5 * self.scaling_factor), sticky='w')
    entrada_bcc = ttk.Entry(frame_principal, font=font_entries, width=int(30 * self.scaling_factor))
    entrada_bcc.insert(0, correo_bcc_actual)
    entrada_bcc.grid(row=3, column=1, pady=int(5 * self.scaling_factor), padx=int(5 * self.scaling_factor), sticky='ew')

    # Botón para guardar los cambios
    boton_guardar = ttk.Button(frame_principal, text="Guardar", command=guardar_cambio, style='GreenSuccess.TButton')
    boton_guardar.grid(row=4, column=0, columnspan=2, pady=int(15 * self.scaling_factor),
                       ipadx=int(10 * self.scaling_factor), ipady=int(5 * self.scaling_factor), sticky='ew')

    # Configurar la ventana como modal
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

    selected_items = tree.selection()
    if not selected_items:
        messagebox.showwarning("Advertencia", "Por favor, selecciona un cliente de la lista.")
        return

    # Confirmación de eliminación
    if len(selected_items) == 1:
        cliente_nombre = tree.item(selected_items[0], 'values')[1]
        mensaje = f"¿Estás seguro de eliminar el cliente '{cliente_nombre}' y todos sus aliases?"
    else:
        mensaje = f"¿Estás seguro de eliminar los {len(selected_items)} clientes seleccionados y todos sus aliases?"

    respuesta = messagebox.askyesno("Confirmar Eliminación", mensaje)
    if respuesta:
        try:
            # Almacenar la información de los clientes eliminados para deshacer
            deleted_clients = []
            for item in selected_items:
                client_id = int(tree.item(item, 'values')[0])
                client_name = tree.item(item, 'values')[1]
                client_email = tree.item(item, 'values')[2]
                aliases = self.db.get_aliases_for_client(client_id)
                deleted_clients.append({
                    'id': client_id,
                    'nombre': client_name,
                    'correo': client_email,
                    'aliases': aliases
                })
                self.db.delete_client(client_id)

            # Agregar la acción de deshacer a la pila
            self.undo_stack.append({
                'action': 'delete_clients',
                'data': deleted_clients,
                'list_id': self.list_ids.get(self.selected_lista.get(), None) if current_tab == "Por Lista" else None
            })

            # Recargar las tablas después de eliminar los clientes
            if current_tab == "Por Lista":
                self.load_clients_in_selected_list()
            elif current_tab == "Todos los Clientes":
                self.load_all_clients()
            # Limpiar la lista de aliases
            listbox.delete(0, tk.END)
            messagebox.showinfo("Éxito", "Cliente(s) eliminado(s) exitosamente.")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error al eliminar el cliente(s): {str(e)}")

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
            # messagebox.showinfo("Éxito", "Alias agregado exitosamente.")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error al agregar el alias: {str(e)}")

    # Crear la ventana de agregar alias
    ventana_alias = ttk.Toplevel(self.client_manager)
    ventana_alias.title("Agregar Alias")

    # Ajustar el tamaño inicial de la ventana según el escalado
    initial_width = int(600 * self.scaling_factor)
    initial_height = int(250 * self.scaling_factor)

    self.center_window(ventana_alias, initial_width, initial_height)

    # Configurar el grid para una mejor disposición de los widgets
    ventana_alias.columnconfigure(0, weight=1)
    ventana_alias.rowconfigure(0, weight=1)
    ventana_alias.rowconfigure(1, weight=1)
    ventana_alias.rowconfigure(2, weight=1)

    # Frame principal dentro de la ventana de alias
    frame_principal = ttk.Frame(ventana_alias, padding=int(10 * self.scaling_factor))
    frame_principal.grid(row=0, column=0, sticky='nsew')
    frame_principal.columnconfigure(1, weight=1)

    # Etiqueta y entrada para el nuevo alias
    etiqueta_alias = ttk.Label(
        frame_principal,
        text="Nuevo Alias:",
        style='Custom.TLabel',
        # , int(12 * self.scaling_factor))
    )
    etiqueta_alias.grid(row=0, column=0, pady=int(5 * self.scaling_factor), sticky='w')

    entrada_alias = ttk.Entry(
        frame_principal,
        width=int(30 * self.scaling_factor),
        # , int(10 * self.scaling_factor))
    )
    entrada_alias.grid(row=0, column=1, pady=int(5 * self.scaling_factor), padx=int(5 * self.scaling_factor),
                       sticky='ew')

    # Botón para guardar el alias
    boton_guardar = ttk.Button(
        frame_principal,
        text="Guardar",
        command=guardar_alias,
        style='GreenSuccess.TButton'
    )
    boton_guardar.grid(row=1, column=0, columnspan=2, pady=int(20 * self.scaling_factor),
                       ipadx=int(10 * self.scaling_factor), ipady=int(5 * self.scaling_factor))

    # Configurar la ventana como modal
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
            # messagebox.showinfo("Éxito", "Alias eliminado exitosamente.")
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
            self.update_list_combobox_ui()  # Actualizar el Combobox
            ventana_nueva_lista.destroy()
            # messagebox.showinfo("Éxito", f"Lista '{nombre_lista}' creada exitosamente.")
        except sqlite3.IntegrityError:
            messagebox.showerror("Error", f"La lista '{nombre_lista}' ya existe.")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error al crear la lista: {str(e)}")

    # Crear la ventana de agregar lista
    ventana_nueva_lista = ttk.Toplevel(self.client_manager)
    ventana_nueva_lista.title("Crear Nueva Lista")

    ventana_nueva_lista.configure(background='white')
    # Ajustar el tamaño inicial de la ventana según el escalado
    initial_width = int(400 * self.scaling_factor)
    initial_height = int(150 * self.scaling_factor)

    self.center_window(ventana_nueva_lista, initial_width, initial_height)

    # Configurar el grid para una mejor disposición de los widgets
    ventana_nueva_lista.columnconfigure(0, weight=1)
    ventana_nueva_lista.rowconfigure(0, weight=1)
    ventana_nueva_lista.rowconfigure(1, weight=1)
    ventana_nueva_lista.rowconfigure(2, weight=1)

    # Frame principal dentro de la ventana de nueva lista
    frame_principal = ttk.Frame(ventana_nueva_lista, padding=int(10 * self.scaling_factor))
    frame_principal.grid(row=0, column=0, sticky='nsew')
    frame_principal.columnconfigure(1, weight=1)

    # Etiqueta y entrada para el nombre de la lista
    etiqueta_nombre = ttk.Label(
        frame_principal,
        text="Nombre de la Lista:",
        style='Custom.TLabel',
        # , int(12 * self.scaling_factor))
    )
    etiqueta_nombre.grid(row=0, column=0, pady=int(5 * self.scaling_factor), sticky='w')

    entrada_nombre = ttk.Entry(
        frame_principal,
        width=int(30 * self.scaling_factor),
        # , int(10 * self.scaling_factor))
    )
    entrada_nombre.grid(row=0, column=1, pady=int(5 * self.scaling_factor), padx=int(5 * self.scaling_factor),
                        sticky='ew')

    # Botón para guardar la nueva lista
    boton_guardar = ttk.Button(
        frame_principal,
        text="Guardar",
        command=guardar_lista,
        style='GreenSuccess.TButton'
    )
    boton_guardar.grid(row=1, column=0, columnspan=2, pady=int(10 * self.scaling_factor),
                       ipadx=int(10 * self.scaling_factor), ipady=int(5 * self.scaling_factor), sticky='ew')

    # Configurar la ventana como modal
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
            self.update_list_combobox_ui()  # Actualizar el Combobox
            # messagebox.showinfo("Éxito", f"Lista '{selected_lista}' eliminada exitosamente.")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error al eliminar la lista: {str(e)}")

def update_list_combobox_ui(self):
    """
    Actualiza los valores del Combobox de listas en la interfaz.
    """
    try:
        self.listas = self.db.get_all_lists()  # Lista de tuplas (id, nombre)
        self.list_names = [nombre for (id, nombre) in self.listas]
        self.list_ids = {nombre: id for (id, nombre) in self.listas}
        self.combobox_listas['values'] = self.list_names
        if self.list_names:
            self.combobox_listas.set(self.list_names[0])
        else:
            self.combobox_listas.set('')
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al actualizar el Combobox de listas: {str(e)}")

def add_client_to_selected_list(self):
    selected_list_name = self.selected_lista.get()
    if not selected_list_name:
        messagebox.showwarning("Advertencia", "Por favor, selecciona una lista.")
        return

    selected_items = self.tree_all_clients.selection()
    if not selected_items:
        messagebox.showwarning("Advertencia", "Por favor, selecciona al menos un cliente para agregar.")
        return

    list_id = self.list_ids.get(selected_list_name)
    try:
        for item in selected_items:
            client_id = int(self.tree_all_clients.item(item, 'values')[0])
            self.db.add_client_to_list(client_id, list_id)
        self.load_clients()
        self.load_clients_in_selected_list()
        messagebox.showinfo("Éxito", "Clientes agregados a la lista exitosamente.")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al agregar los clientes a la lista: {str(e)}")

def remove_client_from_selected_list(self):
    selected_list_name = self.selected_lista.get()
    if not selected_list_name:
        messagebox.showwarning("Advertencia", "Por favor, selecciona una lista.")
        return

    selected_items = self.tree_clients_por_lista.selection()
    if not selected_items:
        messagebox.showwarning("Advertencia", "Por favor, selecciona al menos un cliente de la lista.")
        return

    list_id = self.list_ids.get(selected_list_name)
    respuesta = messagebox.askyesno("Confirmar",
                                    f"¿Estás seguro de eliminar los clientes seleccionados de la lista '{selected_list_name}'?")
    if respuesta:
        try:
            for item in selected_items:
                client_id = int(self.tree_clients_por_lista.item(item, 'values')[0])
                client_name = self.tree_clients_por_lista.item(item, 'values')[1]
                self.db.remove_client_from_list(client_id, list_id)
            self.on_list_select_por_lista(None)  # Actualizar la tabla de clientes
            messagebox.showinfo("Éxito", "Clientes eliminados de la lista exitosamente.")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error al eliminar los clientes de la lista: {str(e)}")

def export_clients(self):
    """
    Exporta todos los clientes y sus asociaciones a listas a un archivo JSON.
    """
    import tkinter.filedialog as fd
    from tkinter import messagebox
    import json

    # Selecciona la ruta de exportación
    file_path = fd.asksaveasfilename(defaultextension=".json", filetypes=[("JSON Files", "*.json")])
    if not file_path:
        return  # Si no se selecciona un archivo, se cancela la exportación

    # Prepara la estructura de datos
    export_data = {"listas": []}

    # Obtén todas las listas con sus clientes asignados
    listas = self.db.get_all_lists()  # Se espera una lista de tuplas (id, nombre)
    for list_id, list_name in listas:
        # Se espera que get_clients_by_list retorne tuplas con:
        # (id, nombre, correo, correo_cc, correo_bcc)
        clientes = self.db.get_clients_by_list(list_id)
        lista_data = {
            "id": list_id,
            "nombre": list_name,
            "clientes": []
        }
        for client in clientes:
            client_id, client_name, client_email, client_cc, client_bcc = client
            aliases = self.db.get_aliases_for_client(client_id)
            lista_data["clientes"].append({
                "id": client_id,
                "nombre": client_name,
                "correo": client_email,
                "correo_cc": client_cc,
                "correo_bcc": client_bcc,
                "aliases": [alias for _, alias in aliases]
            })
        export_data["listas"].append(lista_data)

    # Obtén los clientes sin lista
    all_clients = self.db.get_all_clients()  # Se espera que retorne tuplas (id, nombre, correo, correo_cc, correo_bcc)
    # Extraer los IDs de clientes que ya están asignados a alguna lista
    clients_with_lists = set()
    for lista in export_data["listas"]:
        for client in lista["clientes"]:
            clients_with_lists.add(client["id"])

    unassigned_clients = [client for client in all_clients if client[0] not in clients_with_lists]

    if unassigned_clients:
        lista_sin_asignar = {
            "id": None,
            "nombre": "Sin asignar",
            "clientes": []
        }
        for client in unassigned_clients:
            client_id, client_name, client_email, client_cc, client_bcc = client
            aliases = self.db.get_aliases_for_client(client_id)
            lista_sin_asignar["clientes"].append({
                "id": client_id,
                "nombre": client_name,
                "correo": client_email,
                "correo_cc": client_cc,
                "correo_bcc": client_bcc,
                "aliases": [alias for _, alias in aliases]
            })
        export_data["listas"].append(lista_sin_asignar)

    # Guarda el JSON
    try:
        with open(file_path, "w", encoding="utf-8") as file:
            json.dump(export_data, file, indent=4, ensure_ascii=False)
        messagebox.showinfo("Exportación completa", f"Los datos se han exportado a {file_path}")
    except Exception as e:
        messagebox.showerror("Error de exportación", f"No se pudo exportar: {e}")

def import_clients(self):
    # Abrir el diálogo de selección de archivo
    file_path = filedialog.askopenfilename(
        title="Seleccione el archivo JSON de clientes",
        filetypes=[("JSON Files", "*.json")]
    )

    if not file_path:
        print("No se seleccionó ningún archivo.")
        return

    # Cargar los datos desde el archivo JSON
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            data = json.load(file)
    except Exception as e:
        print(f"Error al leer el archivo JSON: {e}")
        return

    # Procesar cada lista y cliente en el JSON, generando nuevos IDs
    for lista in data.get("listas", []):
        list_name = lista.get("nombre")

        # Agregar la lista si no existe
        if not self.db.list_exists_by_name(list_name):
            self.db.add_list(list_name)
        list_id = self.db.get_list_id_by_name(list_name)

        for cliente in lista.get("clientes", []):
            client_name = cliente.get("nombre")
            client_email = cliente.get("correo")
            # Recuperar los nuevos campos; si no existen, se asigna cadena vacía
            client_cc = cliente.get("correo_cc", "")
            client_bcc = cliente.get("correo_bcc", "")
            aliases = cliente.get("aliases", [])

            # Agregar el cliente si no existe
            if not self.db.client_exists_by_name(client_name):
                self.db.add_client(client_name, client_email, client_cc, client_bcc)
            client_id = self.db.get_client_id_by_name(client_name)

            # Agregar alias al cliente si no existen
            for alias in aliases:
                if not self.db.alias_exists(client_id, alias):
                    self.db.add_alias(client_id, alias)
                    print(f"Alias '{alias}' agregado para cliente '{client_name}'.")

            # Asociar el cliente a la lista
            try:
                self.db.add_client_to_list(client_id, list_id)
            except sqlite3.IntegrityError:
                print(f"La asociación del cliente '{client_name}' y la lista '{list_name}' ya existe.")
            except ValueError as e:
                print(e)

    # Mostrar mensaje de confirmación
    messagebox.showinfo("Importación completada", "Los datos se importaron correctamente.")

    # Eliminar la lista 'Sin asignar' si existe
    if self.db.list_exists_by_name("Sin asignar"):
        list_id = self.db.get_list_id_by_name("Sin asignar")
        self.db.delete_list(list_id)
        print("La lista 'Sin asignar' fue eliminada después de la importación.")

def open_client_manager(self):
    # Obtener la resolución de pantalla y calcular el factor de escala
    screen_width = self.winfo_screenwidth()
    screen_height = self.winfo_screenheight()
    self.scaling_factor = min(screen_width / 1920, screen_height / 1080)  # Basado en 1920x1080
    window_width = min(int(screen_width * 1), 1200)
    window_height = min(int(screen_height * 0.8), 1400)  # Ajustado para evitar ventanas demasiado altas

    # Calcular la posición x e y para centrar la ventana
    x = int((screen_width / 2) - (window_width / 2))
    y = int((screen_height / 2) - (window_height / 2))

    # Crear la ventana del gestor de clientes
    self.client_manager = tk.Toplevel(self)
    self.client_manager.title("Gestor de Clientes")
    self.client_manager.geometry(f"{window_width}x{window_height}+{x}+{y}")
    self.client_manager.configure(background='white')

    self.center_window(self.client_manager, window_width, window_height)

    # Configuración del grid principal
    self.client_manager.columnconfigure(0, weight=1)
    self.client_manager.rowconfigure(0, weight=1)

    # Crear el Notebook
    notebook = ttk.Notebook(self.client_manager)
    notebook.grid(row=0, column=0, sticky='nsew')

    # Pestaña 1: Por Lista
    tab_por_lista = ttk.Frame(notebook, padding=int(10 * self.scaling_factor))
    #notebook.add(tab_por_lista, text="Por Lista")
    tab_por_lista.columnconfigure(0, weight=1)
    tab_por_lista.rowconfigure(0, weight=1)

    # Frame para la selección de listas
    frame_selection = ttk.Frame(tab_por_lista, padding=int(10 * self.scaling_factor))
    frame_selection.grid(row=0, column=0, padx=int(10 * self.scaling_factor),
                         pady=int(10 * self.scaling_factor), sticky='ew')
    frame_selection.columnconfigure(1, weight=1)

    # Etiqueta para seleccionar lista
    label_select_lista = ttk.Label(
        frame_selection,
        text="Seleccionar Lista:",
        style='Custom.TLabel'
    )
    label_select_lista.grid(row=0, column=0, pady=int(5 * self.scaling_factor), sticky='w')

    # Obtener todas las listas
    self.listas = self.db.get_all_lists()
    self.list_names = [nombre for (id, nombre) in self.listas]
    self.list_ids = {nombre: id for (id, nombre) in self.listas}

    # Variable para la lista seleccionada
    self.selected_lista = tk.StringVar()
    self.selected_lista.set(self.list_names[0] if self.list_names else '')

    # Definir el combobox en la inicialización del gestor de clientes
    self.combobox_listas = ttk.Combobox(
        frame_selection,
        textvariable=self.selected_lista,
        values=self.list_names,
        state='readonly',
        width=int(20 * self.scaling_factor)
    )
    self.combobox_listas.grid(row=0, column=1, padx=int(5 * self.scaling_factor),
                              pady=int(5 * self.scaling_factor), sticky='ew')
    self.combobox_listas.bind("<<ComboboxSelected>>", self.on_list_select_por_lista)

    # Ahora, llamar a update_list_combobox_ui para asegurarse de que el combobox está actualizado
    self.update_list_combobox_ui()

    # Botones para crear y eliminar listas
    boton_crear_lista = ttk.Button(
        frame_selection,
        text="Crear Lista",
        command=self.create_new_list,
        style='BlueZodiac.TButton'
    )
    boton_crear_lista.grid(row=1, column=0, padx=int(5 * self.scaling_factor),
                           pady=int(5 * self.scaling_factor), sticky='w')

    boton_eliminar_lista = ttk.Button(
        frame_selection,
        text="Eliminar Lista",
        command=self.delete_selected_list,
        style='RedColor.TButton'
    )
    boton_eliminar_lista.grid(row=1, column=1, padx=int(5 * self.scaling_factor),
                              pady=int(5 * self.scaling_factor), sticky='w')

    # Crear el menú contextual para los Treeview
    self.context_menu = tk.Menu(self.client_manager, tearoff=0)
    self.context_menu.add_command(label="Agregar a Lista...", command=self.context_menu_add_to_list)
    self.context_menu.add_command(label="Agregar a Nueva Lista...", command=self.context_menu_add_to_new_list)
    self.context_menu.add_separator()  # Agregar una línea separadora para mejor organización
    self.context_menu.add_command(label="Eliminar Cliente(s)", command=self.delete_client_from_context_menu)

    # Frame para la tabla de clientes por lista (Pestaña "Por Lista")
    frame_clientes = ttk.Frame(tab_por_lista, padding=int(10 * self.scaling_factor))
    frame_clientes.grid(row=1, column=0, padx=int(10 * self.scaling_factor),
                        pady=int(10 * self.scaling_factor), sticky='nsew')
    frame_clientes.columnconfigure(0, weight=1)
    frame_clientes.rowconfigure(0, weight=1)

    # Definir las columnas incluyendo "Correo CC" y "Correo BCC"
    columns_por_lista = ('ID', 'Nombre', 'Correo', 'Correo CC', 'Correo BCC')
    self.tree_clients_por_lista = ttk.Treeview(
        frame_clientes,
        columns=columns_por_lista,
        show='headings',
        selectmode='extended',  # Permite selección múltiple
        height=10
    )

    # Configurar etiquetas (tags) para filas (opcional)
    self.tree_clients_por_lista.tag_configure('has_report', background='#ccffcc')
    self.tree_clients_por_lista.tag_configure('no_report', background='#ffcccc')

    # Configurar los encabezados y las columnas
    self.tree_clients_por_lista.heading('ID', text='ID')
    self.tree_clients_por_lista.heading('Nombre', text='Nombre')
    self.tree_clients_por_lista.heading('Correo', text='Correo')
    self.tree_clients_por_lista.heading('Correo CC', text='Correo CC')
    self.tree_clients_por_lista.heading('Correo BCC', text='Correo BCC')

    self.tree_clients_por_lista.column('ID', width=int(50 * self.scaling_factor), anchor='center')
    self.tree_clients_por_lista.column('Nombre', width=int(200 * self.scaling_factor), anchor='w')
    self.tree_clients_por_lista.column('Correo', width=int(250 * self.scaling_factor), anchor='w')
    self.tree_clients_por_lista.column('Correo CC', width=int(200 * self.scaling_factor), anchor='w')
    self.tree_clients_por_lista.column('Correo BCC', width=int(200 * self.scaling_factor), anchor='w')

    self.tree_clients_por_lista.grid(row=0, column=0, sticky='nsew')

    # Vincular eventos para la selección y el clic derecho
    self.tree_clients_por_lista.bind('<<TreeviewSelect>>', self.on_client_select)
    self.tree_clients_por_lista.bind("<Button-3>", self.show_context_menu)

    # Scrollbars para el Treeview
    scrollbar_vertical = ttk.Scrollbar(frame_clientes, orient=tk.VERTICAL, command=self.tree_clients_por_lista.yview)
    scrollbar_vertical.grid(row=0, column=1, sticky='ns')
    self.tree_clients_por_lista.configure(yscrollcommand=scrollbar_vertical.set)

    scrollbar_horizontal = ttk.Scrollbar(frame_clientes, orient=tk.HORIZONTAL,
                                         command=self.tree_clients_por_lista.xview)
    scrollbar_horizontal.grid(row=1, column=0, sticky='ew')
    self.tree_clients_por_lista.configure(xscrollcommand=scrollbar_horizontal.set)

    # Panel de botones de gestión de clientes por lista
    frame_botones_clientes = ttk.Frame(frame_clientes, padding=int(10 * self.scaling_factor))
    frame_botones_clientes.grid(row=2, column=0, padx=int(10 * self.scaling_factor),
                                pady=int(5 * self.scaling_factor), sticky='ew')
    frame_botones_clientes.columnconfigure((0, 1, 2, 3), weight=1)

    # Vincular Ctrl + Z a la acción de deshacer
    self.client_manager.bind_all('<Control-z>', self.undo_action)

    boton_anadir_cliente = ttk.Button(
        frame_botones_clientes,
        text="Añadir Cliente",
        command=self.add_client,
        style='BlueZodiac.TButton'
    )
    boton_anadir_cliente.grid(row=0, column=0, padx=int(5 * self.scaling_factor),
                              pady=int(5 * self.scaling_factor), sticky='ew')

    boton_modificar_cliente = ttk.Button(
        frame_botones_clientes,
        text="Modificar Cliente",
        command=self.edit_client,
        style='Warning.TButton'
    )
    boton_modificar_cliente.grid(row=0, column=1, padx=int(5 * self.scaling_factor),
                                 pady=int(5 * self.scaling_factor), sticky='ew')


    # Nuevo botón para agregar a una lista existente
    boton_anadir_a_lista_all = ttk.Button(
        frame_botones_clientes,
        text="Añadir a Lista",
        command=self.add_client_to_existing_list,
        style='BlueZodiac.TButton'
    )
    boton_anadir_a_lista_all.grid(row=0, column=2, padx=int(5 * self.scaling_factor),
                                  pady=int(5 * self.scaling_factor), sticky='ew')

    # Nuevo botón para eliminar de una lista
    boton_eliminar_de_lista = ttk.Button(
        frame_botones_clientes,
        text="Eliminar de Lista",
        command=self.remove_client_from_selected_list,
        style='RedColor.TButton'
    )
    boton_eliminar_de_lista.grid(row=0, column=3, padx=int(5 * self.scaling_factor),
                                 pady=int(5 * self.scaling_factor), sticky='ew')

    # Frame para gestión de aliases en "Por Lista"
    frame_aliases = ttk.Frame(tab_por_lista, padding=int(10 * self.scaling_factor))
    frame_aliases.grid(row=2, column=0, padx=int(10 * self.scaling_factor),
                       pady=int(10 * self.scaling_factor), sticky='nsew')
    frame_aliases.columnconfigure(0, weight=1)
    frame_aliases.rowconfigure(0, weight=1)

    label_aliases = ttk.Label(
        frame_aliases,
        text="Consignatario(s):",
        style='Custom.TLabel'
    )
    label_aliases.grid(row=0, column=0, pady=int(5 * self.scaling_factor), sticky='w')

    # Frame que contiene la listbox y los botones de aliases
    frame_aliases_content = ttk.Frame(frame_aliases)
    frame_aliases_content.grid(row=1, column=0, sticky='nsew')
    frame_aliases_content.columnconfigure(0, weight=1)
    frame_aliases_content.rowconfigure(0, weight=1)

    # Lista para mostrar los aliases del cliente seleccionado en "Por Lista"
    self.listbox_aliases_por_lista = tk.Listbox(
        frame_aliases_content,
        height=12  # Establecer la altura en 12 filas visibles
    )
    self.listbox_aliases_por_lista.grid(row=0, column=0, sticky='nsew')
    self.listbox_aliases_por_lista.bind("<Double-1>", self.edit_alias_inline)

    # Scrollbar para la lista de aliases por lista
    scrollbar_aliases = ttk.Scrollbar(
        frame_aliases_content,
        orient=tk.VERTICAL,
        command=self.listbox_aliases_por_lista.yview
    )
    scrollbar_aliases.grid(row=0, column=1, sticky='ns')
    self.listbox_aliases_por_lista.configure(yscrollcommand=scrollbar_aliases.set)

    # Panel de botones para gestión de aliases por lista (al costado derecho)
    frame_botones_aliases = ttk.Frame(frame_aliases_content)
    frame_botones_aliases.grid(row=0, column=2, padx=int(5 * self.scaling_factor), sticky='ns')
    frame_botones_aliases.columnconfigure(0, weight=1)

    boton_anadir_alias = ttk.Button(
        frame_botones_aliases,
        text="Añadir Alias",
        command=self.add_alias,
        style='BlueZodiac.TButton'
    )
    boton_anadir_alias.grid(row=0, column=0, padx=0, pady=int(5 * self.scaling_factor), sticky='ew')

    boton_borrar_alias = ttk.Button(
        frame_botones_aliases,
        text="Borrar Alias",
        command=self.delete_alias,
        style='RedColor.TButton'
    )
    boton_borrar_alias.grid(row=1, column=0, padx=0, pady=int(5 * self.scaling_factor), sticky='ew')

    # Pestaña 2: "Todos los Clientes"
    tab_todos_clientes = ttk.Frame(notebook, padding=int(10 * self.scaling_factor))
    notebook.add(tab_todos_clientes, text="Todos los Clientes")
    notebook.add(tab_por_lista, text="Por Lista")

    tab_todos_clientes.columnconfigure(0, weight=1)
    tab_todos_clientes.rowconfigure(0, weight=1)

    # Frame para la lista de todos los clientes en "Todos los Clientes"
    frame_all_clients = ttk.Frame(tab_todos_clientes, padding=int(10 * self.scaling_factor))
    frame_all_clients.grid(row=0, column=0, padx=int(10 * self.scaling_factor),
                           pady=int(10 * self.scaling_factor), sticky='nsew')
    frame_all_clients.columnconfigure(0, weight=1)
    frame_all_clients.rowconfigure(0, weight=1)

    # Actualizamos la definición de columnas para mostrar los nuevos campos CC y BCC
    columns_all = ('ID', 'Nombre', 'Correo', 'Correo CC', 'Correo BCC')
    self.tree_all_clients = ttk.Treeview(
        frame_all_clients,
        columns=columns_all,
        show='headings',
        selectmode='extended',
        height=10  # Número de filas visibles
    )

    # Configurar los encabezados y las columnas
    self.tree_all_clients.heading('ID', text='ID')
    self.tree_all_clients.heading('Nombre', text='Nombre')
    self.tree_all_clients.heading('Correo', text='Correo')
    self.tree_all_clients.heading('Correo CC', text='Correo CC')
    self.tree_all_clients.heading('Correo BCC', text='Correo BCC')
    self.tree_all_clients.column('ID', width=int(50 * self.scaling_factor), anchor='center')
    self.tree_all_clients.column('Nombre', width=int(200 * self.scaling_factor), anchor='w')
    self.tree_all_clients.column('Correo', width=int(250 * self.scaling_factor), anchor='w')
    self.tree_all_clients.column('Correo CC', width=int(200 * self.scaling_factor), anchor='w')
    self.tree_all_clients.column('Correo BCC', width=int(200 * self.scaling_factor), anchor='w')

    self.tree_all_clients.grid(row=0, column=0, sticky='nsew')

    self.tree_all_clients.tag_configure('has_report', background='#ccffcc')
    self.tree_all_clients.tag_configure('no_report', background='#ffcccc')

    # Vincular eventos para la selección y clic derecho en "Todos los Clientes"
    self.tree_all_clients.bind('<<TreeviewSelect>>', self.on_all_client_select)
    self.tree_all_clients.bind("<Button-3>", self.show_context_menu)

    # Scrollbars para la tabla de "Todos los Clientes"
    scrollbar_vertical_all = ttk.Scrollbar(frame_all_clients, orient=tk.VERTICAL, command=self.tree_all_clients.yview)
    scrollbar_vertical_all.grid(row=0, column=1, sticky='ns')
    self.tree_all_clients.configure(yscrollcommand=scrollbar_vertical_all.set)

    scrollbar_horizontal_all = ttk.Scrollbar(frame_all_clients, orient=tk.HORIZONTAL,
                                             command=self.tree_all_clients.xview)
    scrollbar_horizontal_all.grid(row=1, column=0, sticky='ew')
    self.tree_all_clients.configure(xscrollcommand=scrollbar_horizontal_all.set)

    # Panel de botones para gestión de todos los clientes (se mantiene similar, sin cambios en este punto)
    frame_botones_all_clients = ttk.Frame(frame_all_clients, padding=int(10 * self.scaling_factor))
    frame_botones_all_clients.grid(row=2, column=0, padx=int(10 * self.scaling_factor),
                                   pady=int(5 * self.scaling_factor), sticky='ew')
    frame_botones_all_clients.columnconfigure((0, 1, 2, 3, 4, 5), weight=1)

    boton_anadir_cliente_all = ttk.Button(
        frame_botones_all_clients,
        text="Añadir Cliente",
        command=self.add_client,
        style='BlueZodiac.TButton'
    )
    boton_anadir_cliente_all.grid(row=0, column=0, padx=int(5 * self.scaling_factor),
                                  pady=int(5 * self.scaling_factor), sticky='ew')

    boton_modificar_cliente_all = ttk.Button(
        frame_botones_all_clients,
        text="Modificar Cliente",
        command=self.edit_client,
        style='Warning.TButton'
    )
    boton_modificar_cliente_all.grid(row=0, column=1, padx=int(5 * self.scaling_factor),
                                     pady=int(5 * self.scaling_factor), sticky='ew')

    boton_borrar_cliente_all = ttk.Button(
        frame_botones_all_clients,
        text="Borrar Cliente",
        command=self.delete_client,
        style='RedColor.TButton'
    )
    boton_borrar_cliente_all.grid(row=0, column=2, padx=int(5 * self.scaling_factor),
                                  pady=int(5 * self.scaling_factor), sticky='ew')

    boton_anadir_a_lista_all = ttk.Button(
        frame_botones_all_clients,
        text="Añadir a Lista",
        command=self.add_client_to_existing_list,
        style='BlueZodiac.TButton'
    )
    boton_anadir_a_lista_all.grid(row=0, column=3, padx=int(5 * self.scaling_factor),
                                  pady=int(5 * self.scaling_factor), sticky='ew')

    boton_enviar_reportes_all = ttk.Button(
        frame_botones_all_clients,
        text="Enviar Reportes",
        command=self.send_reports_to_selected_clients,
        style='EmeraldGreen.TButton'
    )
    boton_enviar_reportes_all.grid(row=0, column=4, padx=int(5 * self.scaling_factor),
                                   pady=int(5 * self.scaling_factor), sticky='ew')

    boton_seleccionar_con_reporte = ttk.Button(
        frame_botones_all_clients,
        text="Seleccionar con Reporte",
        command=self.select_all_with_reports_in_all_clients,
        style='Warning.TButton'
    )
    boton_seleccionar_con_reporte.grid(row=0, column=5, padx=int(5 * self.scaling_factor),
                                       pady=int(5 * self.scaling_factor), sticky='ew')

    # Frame para gestión de aliases en "Todos los Clientes"
    frame_aliases_all = ttk.Frame(tab_todos_clientes, padding=int(10 * self.scaling_factor))
    frame_aliases_all.grid(row=1, column=0, padx=int(10 * self.scaling_factor),
                           pady=int(10 * self.scaling_factor), sticky='nsew')
    frame_aliases_all.columnconfigure(0, weight=1)
    frame_aliases_all.rowconfigure(0, weight=1)

    label_aliases_all = ttk.Label(
        frame_aliases_all,
        text="Consignatario(s):",
        style='Custom.TLabel'
    )
    label_aliases_all.grid(row=0, column=0, pady=int(5 * self.scaling_factor), sticky='w')

    # Frame que contiene la listbox y los botones de aliases
    frame_aliases_content_all = ttk.Frame(frame_aliases_all)
    frame_aliases_content_all.grid(row=1, column=0, sticky='nsew')
    frame_aliases_content_all.columnconfigure(0, weight=1)
    frame_aliases_content_all.rowconfigure(0, weight=1)

    # Lista para mostrar los aliases del cliente seleccionado en "Todos los Clientes"
    self.listbox_aliases_all_clients = tk.Listbox(
        frame_aliases_content_all,
        height=15  # Establecer la altura en 15 filas visibles
    )
    self.listbox_aliases_all_clients.grid(row=0, column=0, sticky='nsew')
    self.listbox_aliases_all_clients.bind("<Double-1>", self.edit_alias_inline)

    # Scrollbar para la lista de aliases de todos los clientes
    scrollbar_aliases_all_clients = ttk.Scrollbar(
        frame_aliases_content_all,
        orient=tk.VERTICAL,
        command=self.listbox_aliases_all_clients.yview
    )
    scrollbar_aliases_all_clients.grid(row=0, column=1, sticky='ns')
    self.listbox_aliases_all_clients.configure(yscrollcommand=scrollbar_aliases_all_clients.set)

    # Panel de botones para gestión de aliases de todos los clientes (al costado derecho)
    frame_botones_aliases_all_clients = ttk.Frame(frame_aliases_content_all)
    frame_botones_aliases_all_clients.grid(row=0, column=2, padx=int(5 * self.scaling_factor), sticky='ns')
    frame_botones_aliases_all_clients.columnconfigure(0, weight=1)

    boton_anadir_alias_all = ttk.Button(
        frame_botones_aliases_all_clients,
        text="Añadir Alias",
        command=self.add_alias,
        style='BlueZodiac.TButton'
    )
    boton_anadir_alias_all.grid(row=0, column=0, padx=0, pady=int(5 * self.scaling_factor), sticky='ew')

    boton_borrar_alias_all = ttk.Button(
        frame_botones_aliases_all_clients,
        text="Borrar Alias",
        command=self.delete_alias,
        style='RedColor.TButton'
    )
    boton_borrar_alias_all.grid(row=1, column=0, padx=0, pady=int(5 * self.scaling_factor), sticky='ew')

    # Cargar todos los clientes
    self.load_all_clients()

    # Llamar a la función para cargar la lista predeterminada automáticamente
    if self.list_names:
        self.combobox_listas.current(0)  # Seleccionar la primera lista
        self.on_list_select_por_lista(None)  # Cargar los clientes de la lista predeterminada

    # Bind para el cambio de pestaña
    notebook.bind('<<NotebookTabChanged>>', self.on_tab_change)

    # Configurar la ventana como modal
    self.client_manager.transient(self)
    self.client_manager.grab_set()
    self.wait_window(self.client_manager)

def context_menu_add_to_list(self):
    """
    Maneja la acción de agregar los clientes seleccionados a una lista existente.
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

    # Obtener todas las listas
    try:
        all_lists = self.db.get_all_lists()  # Retorna una lista de tuplas (id, nombre)
        available_lists = {nombre: id for (id, nombre) in all_lists}

        # Filtrar las listas en las que todos los clientes ya están presentes
        for list_id, list_name in all_lists:
            # Verificar si **todos** los clientes seleccionados ya están en esta lista
            all_present = True
            for item in selected_items:
                client_id = int(tree.item(item, 'values')[0])
                if not self.db.is_client_in_list(client_id, list_id):
                    all_present = False
                    break
            if all_present:
                # Si todos los clientes ya están en la lista, eliminarla de available_lists
                available_lists.pop(list_name, None)

        if not available_lists:
            messagebox.showinfo("Información",
                                "Los clientes seleccionados ya están en todas las listas existentes.")
            return
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al obtener las listas: {str(e)}")
        return

    # Mostrar la ventana para seleccionar la lista
    self.show_add_to_list_dialog(tree, selected_items, available_lists)

def show_context_menu(self, event):
    """
    Muestra el menú contextual al hacer clic derecho en un Treeview.
    Preserva la selección múltiple si el elemento clicado ya está seleccionado.
    """
    # Obtener el widget Treeview
    widget = event.widget
    # Identificar la región y el ítem bajo el cursor
    region = widget.identify('region', event.x, event.y)
    if region == 'cell':
        iid = widget.identify_row(event.y)
        if iid:
            # Verificar si el ítem ya está seleccionado
            if iid not in widget.selection():
                # Si no está seleccionado, seleccionar solo este ítem
                widget.selection_set(iid)
            # Si ya está seleccionado, mantener la selección existente
        else:
            # Si se hace clic fuera de cualquier ítem, limpiar la selección
            widget.selection_clear()
        # Crear el menú contextual
        context_menu = tk.Menu(self.client_manager, tearoff=0)
        context_menu.add_command(label="Enviar Reporte(s)", command=self.send_reports_to_selected_clients)
        context_menu.add_separator()  # Agregar una línea separadora para mejor organización
        context_menu.add_command(label="Agregar a Lista...", command=self.context_menu_add_to_list)
        context_menu.add_command(label="Agregar a Nueva Lista...", command=self.context_menu_add_to_new_list)
        context_menu.add_separator()  # Agregar una línea separadora para mejor organización
        context_menu.add_command(label="Eliminar Cliente(s)", command=self.delete_client_from_context_menu)

        # -- AQUI VIENE LO NUEVO --
        # Verificar si al menos uno de los items seleccionados tiene el tag 'has_report'.
        selected_items = widget.selection()
        has_report_any = False
        for sel in selected_items:
            tags = widget.item(sel, 'tags')
            if 'has_report' in tags:
                has_report_any = True
                break

        if has_report_any:
            context_menu.add_separator()
            context_menu.add_command(label="Abrir Reporte en Excel", command=self.open_reports_in_excel)
        # -- FIN DE LO NUEVO --

        # Mostrar el menú
        context_menu.tk_popup(event.x_root, event.y_root)
    # Evitar que el menú contextual permanezca en la pantalla
    context_menu.grab_release()

def delete_client_from_context_menu(self):
    """
    Elimina los clientes seleccionados desde el menú contextual.
    """
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

    selected_items = tree.selection()
    if not selected_items:
        messagebox.showwarning("Advertencia", "Por favor, selecciona al menos un cliente de la lista.")
        return

    # Confirmación de eliminación
    if len(selected_items) == 1:
        cliente_nombre = tree.item(selected_items[0], 'values')[1]
        mensaje = f"¿Estás seguro de eliminar el cliente '{cliente_nombre}' y todos sus aliases?"
    else:
        mensaje = f"¿Estás seguro de eliminar los {len(selected_items)} clientes seleccionados y todos sus aliases?"

    respuesta = messagebox.askyesno("Confirmar Eliminación", mensaje)
    if respuesta:
        try:
            # Almacenar la información de los clientes eliminados para deshacer
            deleted_clients = []
            for item in selected_items:
                client_id = int(tree.item(item, 'values')[0])
                client_name = tree.item(item, 'values')[1]
                client_email = tree.item(item, 'values')[2]
                aliases = self.db.get_aliases_for_client(client_id)
                deleted_clients.append({
                    'id': client_id,
                    'nombre': client_name,
                    'correo': client_email,
                    'aliases': aliases
                })
                self.db.delete_client(client_id)

            # Agregar la acción de deshacer a la pila
            self.undo_stack.append({
                'action': 'delete_clients',
                'data': deleted_clients,
                'list_id': self.list_ids.get(self.selected_lista.get(), None) if current_tab == "Por Lista" else None
            })

            # Recargar las tablas después de eliminar los clientes
            if current_tab == "Por Lista":
                self.load_clients_in_selected_list()
            elif current_tab == "Todos los Clientes":
                self.load_all_clients()
            # Limpiar la lista de aliases
            listbox.delete(0, tk.END)
            messagebox.showinfo("Éxito", "Cliente(s) eliminado(s) exitosamente.")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error al eliminar el cliente(s): {str(e)}")

def edit_alias_inline(self, event):
    """
    Permite editar un alias directamente en el Listbox haciendo doble clic.
    """
    # Obtener el widget Listbox que disparó el evento
    widget = event.widget
    # Obtener la posición del clic
    x, y = event.x, event.y
    # Obtener el índice del alias bajo el cursor
    index = widget.nearest(y)
    if index >= 0:
        # Obtener el alias seleccionado
        alias_item = widget.get(index)
        try:
            alias_id, alias = alias_item.split(':', 1)
        except ValueError:
            messagebox.showerror("Error", "Formato de alias inválido.")
            return

        # Obtener la posición del alias en el Listbox
        bbox = widget.bbox(index)
        if not bbox:
            return
        entry_x, entry_y, entry_width, entry_height = bbox
        # Crear un Entry widget sobre el alias para editar
        self.entry_alias = ttk.Entry(widget, width=30)
        self.entry_alias.insert(0, alias)
        self.entry_alias.select_range(0, tk.END)
        self.entry_alias.focus_set()
        # Posicionar el Entry sobre el alias
        self.entry_alias.place(x=entry_x, y=entry_y, width=entry_width, height=entry_height)

        def guardar_alias(event=None):
            nuevo_alias = self.entry_alias.get().strip()
            if not nuevo_alias:
                messagebox.showwarning("Advertencia", "El alias no puede estar vacío.")
                return
            try:
                # Actualizar en la base de datos
                self.db.update_alias(int(alias_id), nuevo_alias)
                # Actualizar en el Listbox
                widget.delete(index)
                widget.insert(index, f"{alias_id}:{nuevo_alias}")
                self.entry_alias.destroy()
                messagebox.showinfo("Éxito", "Alias actualizado exitosamente.")
            except sqlite3.IntegrityError:
                messagebox.showerror("Error", f"El alias '{nuevo_alias}' ya existe para este cliente.")
            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error al actualizar el alias: {str(e)}")

        def cancelar_edicion(event=None):
            self.entry_alias.destroy()

        # Bindings para guardar o cancelar la edición
        self.entry_alias.bind("<Return>", guardar_alias)
        self.entry_alias.bind("<FocusOut>", guardar_alias)
        self.entry_alias.bind("<Escape>", cancelar_edicion)

def undo_action(self, event=None):
    """
    Deshace la última acción realizada.
    """
    if not self.undo_stack:
        messagebox.showinfo("Deshacer", "No hay acciones para deshacer.")
        return

    last_action = self.undo_stack.pop()
    action_type = last_action.get('action')

    if action_type == 'delete_clients':
        deleted_clients = last_action.get('data', [])
        list_id = last_action.get('list_id', None)

        try:
            self.db.connection.execute('BEGIN')  # Iniciar transacción

            for client in deleted_clients:
                correo = client['correo']
                if isinstance(correo, tuple):
                    correo = correo[0]  # Ajusta el índice según la estructura del tuple

                # Reinsertar el cliente con su ID original
                self.db.add_client(
                    nombre=client['nombre'],
                    correo=correo,
                    client_id=client['id']
                )

                for alias in client['aliases']:
                    alias_str = alias[1]
                    if isinstance(alias_str, tuple):
                        alias_str = alias_str[0]
                    self.db.add_alias(client_id=client['id'], alias=alias_str)

                if list_id:
                    self.db.add_client_to_list(client_id=client['id'], list_id=list_id)

            self.db.connection.commit()  # Confirmar transacción

            # Recargar las tablas para reflejar los cambios
            self.load_all_clients()
            if self.selected_lista.get():
                self.load_clients_in_selected_list()

            messagebox.showinfo("Deshacer", "La eliminación de cliente(s) ha sido deshecha exitosamente.")
        except Exception as e:
            self.db.connection.rollback()  # Revertir transacción
            messagebox.showerror("Error", f"Ocurrió un error al deshacer la eliminación: {str(e)}")
    else:
        messagebox.showwarning("Deshacer", "La acción no se puede deshacer.")

def send_reports_to_selected_clients(self):
    """
    Envía reportes por correo electrónico a los clientes seleccionados.
    """
    # Determinar la pestaña activa
    notebook = self.client_manager.nametowidget("!notebook")
    current_tab = notebook.tab(notebook.select(), "text")

    if current_tab == "Por Lista":
        tree = self.tree_clients_por_lista
    elif current_tab == "Todos los Clientes":
        tree = self.tree_all_clients
    else:
        messagebox.showwarning("Advertencia", "No se pudo determinar la pestaña activa.")
        return

    selected_items = tree.selection()
    if not selected_items:
        messagebox.showwarning("Advertencia", "Por favor, selecciona al menos un cliente.")
        return

    # Obtener la fecha de última actualización desde config.json
    last_update_date = self.config_manager.get_last_update_date()

    if not last_update_date:
        messagebox.showerror("Error", "No se encontró una fecha de última actualización.\n"
                                      "Por favor, asegúrate de que el archivo Excel ha sido procesado correctamente.")
        return

    # Lista de formatos posibles
    formatos_fecha = [
        "%Y-%m-%d %H:%M:%S.%f",  # Formato con microsegundos
        "%Y-%m-%d %H:%M:%S",      # Formato sin microsegundos
        "%Y-%m-%d %H:%M",         # Solo hasta minutos
        "%Y-%m-%d"                # Solo fecha
    ]

    formatted_update_date = last_update_date  # Si no logra formatear, usar el original
    for formato in formatos_fecha:
        try:
            last_update_date_obj = datetime.strptime(last_update_date, formato)
            formatted_update_date = last_update_date_obj.strftime("%Y-%m-%d %H:%M")
            break  # Si encuentra un formato válido, salir del bucle
        except ValueError:
            continue  # Si falla, probar con el siguiente

    mensaje = f"¿Estás seguro de enviar los reportes? Última actualización: {formatted_update_date}"
    respuesta_envio = messagebox.askyesno("Confirmar Envío", mensaje)

    if not respuesta_envio:
        return

    # Iniciar el envío de reportes en un hilo separado
    threading.Thread(target=self.process_sending_reports, args=(selected_items, tree), daemon=True).start()

def open_reports_in_excel(self):
    """
    Abre el archivo de reporte XLSX en Excel para los clientes seleccionados
    que tengan un archivo disponible (tag = 'has_report').
    """
    import os  # Para utilizar os.startfile en Windows

    # Determinar la pestaña activa (Por Lista o Todos los Clientes)
    notebook = self.client_manager.nametowidget("!notebook")
    current_tab = notebook.tab(notebook.select(), "text")

    if current_tab == "Por Lista":
        tree = self.tree_clients_por_lista
    elif current_tab == "Todos los Clientes":
        tree = self.tree_all_clients
    else:
        messagebox.showwarning("Advertencia", "No se pudo determinar la pestaña activa.")
        return

    selected_items = tree.selection()
    if not selected_items:
        messagebox.showwarning("Advertencia", "Por favor, selecciona al menos un cliente.")
        return

    # Intentar abrir cada reporte seleccionado
    for item in selected_items:
        values = tree.item(item, 'values')
        tags = tree.item(item, 'tags')

        # Solo abrimos si el cliente tiene el tag 'has_report'
        if 'has_report' in tags:
            client_name = values[1]  # Asumiendo que en la columna 1 (Nombre) se encuentra el nombre del cliente
            file_path = self.processor.get_client_file(client_name)
            if file_path:
                try:
                    # En Windows, os.startfile abre el archivo con la aplicación predeterminada (Excel).
                    # En otros SO (Linux, Mac) puedes usar subprocess con 'xdg-open' o 'open', respectivamente.
                    os.startfile(file_path)
                except Exception as e:
                    messagebox.showerror(
                        "Error",
                        f"No se pudo abrir el archivo para '{client_name}'.\nRuta: {file_path}\nError: {e}"
                    )
            else:
                messagebox.showinfo(
                    "Reporte no encontrado",
                    f"El cliente '{client_name}' no tiene un archivo de reporte disponible."
                )
        else:
            # Si no tiene el tag 'has_report', no hacemos nada con este cliente.
            pass

def select_all_with_reports_in_all_clients(self):
    """
    Selecciona todos los clientes que tengan el tag 'has_report'
    en la pestaña 'Todos los Clientes'.
    """
    tree = self.tree_all_clients

    # Limpiar cualquier selección previa
    tree.selection_clear()

    # Iterar todos los ítems (clientes) en la tabla
    for item_id in tree.get_children():
        # Obtener las etiquetas (tags) de cada ítem
        tags = tree.item(item_id, 'tags')
        # Verificar si 'has_report' está presente
        if 'has_report' in tags:
            # Agregar este ítem a la selección
            tree.selection_add(item_id)
