import tkinter as tk
import ttkbootstrap as ttk

class MultiEmailField(ttk.Frame):
    def __init__(self, master, initial_value="", **kwargs):
        super().__init__(master, **kwargs)
        # Lista interna que almacenará cada email
        self.email_list = []
        # Si se proporciona un valor inicial, separamos por comas y limpiamos espacios
        if initial_value:
            emails = [email.strip() for email in initial_value.split(",") if email.strip()]
            self.email_list.extend(emails)
        # Entrada para agregar nuevo correo
        self.entry = ttk.Entry(self, width=30)
        self.entry.grid(row=0, column=0, sticky='ew')
        # Botón para agregar el correo de la entrada
        self.btn_add = ttk.Button(self, text="Agregar", command=self.add_email)
        self.btn_add.grid(row=0, column=1, padx=(5,0))
        # Listbox para mostrar los correos ingresados
        self.listbox = tk.Listbox(self, height=3)
        self.listbox.grid(row=1, column=0, columnspan=2, sticky='ew', pady=(5,0))
        # Botón para eliminar el correo seleccionado
        self.btn_remove = ttk.Button(self, text="Eliminar", command=self.remove_selected)
        self.btn_remove.grid(row=2, column=0, columnspan=2, pady=(5,0))
        # Cargar los emails iniciales en el listbox (si hubiera)
        for email in self.email_list:
            self.listbox.insert(tk.END, email)

    def add_email(self):
        email = self.entry.get().strip()
        if email and email not in self.email_list:
            self.email_list.append(email)
            self.listbox.insert(tk.END, email)
            self.entry.delete(0, tk.END)
        elif email in self.email_list:
            tk.messagebox.showwarning("Advertencia", "El correo ya se agregó.")

    def remove_selected(self):
        selection = self.listbox.curselection()
        if selection:
            index = selection[0]
            email = self.listbox.get(index)
            self.listbox.delete(index)
            self.email_list.remove(email)

    def get_value(self):
        """Devuelve los correos en una cadena separados por comas."""
        return ",".join(self.email_list)
