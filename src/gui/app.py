# src/gui/app.py
import tkinter as tk
from tkinter import filedialog, messagebox
from dotenv import load_dotenv
import os

from src.email_service.email_sender import EmailSender
from src.processing.file_processor import FileProcessor


class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Generar Reportes por Cliente y Enviar Correo")
        self.geometry("500x350")
        self.file_path = ''
        self.create_widgets()
        load_dotenv()  # Cargar variables de entorno

    def create_widgets(self):
        # Etiqueta y campo de entrada para seleccionar archivo
        etiqueta_archivo = tk.Label(self, text="Selecciona el archivo XLSX:")
        etiqueta_archivo.pack(pady=5)

        self.entrada_archivo = tk.Entry(self, width=50)
        self.entrada_archivo.pack(pady=5)

        boton_seleccionar = tk.Button(self, text="Seleccionar archivo", command=self.seleccionar_archivo)
        boton_seleccionar.pack(pady=5)

        # Botón para filtrar y generar reportes
        boton_generar = tk.Button(self, text="Filtrar y generar reportes", command=self.filtrar_y_generar_reportes)
        boton_generar.pack(pady=20)

    def seleccionar_archivo(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xlsm")])
        if self.file_path:
            self.entrada_archivo.delete(0, tk.END)
            self.entrada_archivo.insert(0, self.file_path)

    def filtrar_y_generar_reportes(self):
        if not self.file_path:
            messagebox.showwarning("Advertencia", "Por favor, selecciona un archivo.")
            return

        # Obtener las credenciales desde las variables de entorno
        user_email = os.getenv('USER_EMAIL')
        app_password = os.getenv('APP_PASSWORD')

        if not user_email or not app_password:
            messagebox.showwarning("Advertencia", "Por favor, configura tu correo y contraseña de aplicación en el archivo .env.")
            return

        try:
            processor = FileProcessor(self.file_path)
            clientes_unicos = processor.process_file()
            messagebox.showinfo("Éxito",
                                f"Se generaron archivos separados para {len(clientes_unicos)} clientes en '{processor.output_dir}'.")

            # Enviar uno de los archivos por correo como prueba
            archivo_path = processor.get_random_file()
            if archivo_path:
                sender = EmailSender(user_email, app_password)
                success, msg = sender.send_email(
                    user_email,
                    'Prueba de envío de correo',
                    '<h1>Este es un correo de prueba con caracteres especiales</h1>',
                    archivo_path
                )
                if success:
                    messagebox.showinfo("Correo enviado", msg)
                else:
                    messagebox.showerror("Error de envío", f"No se pudo enviar el correo: {msg}")
            else:
                messagebox.showwarning("Advertencia", "No se encontró ningún archivo para enviar.")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error: {str(e)}")
