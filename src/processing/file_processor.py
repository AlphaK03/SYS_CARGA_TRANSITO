# src/processing/file_processor.py
import os
import pandas as pd
import random

class FileProcessor:
    def __init__(self, file_path):
        self.file_path = file_path
        self.output_dir = os.path.join('data', 'Clientes_Separados')

    def process_file(self):
        # Leer el archivo Excel con la fila correcta como encabezado
        df = pd.read_excel(self.file_path, sheet_name='Report SI SYSTEM - Customers', header=7)
        # Remover espacios en blanco de los nombres de las columnas
        df.columns = df.columns.str.strip()

        # Verificar si 'Consignee' está en las columnas
        if 'Consignee' not in df.columns:
            raise ValueError(
                f"La columna 'Consignee' no se encontró. Las columnas disponibles son: {df.columns.tolist()}")

        # Filtrar por clientes únicos en la columna 'Consignee'
        clientes_unicos = df['Consignee'].dropna().unique()

        # Crear un directorio para guardar los archivos
        os.makedirs(self.output_dir, exist_ok=True)

        # Crear un archivo separado por cada cliente
        for cliente in clientes_unicos:
            df_cliente = df[df['Consignee'] == cliente]
            output_file = os.path.join(self.output_dir, f'{cliente}.xlsx')
            df_cliente.to_excel(output_file, index=False)

        return clientes_unicos

    def get_random_file(self):
        archivos = os.listdir(self.output_dir)
        if archivos:
            archivo_a_enviar = random.choice(archivos)
            archivo_path = os.path.join(self.output_dir, archivo_a_enviar)
            return archivo_path
        else:
            return None
