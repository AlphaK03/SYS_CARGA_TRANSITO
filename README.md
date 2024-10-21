## Configuración del Proyecto

1. **Clonar el repositorio**:

    ```bash
    git clone https://github.com/usuario/proyecto.git
    ```

2. **Crear un entorno virtual y activar**:

    ```bash
    python -m venv venv
    source venv/bin/activate  # En Windows: venv\Scripts\activate
    ```

3. **Instalar las dependencias**:

    ```bash
    pip install -r requirements.txt
    ```

4. **Configurar las variables de entorno**:

    - Crear un archivo `.env` en el directorio raíz del proyecto.
    - Añadir las siguientes líneas al archivo `.env`:

        ```dotenv
        USER_EMAIL=tu_correo@gmail.com
        APP_PASSWORD=tu_contraseña_de_aplicación
        ```

    - **Nota**: Asegúrate de generar una contraseña de aplicación si usas Gmail y la autenticación de dos factores está habilitada.

5. **Ejecutar la aplicación**:

    ```bash
    python src/main.py
    ```

