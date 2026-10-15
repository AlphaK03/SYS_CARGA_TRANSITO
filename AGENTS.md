# Guías del repositorio

## Contexto del proyecto

Esta aplicación de escritorio para Windows descarga un libro de cargas desde SharePoint, ejecuta una macro de Excel, genera reportes por cliente y los envía mediante Microsoft Graph. Usa OAuth2 para autenticación y una base de datos local para gestionar clientes y alias.

## Estructura del proyecto

`src/main.py` inicia la aplicación. La interfaz está en `src/gui/` y los reportes en `src/processing/`. Correo, base de datos, clientes y utilidades están en `src/email_service/`, `src/database/`, `src/manager/` y `src/utils/`. Las plantillas están en `template/`, las imágenes en `resources/` y los archivos de trabajo en `data/`. `tests/` contiene módulos vacíos. `CHC_Reportes.spec` define el ejecutable.

## Comandos de desarrollo, pruebas y compilación

- `python -m venv .venv`: crea el entorno virtual.
- `python -m pip install -r requirements.txt`: instala las dependencias declaradas.
- `python src/main.py`: ejecuta la aplicación en Windows. Las funciones de automatización de Excel requieren Excel y `pywin32` para usar `win32com`.
- `python -m unittest discover -s tests -p "test_*.py"`: busca y ejecuta pruebas. Actualmente no encuentra casos porque los módulos están vacíos.
- `python -m PyInstaller CHC_Reportes.spec`: genera el ejecutable si PyInstaller está instalado.

## Estilo y convenciones

Usa cuatro espacios de sangría en Python. Nombra módulos, funciones y variables en `snake_case`, y clases en `PascalCase`, como `FileProcessor`. Mantén cada lógica en su paquete correspondiente. No hay formateador ni linter configurado; sigue el estilo cercano.

## Pruebas

Al cambiar comportamiento, agrega pruebas en `tests/test_<modulo>.py`. Prioriza reportes, configuración y base de datos. Aísla la interfaz, Excel, la red y las credenciales con simulaciones o datos de prueba. Ejecuta las pruebas antes de enviar cambios y documenta las comprobaciones manuales de Windows. No hay umbral de cobertura.

## Commits y solicitudes de incorporación

Los commits recientes usan descripciones breves en español sobre correcciones, funciones o versiones del ejecutable, sin prefijo obligatorio. Escribe un resumen específico en imperativo, por ejemplo, `Corrige formato de fechas en reportes`. En cada solicitud, explica el cambio, cómo lo verificaste y qué plantillas o datos de clientes afecta. Enlaza la incidencia relacionada si existe e incluye capturas cuando cambie la interfaz.

## Configuración y datos sensibles

Mantén `.env` fuera de los commits; el README documenta `USER_EMAIL` y `APP_PASSWORD`. No incorpores credenciales, cachés de tokens ni bases de datos de clientes. Antes de empaquetar, revisa `CHC_Reportes.spec`: actualmente incluye `.env` en el ejecutable.
