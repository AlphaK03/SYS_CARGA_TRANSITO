# Revisión acotada de DEV — Astra

Fecha: 4 de octubre de 2026.

Se revisó el código actual del directorio de trabajo, incluidos sus cambios
preexistentes. Esta revisión no modifica el comportamiento de la aplicación.

## Comprobaciones

```powershell
python -m unittest discover -s tests -p "test_*.py" -v
```

Se usó el intérprete disponible con `python`, que tiene las dependencias del
proyecto. El `.venv` local aún no las tiene instaladas.

Resultado: **11 casos: 6 satisfactorios y 5 bugs reproducidos**. Los cinco
se marcan explícitamente con `unittest.expectedFailure`; el resultado
`OK (expected failures=5)` no significa que esos bugs estén corregidos.
Se verificó que cada fallo corresponde a la aserción del comportamiento
esperado, no a un problema de preparación de las pruebas.

Casos satisfactorios: persistencia de configuración; eliminación de cliente
con sus alias y asociaciones; exportación básica de fechas y ocultación de
columnas internas; composición de correo y adjunto con transporte simulado;
bloqueo del envío sin token; manejo de rechazo simulado del servidor.

Todas las entradas son ficticias. Los libros, adjuntos, configuración y base
SQLite se crearon en una carpeta temporal. Se bloquearon conexiones de red,
acceso a keyring y lanzamiento real de Excel. El transporte de correo fue
sustituido por mocks. No se enviaron mensajes ni se ejecutaron macros reales.

## Bugs confirmados y cambios pequeños posibles

1. **El flujo combinado continúa hacia el envío tras fallar la generación.**
   En `src/gui/app.py`, `generate_reports_simplified` captura el error y retorna
   sin indicar fracaso; `filter_generate_and_send` llama después a
   `send_mails_filtered`. Una macro simulada falló y la llamada al envío
   simulado ocurrió una vez. Si quedan reportes parciales, podrían procesarse.
   Ajuste propuesto: devolver éxito/fracaso y enviar únicamente tras éxito.

2. **Se pierde texto en columnas detectadas como fechas.**
   En `src/processing/file_processor.py`, `_save_excel_with_format` convierte
   valores no interpretables en fechas vacías. Con dos fechas y `PENDIENTE`,
   el archivo generado contiene una celda vacía donde estaba `PENDIENTE`.
   Ajuste propuesto: conservar el valor original cuando no se pueda convertir;
   mantener el formato de fecha para los valores válidos.

3. **Excel no se cierra si falla la macro.**
   En `FileProcessor.run_macro`, `Quit()` solo se ejecuta en el camino de éxito.
   Tras tres fallos simulados, no se llamó a `Quit()`.
   Ajuste propuesto: liberar libro e instancia en `finally`, sin guardar
   cambios parciales cuando la operación falla.

4. **Una fecha válida sin microsegundos produce un error.**
   `Application.filter_and_generate` exige `%Y-%m-%d %H:%M:%S.%f`.
   Con `2026-10-04 12:30:00`, los pasos de generación simulados terminan pero
   la interfaz informa un error. `send_mails` contiene el mismo formato
   estricto. Ajuste propuesto: aceptar fechas con y sin fracción de segundo.

5. **No se elimina la tabla previa de una plantilla.**
   `_save_excel_with_format` intenta `ws._tables.remove(...)` y oculta la
   excepción. Una plantilla con una tabla termina con dos tablas superpuestas
   en el archivo generado. Ajuste propuesto: eliminar por nombre mediante
   `ws.tables` antes de añadir la tabla nueva.

## Observación adicional

Las llamadas HTTP de correo, perfil y descarga no establecen `timeout`.
Una respuesta que no llegue puede dejar la operación esperando sin un límite
definido por la aplicación. Se observó en el código; no se provocó una espera
real. Un límite de conexión/lectura sería un ajuste pequeño. No conviene añadir
reintentos automáticos de envío sin contemplar el riesgo de duplicados.

No se probaron autenticación real, entrega de correo, permisos del tenant,
macros de Excel ni todas las pantallas. No se generó un ejecutable ni se
modificaron plantillas, credenciales o datos de clientes.
