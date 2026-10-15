# Autenticación y permisos de Microsoft

- `manual_seguridad_oauth.tex` y su PDF: fragmentos de autenticación, correo, SharePoint y almacenamiento de tokens, con explicaciones breves.
- `guia_registro_client_id.tex` y su PDF: navegación por Entra para crear la aplicación en el tenant nuevo, configurar los permisos y copiar los identificadores.

Los documentos se compilan sin acceso al código fuente. La guía de registro
utiliza las nueve capturas numeradas de `capturas_registro/`; conserve esa
carpeta junto al archivo `.tex`. El PDF contiene las imágenes y puede
compartirse por separado.

La guía está dirigida al administrador de TI e incluye los permisos delegados,
el consentimiento administrativo, los datos que debe entregar al desarrollador
y la referencia oficial de Microsoft sobre el carácter público del Client ID.
El registro temporal utilizado para las capturas fue eliminado; su trazabilidad
se documenta en `capturas_registro/README.md`.

## Compilación

Desde esta carpeta:

```powershell
pdflatex -interaction=nonstopmode -halt-on-error manual_seguridad_oauth.tex
pdflatex -interaction=nonstopmode -halt-on-error guia_registro_client_id.tex
```

La carpeta `.build/` contiene los archivos auxiliares de la compilación local.
