# Validaciones TT — Mejora continua

App web para el seguimiento de **Validaciones** (mejoras de rendimiento, precio, calidad y disponibilidad de piezas).

- **Frontend:** `index.html` publicado en GitHub Pages.
- **Datos:** planilla de Google "Plan de Mejora Continua" (hojas `VAL_*`), archivos y fotos en Drive.
- **Backend:** `Code.gs` (Apps Script vinculado a la planilla) + `index.html` (en Apps Script va como archivo HTML llamado `Index`).

## Acceso
La primera vez pide la **clave de acceso** (en la planilla: menú *Validaciones → Ver clave de acceso*).

## Actualizar backend
Pegar `Code.gs` e `index.html` (como `Index`) en *Extensiones → Apps Script*, y *Implementar → Gestionar implementaciones → Editar → Nueva versión* (la URL no cambia).
