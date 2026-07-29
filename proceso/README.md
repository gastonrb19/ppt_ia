# Proceso de generación

Pipeline reproducible que arma las dos presentaciones desde los Excel de la semana.

## Pasos

1. **Extraer y filtrar datos** (genera `proceso/_data.json`):

   ```bash
   python3 proceso/extraer_datos.py <carpeta_datos> <dd-mm-yyyy>
   ```
   - Filtros aplicados: excluye tickets `Cerrado` y `Cancelado`; incluye INC + RITM; todas las sociedades.
   - Por defecto lee `Ejecuciones/<hoy>/datos`.

2. **Generar las 2 PPTX**:

   ```bash
   node proceso/generar_ppt.js <carpeta_salida>
   ```
   - Requiere `pptxgenjs` (`npm install pptxgenjs` si falta).
   - Por defecto escribe en `Ejecuciones/<fecha>/`.

## Ejemplos

- **Flujo real** (semana en curso):
  ```bash
  python3 proceso/extraer_datos.py Ejecuciones/05-08-2026/datos 05-08-2026
  node proceso/generar_ppt.js Ejecuciones/05-08-2026
  ```

- **Prueba** (no cuenta como reunión anterior):
  ```bash
  python3 proceso/extraer_datos.py test/29-07-2026/datos 29-07-2026
  node proceso/generar_ppt.js test/29-07-2026
  ```

La especificación completa (mapeos, columnas, estilo, colores) está en `../ESPECIFICACION.md`.
