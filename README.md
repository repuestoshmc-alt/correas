# Actualizador de precios desde Excel (Python + pandas)

Este proyecto procesa dos archivos Excel y genera un tercero con precios actualizados.

## Archivos de entrada/salida

- **Entrada 1:** `proveedor.xlsx`
- **Entrada 2:** `lista_base.xlsx`
- **Salida:** `resultado.xlsx`

## Reglas implementadas

1. En `lista_base.xlsx` se usa la columna **`Codigo`**.
2. Antes de comparar, se elimina el prefijo **`KEE_`** del código base.
3. En `proveedor.xlsx` se detecta automáticamente:
   - columna de código (ej: `Codigo`, `Cod`, `SKU`, etc.)
   - columna de precio sin IVA (buscando términos como `costo`, `neto`, `sin iva`)
4. Convierte precios de texto a número (por ejemplo `"$ 42.148,76"` → `42148.76`).
5. Hace match por código.
6. Completa la columna **`Precio bruto sin iva`** en `lista_base.xlsx`.
7. Si no hay coincidencia, deja la celda vacía.
8. Genera archivo adicional **`no_encontrados.xlsx`** con columnas `Codigo` y `Nombre`.
9. Informa por consola un resumen con logging (`[INFO]` / `[WARNING]`), incluyendo columnas detectadas y totales.
10. Si hay códigos duplicados en proveedor, conserva el último precio y muestra advertencia.

## Requisitos

```bash
pip install -r requirements.txt
```

## Ejecución

Con nombres de archivo por defecto:

```bash
python actualizar_precios.py
```

Al ejecutar sin parámetros, el script solicita por consola:

- ruta del archivo proveedor
- ruta del archivo base

Con rutas personalizadas:

```bash
python actualizar_precios.py --proveedor proveedor.xlsx --base lista_base.xlsx --output resultado.xlsx
```

Al finalizar, se genera el archivo `resultado.xlsx`.

## Estructura

- `actualizar_precios.py`: lógica completa de detección de columnas, normalización y exportación.
- `tests/test_actualizar_precios.py`: pruebas unitarias de parseo, detección y proceso principal.
