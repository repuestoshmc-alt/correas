# Buscador de correas desde Excel

Este proyecto ahora usa **pandas** para leer el archivo Excel y se ejecuta de forma simple con:

```bash
python main.py
```

## Requisitos

```bash
pip install -r requirements.txt
```

## Qué hace `main.py`

1. Carga un archivo Excel real llamado `correas.xlsx` usando `pandas`.
2. Pide por consola:
   - `largo`
   - `ancho` (acepta decimal o fracción, por ejemplo `0.625` o `5/8`)
   - `tolerancia` (por ejemplo `1.5`)
3. Ejecuta la búsqueda.
4. Muestra los resultados en pantalla.

## Formato del archivo `correas.xlsx`

El Excel debe contener estas columnas:

- `Codigo`
- `Original`
- `Marca`
- `Largo_in`
- `Ancho_in`
- `Tipo`

## Regla de búsqueda

- filtra por `Ancho_in` exacto,
- acepta `Largo_in` dentro de la tolerancia ingresada,
- ordena por cercanía,
- muestra: `Codigo - Original - Largo - Diferencia`.

## Uso

1. Copiá tu archivo `correas.xlsx` en la misma carpeta del proyecto.
2. Ejecutá:

```bash
python main.py
```

3. Ingresá, por ejemplo:

```text
Ingresá el largo: 75
Ingresá el ancho (ej: 0.625 o 5/8): 5/8
Ingresá la tolerancia en pulgadas (default 1.5): 1.5
```

## Archivos

- `main.py`: punto de entrada principal para consola.
- `buscar_correas.py`: lógica de lectura con pandas, tolerancia configurable y búsqueda.
