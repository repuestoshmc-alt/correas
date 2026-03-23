from __future__ import annotations

from pathlib import Path

from buscar_correas import (
    ArchivoInvalidoError,
    DEFAULT_TOLERANCE,
    _parsear_numero_ingresado,
    buscar_correas_por_ancho,
    cargar_correas_desde_excel,
    formatear_resultados,
    indexar_correas_por_ancho,
)


ARCHIVO_EXCEL = Path("correas.xlsx")


def ejecutar_busqueda() -> str:
    if not ARCHIVO_EXCEL.exists():
        raise FileNotFoundError(
            "No se encontró el archivo 'correas.xlsx' en la carpeta actual. "
            "Copialo junto a main.py antes de ejecutar el sistema."
        )

    correas = cargar_correas_desde_excel(ARCHIVO_EXCEL)
    indice = indexar_correas_por_ancho(correas)

    largo = _parsear_numero_ingresado(input("Ingresá el largo: "), "largo")
    ancho = _parsear_numero_ingresado(input("Ingresá el ancho (ej: 0.625 o 5/8): "), "ancho")
    tolerancia = _parsear_numero_ingresado(
        input(f"Ingresá la tolerancia en pulgadas (default {DEFAULT_TOLERANCE}): ") or str(DEFAULT_TOLERANCE),
        "tolerancia",
    )

    resultados = buscar_correas_por_ancho(
        indice,
        largo_objetivo=largo,
        ancho_objetivo=ancho,
        tolerancia=tolerancia,
    )
    return formatear_resultados(resultados)


def main() -> None:
    print("Buscador de correas")
    print("Archivo esperado: correas.xlsx")

    try:
        salida = ejecutar_busqueda()
    except (ArchivoInvalidoError, FileNotFoundError, ValueError) as exc:
        print(f"Error: {exc}")
        return

    print()
    print(salida)


if __name__ == "__main__":
    main()
