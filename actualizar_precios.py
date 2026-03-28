from __future__ import annotations

import argparse
import logging
import re
import sys
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import pandas as pd

PRICE_OUTPUT_COLUMN = "Precio bruto sin iva"
BASE_CODE_COLUMN = "Codigo"
MISSING_OUTPUT_FILENAME = "no_encontrados.xlsx"

logger = logging.getLogger(__name__)


class ConfiguracionInvalidaError(ValueError):
    """Error para columnas requeridas faltantes o ambiguas."""


@dataclass(frozen=True)
class ResumenActualizacion:
    total_productos: int
    con_precio: int
    sin_precio: int


@dataclass(frozen=True)
class ColumnasDetectadas:
    codigo_proveedor: str
    precio_proveedor: str


def normalizar_texto(texto: object) -> str:
    raw = str(texto or "").strip().lower()
    sin_acentos = "".join(
        c for c in unicodedata.normalize("NFD", raw) if unicodedata.category(c) != "Mn"
    )
    return re.sub(r"\s+", " ", sin_acentos)


def limpiar_codigo_base(codigo: object) -> str:
    texto = str(codigo or "").strip()
    if texto.upper().startswith("KEE_"):
        texto = texto[4:]
    return texto.strip().upper()


def limpiar_codigo_general(codigo: object) -> str:
    return str(codigo or "").strip().upper()


def es_columna_codigo(nombre_columna: str) -> bool:
    nombre = normalizar_texto(nombre_columna)
    tokens = {
        "codigo",
        "cod",
        "sku",
        "articulo",
        "producto",
        "item",
        "referencia",
        "ref",
    }
    return any(token in nombre for token in tokens)


def puntaje_columna_precio(nombre_columna: str) -> int:
    nombre = normalizar_texto(nombre_columna)
    puntaje = 0
    if "costo" in nombre:
        puntaje += 3
    if "neto" in nombre:
        puntaje += 3
    if "sin iva" in nombre:
        puntaje += 4
    if "precio" in nombre:
        puntaje += 1
    if "iva" in nombre and "sin iva" not in nombre:
        puntaje -= 2
    return puntaje


def detectar_columna_codigo(df: pd.DataFrame) -> str:
    columnas = list(df.columns)

    exactas = [c for c in columnas if normalizar_texto(c) == "codigo"]
    if exactas:
        return exactas[0]

    candidatas = [c for c in columnas if es_columna_codigo(c)]
    if len(candidatas) == 1:
        return candidatas[0]
    if len(candidatas) > 1:
        raise ConfiguracionInvalidaError(
            f"Hay múltiples columnas candidatas a código en proveedor.xlsx: {candidatas}."
        )

    raise ConfiguracionInvalidaError(
        "No se pudo detectar automáticamente la columna de código en proveedor.xlsx."
    )


def detectar_columna_precio(df: pd.DataFrame) -> str:
    columnas = list(df.columns)
    puntajes = {c: puntaje_columna_precio(c) for c in columnas}
    mejores = sorted(columnas, key=lambda c: (puntajes[c], normalizar_texto(c)), reverse=True)

    if not mejores or puntajes[mejores[0]] <= 0:
        raise ConfiguracionInvalidaError(
            "No se pudo detectar una columna de precio sin IVA en proveedor.xlsx "
            "(se buscó: costo, neto, sin iva)."
        )

    top = mejores[0]
    empatadas = [c for c in columnas if puntajes[c] == puntajes[top] and puntajes[c] > 0]
    if len(empatadas) > 1:
        # Preferir la que contenga explícitamente "sin iva"
        sin_iva = [c for c in empatadas if "sin iva" in normalizar_texto(c)]
        if len(sin_iva) == 1:
            return sin_iva[0]
        raise ConfiguracionInvalidaError(
            f"Hay múltiples columnas candidatas a precio sin IVA en proveedor.xlsx: {empatadas}."
        )

    return top


def parsear_precio(valor: object) -> float | None:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None

    if isinstance(valor, (int, float)):
        return float(valor)

    texto = str(valor).strip()
    if not texto:
        return None

    texto = re.sub(r"[^\d,.-]", "", texto)
    if not texto:
        return None

    if "," in texto and "." in texto:
        if texto.rfind(",") > texto.rfind("."):
            texto = texto.replace(".", "").replace(",", ".")
        else:
            texto = texto.replace(",", "")
    elif "," in texto:
        partes = texto.split(",")
        if len(partes[-1]) in {1, 2}:
            texto = texto.replace(".", "").replace(",", ".")
        else:
            texto = texto.replace(",", "")
    else:
        texto = texto.replace(",", "")

    try:
        return float(texto)
    except ValueError:
        return None


def construir_mapa_precios(df_proveedor: pd.DataFrame, col_codigo: str, col_precio: str) -> dict[str, float]:
    codigos = df_proveedor[col_codigo].map(limpiar_codigo_general)
    precios = df_proveedor[col_precio].map(parsear_precio)

    resultado: dict[str, float] = {}
    for codigo, precio in zip(codigos, precios):
        if not codigo or precio is None:
            continue
        if codigo in resultado:
            logger.warning("Código duplicado detectado: %s (se usa el último valor)", codigo)
        resultado[codigo] = precio
    return resultado


def construir_dataframe_no_encontrados(df_base: pd.DataFrame) -> pd.DataFrame:
    columnas_requeridas = [BASE_CODE_COLUMN, "Nombre"]
    no_encontrados = df_base[df_base[PRICE_OUTPUT_COLUMN].isna()].copy()

    for columna in columnas_requeridas:
        if columna not in no_encontrados.columns:
            no_encontrados[columna] = ""

    return no_encontrados[columnas_requeridas]


def exportar_no_encontrados(df_base: pd.DataFrame, output_path: Path) -> Path:
    no_encontrados = construir_dataframe_no_encontrados(df_base)
    destino = output_path.with_name(MISSING_OUTPUT_FILENAME)
    no_encontrados.to_excel(destino, index=False)
    return destino


def calcular_resumen(df_base: pd.DataFrame) -> ResumenActualizacion:
    total = len(df_base)
    con_precio = int(df_base[PRICE_OUTPUT_COLUMN].notna().sum())
    sin_precio = total - con_precio
    return ResumenActualizacion(total_productos=total, con_precio=con_precio, sin_precio=sin_precio)


def loguear_resumen(columnas: ColumnasDetectadas, resumen: ResumenActualizacion) -> None:
    logger.info("Columna código proveedor: %s", columnas.codigo_proveedor)
    logger.info("Columna precio proveedor: %s", columnas.precio_proveedor)
    logger.info("Total productos: %s", resumen.total_productos)
    logger.info("Con precio: %s", resumen.con_precio)
    logger.info("Sin precio: %s", resumen.sin_precio)


def validar_columnas_base(df_base: pd.DataFrame) -> None:
    if BASE_CODE_COLUMN not in df_base.columns:
        raise ConfiguracionInvalidaError(
            f"lista_base.xlsx debe contener la columna '{BASE_CODE_COLUMN}'."
        )


def actualizar_precios(
    proveedor_path: str | Path = "proveedor.xlsx",
    lista_base_path: str | Path = "lista_base.xlsx",
    output_path: str | Path = "resultado.xlsx",
) -> Path:
    proveedor_path = Path(proveedor_path)
    lista_base_path = Path(lista_base_path)
    output_path = Path(output_path)

    df_proveedor = pd.read_excel(proveedor_path)
    df_base = pd.read_excel(lista_base_path)

    validar_columnas_base(df_base)

    columnas = ColumnasDetectadas(
        codigo_proveedor=detectar_columna_codigo(df_proveedor),
        precio_proveedor=detectar_columna_precio(df_proveedor),
    )

    precios_por_codigo = construir_mapa_precios(
        df_proveedor,
        columnas.codigo_proveedor,
        columnas.precio_proveedor,
    )

    codigos_base = df_base[BASE_CODE_COLUMN].map(limpiar_codigo_base)
    df_base[PRICE_OUTPUT_COLUMN] = codigos_base.map(precios_por_codigo)

    df_base.to_excel(output_path, index=False)
    exportar_no_encontrados(df_base, output_path)
    loguear_resumen(columnas, calcular_resumen(df_base))
    return output_path


def crear_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Procesa Excel de proveedor y actualiza precios sin IVA en lista base."
    )
    parser.add_argument("--proveedor", default="proveedor.xlsx", help="Ruta a proveedor.xlsx")
    parser.add_argument("--base", default="lista_base.xlsx", help="Ruta a lista_base.xlsx")
    parser.add_argument("--output", default="resultado.xlsx", help="Ruta de salida resultado.xlsx")
    return parser


def main(argv: Iterable[str] | None = None) -> None:
    logging.basicConfig(level=logging.INFO, format="[%(levelname)s] %(message)s")

    argumentos = list(argv) if argv is not None else sys.argv[1:]
    if not argumentos:
        proveedor = input("Ingrese ruta archivo proveedor: ").strip() or "proveedor.xlsx"
        base = input("Ingrese ruta archivo base: ").strip() or "lista_base.xlsx"
        salida = actualizar_precios(proveedor, base, "resultado.xlsx")
    else:
        parser = crear_parser()
        args = parser.parse_args(argumentos)
        salida = actualizar_precios(args.proveedor, args.base, args.output)

    print(f"Archivo generado: {salida}")


if __name__ == "__main__":
    main()
