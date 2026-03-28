from pathlib import Path

import pandas as pd

from actualizar_precios import (
    MISSING_OUTPUT_FILENAME,
    PRICE_OUTPUT_COLUMN,
    actualizar_precios,
    construir_mapa_precios,
    detectar_columna_codigo,
    detectar_columna_precio,
    parsear_precio,
)


def test_parsear_precio_con_formato_moneda_latam() -> None:
    assert parsear_precio("$ 42.148,76") == 42148.76


def test_detectar_columna_codigo_y_precio() -> None:
    df = pd.DataFrame(
        {
            "SKU": ["ABC"],
            "Costo Neto": ["$ 1.234,56"],
            "Descripcion": ["Producto"],
        }
    )

    assert detectar_columna_codigo(df) == "SKU"
    assert detectar_columna_precio(df) == "Costo Neto"


def test_actualizar_precios_genera_resultado_con_match_y_vacio(tmp_path: Path) -> None:
    proveedor = tmp_path / "proveedor.xlsx"
    base = tmp_path / "lista_base.xlsx"
    out = tmp_path / "resultado.xlsx"

    pd.DataFrame(
        {
            "Cod": ["12345", "99999"],
            "Precio costo sin IVA": ["$ 42.148,76", "$ 100,00"],
        }
    ).to_excel(proveedor, index=False)

    pd.DataFrame(
        {
            "Codigo": ["KEE_12345", "KEE_00000"],
            "Descripcion": ["A", "B"],
        }
    ).to_excel(base, index=False)

    actualizar_precios(proveedor, base, out)

    resultado = pd.read_excel(out)
    assert PRICE_OUTPUT_COLUMN in resultado.columns
    assert resultado.loc[0, PRICE_OUTPUT_COLUMN] == 42148.76
    assert pd.isna(resultado.loc[1, PRICE_OUTPUT_COLUMN])

    no_encontrados = pd.read_excel(tmp_path / MISSING_OUTPUT_FILENAME)
    assert list(no_encontrados.columns) == ["Codigo", "Nombre"]
    assert len(no_encontrados) == 1
    assert no_encontrados.loc[0, "Codigo"] == "KEE_00000"


def test_construir_mapa_precios_loguea_duplicados_y_usa_ultimo_valor(caplog) -> None:
    df = pd.DataFrame(
        {
            "Cod": ["10-09628-001", "10-09628-001"],
            "Costo Neto": ["$ 10,00", "$ 12,00"],
        }
    )

    with caplog.at_level("WARNING"):
        resultado = construir_mapa_precios(df, "Cod", "Costo Neto")

    assert resultado["10-09628-001"] == 12.0
    assert "Código duplicado detectado: 10-09628-001 (se usa el último valor)" in caplog.text
