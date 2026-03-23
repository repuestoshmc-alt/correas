from pathlib import Path
from types import SimpleNamespace

import pytest

import buscar_correas
from buscar_correas import (
    _parsear_numero_ingresado,
    buscar_correas_por_ancho,
    cargar_correas_desde_excel,
    formatear_resultados,
    indexar_correas_por_ancho,
)


class FakeDataFrame:
    def __init__(self, records):
        self._records = records
        self.columns = list(records[0].keys()) if records else []

    def __getitem__(self, columns):
        if isinstance(columns, list):
            return FakeDataFrame([{column: record.get(column) for column in columns} for record in self._records])
        raise TypeError("Unsupported key type")

    def to_dict(self, orient="records"):
        assert orient == "records"
        return self._records


def test_cargar_correas_desde_excel_usa_pandas(monkeypatch, tmp_path: Path) -> None:
    archivo = tmp_path / "correas.xlsx"
    archivo.write_text("placeholder")

    fake_df = FakeDataFrame(
        [
            {"Codigo": "A1", "Original": "OEM-1", "Marca": "MarcaX", "Largo_in": 75, "Ancho_in": 0.625, "Tipo": "V"},
            {"Codigo": "A2", "Original": "OEM-2", "Marca": "MarcaX", "Largo_in": 74, "Ancho_in": 0.625, "Tipo": "V"},
        ]
    )

    llamadas = []

    def fake_read_excel(ruta, engine):
        llamadas.append((ruta, engine))
        return fake_df

    monkeypatch.setattr(buscar_correas, "_obtener_modulo", lambda nombre: SimpleNamespace(read_excel=fake_read_excel) if nombre == "pandas" else object())

    correas = cargar_correas_desde_excel(archivo)

    assert len(correas) == 2
    assert llamadas == [(archivo, "openpyxl")]


def test_busqueda_filtra_por_ancho_y_tolerancia_y_ordena_por_cercania() -> None:
    correas = [
        buscar_correas.Correa("A1", "OEM-1", "MarcaX", 75.0, 0.625, "V"),
        buscar_correas.Correa("A2", "OEM-2", "MarcaX", 74.0, 0.625, "V"),
        buscar_correas.Correa("A3", "OEM-3", "MarcaY", 76.4, 0.625, "V"),
        buscar_correas.Correa("A4", "OEM-4", "MarcaY", 79.0, 0.625, "V"),
        buscar_correas.Correa("B1", "OEM-5", "MarcaZ", 75.0, 0.5, "V"),
    ]

    indice = indexar_correas_por_ancho(correas)
    resultados = buscar_correas_por_ancho(indice, largo_objetivo=75.0, ancho_objetivo=0.625)

    assert [resultado.codigo for resultado in resultados] == ["A1", "A2", "A3"]
    assert [round(resultado.diferencia, 3) for resultado in resultados] == [0.0, 1.0, 1.4]


def test_formateo_muestra_columnas_esperadas() -> None:
    correas = [buscar_correas.Correa("A1", "OEM-1", "MarcaX", 75.0, 0.625, "V")]
    resultados = buscar_correas_por_ancho(indexar_correas_por_ancho(correas), 75.0, 0.625)
    salida = formatear_resultados(resultados)

    assert "Codigo - Original - Largo - Diferencia" in salida
    assert "A1 - OEM-1 - 75.000 - 0.000" in salida


def test_parsea_numeros_con_coma_decimal() -> None:
    assert _parsear_numero_ingresado("0,625", "ancho") == 0.625


def test_parsea_fracciones() -> None:
    assert _parsear_numero_ingresado("5/8", "ancho") == 0.625


def test_soporta_tolerancia_configurable() -> None:
    correas = [
        buscar_correas.Correa("A1", "OEM-1", "MarcaX", 75.0, 0.625, "V"),
        buscar_correas.Correa("A2", "OEM-2", "MarcaX", 74.0, 0.625, "V"),
        buscar_correas.Correa("A3", "OEM-3", "MarcaY", 76.4, 0.625, "V"),
    ]
    indice = indexar_correas_por_ancho(correas)
    resultados = buscar_correas_por_ancho(indice, largo_objetivo=75.0, ancho_objetivo=0.625, tolerancia=1.0)

    assert [resultado.codigo for resultado in resultados] == ["A1", "A2"]


def test_error_si_faltan_columnas(monkeypatch, tmp_path: Path) -> None:
    archivo = tmp_path / "correas.xlsx"
    archivo.write_text("placeholder")

    fake_df = FakeDataFrame([{"Codigo": "A1", "Original": "OEM-1"}])
    monkeypatch.setattr(buscar_correas, "_leer_excel_con_pandas", lambda _ruta: fake_df)

    with pytest.raises(buscar_correas.ArchivoInvalidoError):
        cargar_correas_desde_excel(archivo)
