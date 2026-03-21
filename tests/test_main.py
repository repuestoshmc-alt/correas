from pathlib import Path

import main
from buscar_correas import Correa


def test_main_ejecuta_busqueda_con_correas_xlsx(monkeypatch, tmp_path) -> None:
    archivo = tmp_path / "correas.xlsx"
    archivo.write_text("placeholder")

    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(main, "cargar_correas_desde_excel", lambda _ruta: [
        Correa("A1", "OEM-1", "MarcaX", 75.0, 0.625, "V"),
        Correa("A2", "OEM-2", "MarcaX", 74.0, 0.625, "V"),
    ])
    respuestas = iter(["75", "5/8", "1.5"])
    monkeypatch.setattr("builtins.input", lambda _prompt: next(respuestas))

    salida = main.ejecutar_busqueda()

    assert "Codigo - Original - Largo - Diferencia" in salida
    assert "A1 - OEM-1 - 75.000 - 0.000" in salida
