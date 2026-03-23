from __future__ import annotations

import importlib
import importlib.util
import re
import tkinter as tk
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, DefaultDict, Iterable, List


REQUIRED_COLUMNS = ["Codigo", "Original", "Marca", "Largo_in", "Ancho_in", "Tipo"]
DEFAULT_TOLERANCE = 1.5


@dataclass(frozen=True)
class Correa:
    codigo: str
    original: str
    marca: str
    largo_in: float
    ancho_in: float
    tipo: str


@dataclass(frozen=True)
class ResultadoBusqueda:
    codigo: str
    original: str
    largo_in: float
    diferencia: float


class ArchivoInvalidoError(ValueError):
    pass


class EntradaInvalidaError(ValueError):
    pass


def _normalizar_numero(valor: object) -> float:
    if valor is None:
        raise ArchivoInvalidoError("Se encontró una celda vacía en una columna numérica requerida.")

    texto = str(valor).strip().replace(",", ".")
    if not texto or texto.lower() == "nan":
        raise ArchivoInvalidoError("Se encontró un valor numérico vacío.")

    try:
        return float(texto)
    except ValueError as exc:
        raise ArchivoInvalidoError(f"No se pudo interpretar el número '{valor}'.") from exc


def _parsear_numero_ingresado(valor: str, nombre_campo: str) -> float:
    texto = valor.strip().replace(",", ".")
    if not texto:
        raise EntradaInvalidaError(f"Debés ingresar un valor para {nombre_campo}.")

    if "/" in texto:
        partes = texto.split("/")
        if len(partes) != 2:
            raise EntradaInvalidaError(f"El valor de {nombre_campo} no es válido: '{valor}'.")

        numerador = partes[0].strip()
        denominador = partes[1].strip()
        try:
            numerador_float = float(numerador)
            denominador_float = float(denominador)
        except ValueError as exc:
            raise EntradaInvalidaError(f"El valor de {nombre_campo} no es válido: '{valor}'.") from exc

        if denominador_float == 0:
            raise EntradaInvalidaError(f"El denominador de {nombre_campo} no puede ser 0.")

        return numerador_float / denominador_float

    try:
        return float(texto)
    except ValueError as exc:
        raise EntradaInvalidaError(f"El valor de {nombre_campo} no es válido: '{valor}'.") from exc


def _obtener_modulo(nombre: str) -> Any:
    if importlib.util.find_spec(nombre) is None:
        raise ArchivoInvalidoError(
            f"Falta la dependencia '{nombre}'. Instalá los requisitos con: pip install -r requirements.txt"
        )
    return importlib.import_module(nombre)


def _leer_excel_con_pandas(ruta_excel: Path) -> Any:
    _obtener_modulo("openpyxl")
    pandas = _obtener_modulo("pandas")
    return pandas.read_excel(ruta_excel, engine="openpyxl")


def cargar_correas_desde_excel(ruta_excel: str | Path) -> List[Correa]:
    ruta = Path(ruta_excel)
    if not ruta.exists():
        raise FileNotFoundError(f"No existe el archivo: {ruta}")

    dataframe = _leer_excel_con_pandas(ruta)
    columnas = [str(columna).strip() for columna in getattr(dataframe, "columns", [])]
    faltantes = [columna for columna in REQUIRED_COLUMNS if columna not in columnas]
    if faltantes:
        raise ArchivoInvalidoError(f"Faltan columnas obligatorias: {', '.join(faltantes)}")

    correas: List[Correa] = []
    registros = dataframe[REQUIRED_COLUMNS].to_dict(orient="records")
    for fila in registros:
        if all(str(valor).strip() in {"", "nan", "None"} for valor in fila.values()):
            continue

        correas.append(
            Correa(
                codigo=str(fila["Codigo"]).strip(),
                original=str(fila["Original"]).strip(),
                marca=str(fila["Marca"]).strip(),
                largo_in=_normalizar_numero(fila["Largo_in"]),
                ancho_in=_normalizar_numero(fila["Ancho_in"]),
                tipo=str(fila["Tipo"]).strip(),
            )
        )

    if not correas:
        raise ArchivoInvalidoError("El archivo Excel está vacío.")

    return correas


def indexar_correas_por_ancho(correas: Iterable[Correa]) -> DefaultDict[float, List[Correa]]:
    indice: DefaultDict[float, List[Correa]] = defaultdict(list)
    for correa in correas:
        indice[correa.ancho_in].append(correa)
    return indice


def buscar_correas(
    correas: Iterable[Correa],
    largo_objetivo: float,
    ancho_objetivo: float,
    tolerancia: float = DEFAULT_TOLERANCE,
) -> List[ResultadoBusqueda]:
    del ancho_objetivo
    resultados: List[ResultadoBusqueda] = []

    for correa in correas:
        diferencia = abs(correa.largo_in - largo_objetivo)
        if diferencia <= tolerancia:
            resultados.append(
                ResultadoBusqueda(
                    codigo=correa.codigo,
                    original=correa.original,
                    largo_in=correa.largo_in,
                    diferencia=diferencia,
                )
            )

    return sorted(resultados, key=lambda item: (item.diferencia, item.largo_in, item.codigo))


def buscar_correas_por_ancho(
    indice_por_ancho: dict[float, List[Correa]],
    largo_objetivo: float,
    ancho_objetivo: float,
    tolerancia: float = DEFAULT_TOLERANCE,
) -> List[ResultadoBusqueda]:
    return buscar_correas(
        correas=indice_por_ancho.get(ancho_objetivo, []),
        largo_objetivo=largo_objetivo,
        ancho_objetivo=ancho_objetivo,
        tolerancia=tolerancia,
    )


def formatear_resultados(resultados: Iterable[ResultadoBusqueda]) -> str:
    resultados = list(resultados)
    if not resultados:
        return "No se encontraron correas para los criterios indicados."

    encabezado = "Codigo - Original - Largo - Diferencia"
    lineas = [encabezado]
    for resultado in resultados:
        lineas.append(
            f"{resultado.codigo} - {resultado.original} - {resultado.largo_in:.3f} - {resultado.diferencia:.3f}"
        )
    return "\n".join(lineas)


class BuscadorCorreasApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Buscador de correas")
        self.root.geometry("960x600")
        self.root.minsize(820, 480)

        self.ruta_excel_var = tk.StringVar()
        self.largo_var = tk.StringVar()
        self.ancho_var = tk.StringVar()
        self.tolerancia_var = tk.StringVar(value=str(DEFAULT_TOLERANCE))
        self.estado_var = tk.StringVar(value="Seleccioná tu archivo Excel real y cargá los datos.")
        self.correas: List[Correa] = []
        self.indice_por_ancho: DefaultDict[float, List[Correa]] = defaultdict(list)

        self._construir_interfaz()

    def _construir_interfaz(self) -> None:
        contenedor = ttk.Frame(self.root, padding=16)
        contenedor.pack(fill="both", expand=True)
        contenedor.columnconfigure(1, weight=1)
        contenedor.rowconfigure(4, weight=1)

        ttk.Label(contenedor, text="Archivo Excel (.xlsx):").grid(row=0, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(contenedor, textvariable=self.ruta_excel_var).grid(row=0, column=1, sticky="ew", pady=(0, 8))

        botones_archivo = ttk.Frame(contenedor)
        botones_archivo.grid(row=0, column=2, sticky="e", padx=(8, 0), pady=(0, 8))
        ttk.Button(botones_archivo, text="Examinar", command=self.seleccionar_archivo).pack(side="left", padx=(0, 6))
        ttk.Button(botones_archivo, text="Cargar", command=self.cargar_excel).pack(side="left")

        filtros = ttk.LabelFrame(contenedor, text="Búsqueda", padding=12)
        filtros.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(0, 12))
        filtros.columnconfigure(1, weight=1)
        filtros.columnconfigure(3, weight=1)
        filtros.columnconfigure(5, weight=1)

        ttk.Label(filtros, text="Largo:").grid(row=0, column=0, sticky="w")
        ttk.Entry(filtros, textvariable=self.largo_var, width=18).grid(row=0, column=1, sticky="ew", padx=(6, 18))
        ttk.Label(filtros, text="Ancho:").grid(row=0, column=2, sticky="w")
        ttk.Entry(filtros, textvariable=self.ancho_var, width=18).grid(row=0, column=3, sticky="ew", padx=(6, 18))
        ttk.Label(filtros, text="Tolerancia:").grid(row=0, column=4, sticky="w")
        ttk.Entry(filtros, textvariable=self.tolerancia_var, width=18).grid(row=0, column=5, sticky="ew", padx=(6, 18))
        ttk.Button(filtros, text="Buscar", command=self.buscar).grid(row=0, column=6, sticky="e")

        ttk.Label(
            contenedor,
            text="Usa pandas para leer Excel. Acepta ancho decimal o fracción (ej: 5/8) y tolerancia configurable.",
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(0, 8))

        ttk.Label(contenedor, textvariable=self.estado_var).grid(row=3, column=0, columnspan=3, sticky="w", pady=(0, 8))

        columnas = ("codigo", "original", "largo", "diferencia")
        self.tabla = ttk.Treeview(contenedor, columns=columnas, show="headings", height=16)
        self.tabla.heading("codigo", text="Codigo")
        self.tabla.heading("original", text="Original")
        self.tabla.heading("largo", text="Largo")
        self.tabla.heading("diferencia", text="Diferencia")
        self.tabla.column("codigo", width=150, anchor="w")
        self.tabla.column("original", width=320, anchor="w")
        self.tabla.column("largo", width=120, anchor="e")
        self.tabla.column("diferencia", width=120, anchor="e")
        self.tabla.grid(row=4, column=0, columnspan=3, sticky="nsew")

        scroll = ttk.Scrollbar(contenedor, orient="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=scroll.set)
        scroll.grid(row=4, column=3, sticky="ns")

    def seleccionar_archivo(self) -> None:
        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Excel", "*.xlsx")],
        )
        if ruta:
            self.ruta_excel_var.set(ruta)

    def cargar_excel(self) -> None:
        ruta = self.ruta_excel_var.get().strip()
        if not ruta:
            messagebox.showerror("Archivo requerido", "Seleccioná un archivo Excel antes de cargarlo.")
            return

        try:
            self.correas = cargar_correas_desde_excel(ruta)
            self.indice_por_ancho = indexar_correas_por_ancho(self.correas)
        except (ArchivoInvalidoError, FileNotFoundError) as exc:
            messagebox.showerror("No se pudo cargar el archivo", str(exc))
            self.estado_var.set("Error al cargar el archivo Excel.")
            self._limpiar_tabla()
            return

        anchos = len(self.indice_por_ancho)
        self.estado_var.set(
            f"Archivo cargado correctamente. Registros: {len(self.correas)} | Anchos distintos: {anchos}"
        )
        self._limpiar_tabla()

    def buscar(self) -> None:
        if not self.correas:
            messagebox.showwarning("Sin datos", "Primero cargá un archivo Excel con correas.")
            return

        try:
            largo = _parsear_numero_ingresado(self.largo_var.get(), "largo")
            ancho = _parsear_numero_ingresado(self.ancho_var.get(), "ancho")
            tolerancia = _parsear_numero_ingresado(self.tolerancia_var.get(), "tolerancia")
        except EntradaInvalidaError as exc:
            messagebox.showerror("Datos inválidos", str(exc))
            return

        resultados = buscar_correas_por_ancho(
            self.indice_por_ancho,
            largo_objetivo=largo,
            ancho_objetivo=ancho,
            tolerancia=tolerancia,
        )
        self._mostrar_resultados(resultados)

    def _mostrar_resultados(self, resultados: List[ResultadoBusqueda]) -> None:
        self._limpiar_tabla()
        for resultado in resultados:
            self.tabla.insert(
                "",
                "end",
                values=(
                    resultado.codigo,
                    resultado.original,
                    f"{resultado.largo_in:.3f}",
                    f"{resultado.diferencia:.3f}",
                ),
            )

        if resultados:
            self.estado_var.set(f"Se encontraron {len(resultados)} resultado(s).")
        else:
            self.estado_var.set("No se encontraron correas para los criterios indicados.")

    def _limpiar_tabla(self) -> None:
        for item in self.tabla.get_children():
            self.tabla.delete(item)


def main() -> None:
    root = tk.Tk()
    BuscadorCorreasApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
