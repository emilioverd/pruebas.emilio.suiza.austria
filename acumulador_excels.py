"""
Acumula automáticamente múltiples archivos Excel en un solo archivo.

Uso rápido:
    python acumulador_excels.py --origen "C:/ruta/carpeta" --salida "C:/ruta/salida/acumulado.xlsx"

Opcionalmente puedes indicar hoja, patrón y columnas de control.
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Iterable

import pandas as pd


EXTENSIONES_VALIDAS = {".xlsx", ".xls", ".xlsm"}


def listar_archivos_excel(carpeta: Path, patron: str = "*.xlsx") -> list[Path]:
    """Devuelve archivos Excel de una carpeta ordenados por nombre."""
    archivos = [
        p
        for p in carpeta.glob(patron)
        if p.is_file() and p.suffix.lower() in EXTENSIONES_VALIDAS
    ]
    return sorted(archivos, key=lambda p: p.name.lower())


def leer_archivo(path_excel: Path, hoja: str | int | None = 0) -> pd.DataFrame:
    """Lee un archivo Excel y agrega metadatos de trazabilidad."""
    df = pd.read_excel(path_excel, sheet_name=hoja)
    df["__archivo_origen"] = path_excel.name
    return df


def acumular_archivos(archivos: Iterable[Path], hoja: str | int | None = 0) -> pd.DataFrame:
    """Concatena todos los Excel en un único DataFrame."""
    dataframes: list[pd.DataFrame] = []

    for archivo in archivos:
        try:
            dataframes.append(leer_archivo(archivo, hoja=hoja))
        except Exception as exc:  # noqa: BLE001
            print(f"[ADVERTENCIA] No se pudo leer '{archivo.name}': {exc}")

    if not dataframes:
        return pd.DataFrame()

    return pd.concat(dataframes, ignore_index=True)


def exportar(df: pd.DataFrame, salida: Path) -> None:
    """Exporta el acumulado a Excel."""
    salida.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(salida, index=False)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Acumula automáticamente múltiples archivos Excel en un solo archivo.",
    )
    parser.add_argument(
        "--origen",
        required=True,
        type=Path,
        help="Carpeta donde están los Excel a acumular.",
    )
    parser.add_argument(
        "--salida",
        required=True,
        type=Path,
        help="Ruta del archivo Excel de salida.",
    )
    parser.add_argument(
        "--patron",
        default="*.xlsx",
        help="Patrón de búsqueda (ejemplo: '*.xlsx' o 'Ventas_*.xlsm').",
    )
    parser.add_argument(
        "--hoja",
        default="0",
        help="Nombre o índice de hoja (por defecto 0).",
    )
    return parser.parse_args()


def normalizar_hoja(valor: str) -> str | int:
    return int(valor) if valor.isdigit() else valor


def main() -> None:
    args = parse_args()
    carpeta_origen = args.origen
    hoja = normalizar_hoja(args.hoja)

    if not carpeta_origen.exists() or not carpeta_origen.is_dir():
        raise FileNotFoundError(f"La carpeta de origen no existe: {carpeta_origen}")

    archivos = listar_archivos_excel(carpeta_origen, patron=args.patron)
    if not archivos:
        print("No se encontraron archivos para acumular.")
        return

    df_final = acumular_archivos(archivos, hoja=hoja)
    if df_final.empty:
        print("No se pudo acumular información (todas las lecturas fallaron).")
        return

    exportar(df_final, args.salida)
    print(
        f"OK: {len(archivos)} archivo(s) acumulado(s) en '{args.salida}'. "
        f"Filas totales: {len(df_final)}"
    )


if __name__ == "__main__":
    main()
