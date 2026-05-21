#!/usr/bin/env python3
"""
ETL Monitor de Cooperativas Financieras Ecuador - Segmento 1
Lee los archivos Excel de la SEPS (2018-2026) y genera cooperativas.json.

Uso:
    python etl.py
    python etl.py --source ./ruta/a/excels --output ./ruta/salida.json

Requiere:
    pip install openpyxl
"""

import argparse
import json
import sys
import openpyxl
from datetime import datetime, date
from pathlib import Path

# ── Configuración ──────────────────────────────────────────────────────────────

YEARS = list(range(2018, 2027))

# Cuentas que extraemos de Análisis Resumen (código en col A)
# None = fila que se identifica por label en col B, no por código
CUENTAS_AR = {
    "1":  "Total Activos",
    "11": "Fondos Disponibles",
    "13": "Inversiones",
    "14": "Cartera de Créditos",
    "21": "Obligaciones con el Público",
    "2":  "Total Pasivos",
    # Total Patrimonio: sin código, se busca por label
}
LABEL_PATRIMONIO_AR = "Total Patrimonio"

# Cuentas del Balance (código exacto en col A)
CUENTAS_BAL = {"2101", "2103"}

# Filas de Clasificación de Cartera a incluir (código en col A)
CLF_INCLUDE = {
    # Totales generales
    "1", "53", "105", "157", "158", "159", "160",
    # Por vencer: subtotales por tipo
    "2", "8", "13", "18", "42", "47",
    # Vencida: subtotales por tipo
    "106", "112", "117", "122", "146", "151",
    # Vencida productivo: bandas temporales
    "107", "108", "109", "110", "111",
    # Vencida consumo
    "113", "114", "115", "116",
    # Vencida inmobiliario
    "118", "119", "120", "121",
    # Vencida microcrédito
    "123", "124", "125", "126", "127",
    # Vencida VIP
    "147", "148", "149", "150",
    # Vencida educativo
    "152", "153", "154", "155", "156",
}


# ── Normalización de nombres ───────────────────────────────────────────────────

def normalize_coop(name: str) -> str:
    """
    Normaliza el nombre de una cooperativa para unificar variantes
    LTDA / LIMITADA que la SEPS usó en distintos años.
    Canonical: siempre termina en ' LTDA' (más corto, más común).
    """
    if not name:
        return name
    n = name.strip()
    # Orden importa: reemplazar primero la forma larga
    if n.endswith(" LIMITADA"):
        n = n[: -len(" LIMITADA")] + " LTDA"
    elif n.endswith(" LTDA."):
        n = n[: -len(" LTDA.")] + " LTDA"
    # Caso especial: 'FERNANDO DAQUILEMA' sin sufijo
    # (en años posteriores aparece como 'FERNANDO DAQUILEMA LIMITADA')
    # No agregamos sufijo si no tenía ninguno → queda sin sufijo para unificar
    return n


# ── Extractor genérico ─────────────────────────────────────────────────────────

def parse_ym(val):
    """Convierte un valor de fecha de Excel a string 'YYYY-MM'."""
    if val is None:
        return None
    if isinstance(val, (datetime, date)):
        return val.strftime("%Y-%m")
    s = str(val)
    if len(s) >= 7:
        return s[:7]
    return None


def build_col_map(date_row, coop_row, start_col):
    """Construye {col_idx: (coop_name, 'YYYY-MM')} para las columnas de datos."""
    col_map = {}
    for ci in range(start_col, len(date_row)):
        d = date_row[ci]
        c = coop_row[ci]
        if d is None or c is None:
            continue
        ym = parse_ym(d)
        if ym and isinstance(c, str) and c.strip():
            col_map[ci] = (normalize_coop(c.strip()), ym)
    return col_map


def extract_rows(ws, date_row_idx, coop_row_idx, data_start_col,
                 label_fn, data_start_row_idx, include_set=None):
    """
    Extrae datos de una hoja wide (cooperativas en columnas).

    Args:
        ws: worksheet openpyxl
        date_row_idx: índice 0-based de la fila con fechas
        coop_row_idx: índice 0-based de la fila con nombres de coops
        data_start_col: índice 0-based de la primera columna de datos
        label_fn: función(row) -> str|None que extrae el label de la fila
        data_start_row_idx: índice 0-based de la primera fila de datos
        include_set: si se provee, solo extrae labels en este set

    Returns:
        {label: {coop: {ym: float}}}
    """
    rows = list(ws.iter_rows(values_only=True))

    date_row = rows[date_row_idx] if date_row_idx < len(rows) else []
    coop_row = rows[coop_row_idx] if coop_row_idx < len(rows) else []
    col_map = build_col_map(date_row, coop_row, data_start_col)

    result = {}
    for ri in range(data_start_row_idx, len(rows)):
        row = rows[ri]
        label = label_fn(row)
        if not label:
            continue
        if include_set is not None and label not in include_set:
            continue

        for ci, (coop, ym) in col_map.items():
            if ci >= len(row):
                continue
            val = row[ci]
            if val is None or val == "":
                continue
            try:
                v = float(val)
            except (TypeError, ValueError):
                continue

            if label not in result:
                result[label] = {}
            if coop not in result[label]:
                result[label][coop] = {}
            result[label][coop][ym] = v

    return result


# ── Extractores por hoja ───────────────────────────────────────────────────────

def extract_indicadores(wb):
    """
    Hoja 'Indicadores': fechas en R1, coops en R2, datos desde R4.
    Labels en col A (índice 0). Sin código numérico separado.
    """
    ws = wb["Indicadores"]

    def label_fn(row):
        a = row[0] if row else None
        if not a or not isinstance(a, str):
            return None
        a = a.strip()
        if not a:
            return None
        # Detectar si la fila tiene al menos un valor numérico en columnas de datos
        has_data = any(isinstance(v, (int, float)) for v in row[2:10])
        return a if has_data else None

    return extract_rows(
        ws,
        date_row_idx=0,
        coop_row_idx=1,
        data_start_col=2,
        label_fn=label_fn,
        data_start_row_idx=3,
    )


def extract_analisis_resumen(wb):
    """
    Hoja 'Análisis Resumen': fechas en R1, coops en R2, datos desde R3.
    Col A = código, Col B = nombre.
    Solo extrae las cuentas definidas en CUENTAS_AR.
    También extrae 'Total Patrimonio' por label en col B.
    """
    ws = wb["Análisis Resumen"]
    target_codes = set(CUENTAS_AR.keys())

    def label_fn(row):
        a = str(row[0]).strip() if row and row[0] is not None else ""
        b = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ""

        if a in target_codes:
            return CUENTAS_AR[a]
        if b == LABEL_PATRIMONIO_AR:
            return LABEL_PATRIMONIO_AR
        return None

    return extract_rows(
        ws,
        date_row_idx=0,
        coop_row_idx=1,
        data_start_col=3,
        label_fn=label_fn,
        data_start_row_idx=2,
    )


def extract_balance(wb):
    """
    Hoja 'Balance': fechas en R1, coops en R2, datos desde R4.
    Col A = código de cuenta (hasta 6 dígitos), Col B = nombre.
    Solo extrae las cuentas en CUENTAS_BAL.
    """
    ws = wb["Balance"]

    def label_fn(row):
        a = str(row[0]).strip() if row and row[0] is not None else ""
        b = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ""
        if a in CUENTAS_BAL:
            return b or a
        return None

    return extract_rows(
        ws,
        date_row_idx=0,
        coop_row_idx=1,
        data_start_col=3,
        label_fn=label_fn,
        data_start_row_idx=3,
    )


def extract_resultados(wb):
    """
    Hoja 'Resultados': fechas en R1 (col E+), coops en R2 (col E+), datos desde R4.
    Col C (índice 2) = nombre del concepto.
    """
    ws = wb["Resultados"]

    def label_fn(row):
        c = row[2] if len(row) > 2 and row[2] is not None else None
        if not c or not isinstance(c, str):
            return None
        c = c.strip()
        return c if c else None

    return extract_rows(
        ws,
        date_row_idx=0,
        coop_row_idx=1,
        data_start_col=4,
        label_fn=label_fn,
        data_start_row_idx=3,
    )


def extract_clasificacion_cartera(wb):
    """
    Hoja 'Clasificación de Cartera': fechas en R1, coops en R2, datos desde R4.
    Col A = código, Col B = nombre. Solo incluye filas en CLF_INCLUDE.
    """
    ws = wb["Clasificación de Cartera"]

    def label_fn(row):
        a = str(row[0]).strip() if row and row[0] is not None else ""
        b = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ""
        if a in CLF_INCLUDE:
            return b or a
        return None

    return extract_rows(
        ws,
        date_row_idx=0,
        coop_row_idx=1,
        data_start_col=3,
        label_fn=label_fn,
        data_start_row_idx=3,
        include_set=None,  # el filtro lo hace label_fn directamente
    )


def extract_patrimonio_tecnico(wb):
    """
    Hoja 'Patrimonio Técnico': fechas en R1, coops en R2, datos desde R4.
    Col A = código (100-150), Col B = nombre.
    """
    ws = wb["Patrimonio Técnico"]

    def label_fn(row):
        a = str(row[0]).strip() if row and row[0] is not None else ""
        b = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ""
        # Las 6 filas de PT tienen códigos 100, 110, 120, 130, 140, 150
        if a in {"100", "110", "120", "130", "140", "150"}:
            return b or a
        return None

    return extract_rows(
        ws,
        date_row_idx=0,
        coop_row_idx=1,
        data_start_col=3,
        label_fn=label_fn,
        data_start_row_idx=3,
    )


# ── Merge y serialización ──────────────────────────────────────────────────────

def build_rename_map(all_sections: dict) -> dict:
    """
    Detecta pares (nombre_sin_sufijo, nombre_sin_sufijo + ' LTDA') en el
    universo completo de cooperativas y devuelve un mapa de renombrado.
    Ejemplo: {'FERNANDO DAQUILEMA': 'FERNANDO DAQUILEMA LTDA'}
    """
    all_coops = set()
    for section in all_sections.values():
        for coop_dict in section.values():
            all_coops.update(coop_dict.keys())

    rename = {}
    for name in all_coops:
        canonical = name + " LTDA"
        if canonical in all_coops:
            rename[name] = canonical
    return rename


def apply_rename(all_sections: dict, rename_map: dict) -> dict:
    """Aplica el mapa de renombrado a todas las secciones."""
    for section in all_sections.values():
        for label, coop_dict in section.items():
            for old_name, new_name in rename_map.items():
                if old_name in coop_dict:
                    if new_name not in coop_dict:
                        coop_dict[new_name] = {}
                    # Solo copia meses que no existen ya en el canónico
                    for ym, val in coop_dict[old_name].items():
                        if ym not in coop_dict[new_name]:
                            coop_dict[new_name][ym] = val
                    del coop_dict[old_name]
    return all_sections


def merge_data(all_data):
    """
    Combina todos los dicts {label: {coop: {ym: value}}} del mismo tipo
    provenientes de múltiples años en uno solo.
    """
    merged = {}
    for year_data in all_data:
        for label, coop_dict in year_data.items():
            if label not in merged:
                merged[label] = {}
            for coop, ym_dict in coop_dict.items():
                if coop not in merged[label]:
                    merged[label][coop] = {}
                merged[label][coop].update(ym_dict)
    return merged


def to_indexed(merged, all_coops, all_months):
    """
    Convierte {label: {coop: {ym: value}}} a arrays indexados para JSON compacto.
    Formato: {label: [[v_c0_m0, v_c0_m1, ...], [v_c1_m0, ...], ...]}
    None donde no hay dato.
    """
    result = {}
    coop_idx = {c: i for i, c in enumerate(all_coops)}
    month_idx = {m: i for i, m in enumerate(all_months)}

    for label, coop_dict in merged.items():
        matrix = [[None] * len(all_months) for _ in range(len(all_coops))]
        for coop, ym_dict in coop_dict.items():
            ci = coop_idx.get(coop)
            if ci is None:
                continue
            for ym, val in ym_dict.items():
                mi = month_idx.get(ym)
                if mi is None:
                    continue
                # Redondear a 8 cifras significativas para reducir tamaño
                matrix[ci][mi] = round(val, 8) if isinstance(val, float) else val
        result[label] = matrix

    return result


def round_compact(v):
    """Redondea para JSON compacto: 4 decimales para ratios, 2 para montos."""
    if v is None:
        return None
    if abs(v) < 100:          # probablemente un ratio o índice
        return round(v, 6)
    return round(v, 2)        # monto en miles


# ── Pipeline principal ─────────────────────────────────────────────────────────

def run(source_dir: Path, output_file: Path):
    print(f"Fuente: {source_dir}")
    print(f"Salida: {output_file}")

    # Acumuladores por tipo de hoja
    acc = {
        "ind": [], "ar": [], "bal": [],
        "res": [], "clf": [], "pat": [],
    }

    for year in YEARS:
        fname = source_dir / f"Coop Ind Financieros {year} sin f.xlsx"
        if not fname.exists():
            print(f"  [{year}] OMITIDO — archivo no encontrado: {fname.name}")
            continue

        print(f"  [{year}] Leyendo...", end=" ", flush=True)
        wb = openpyxl.load_workbook(fname, read_only=True, data_only=True)

        acc["ind"].append(extract_indicadores(wb))
        acc["ar"].append(extract_analisis_resumen(wb))
        acc["bal"].append(extract_balance(wb))
        acc["res"].append(extract_resultados(wb))
        acc["clf"].append(extract_clasificacion_cartera(wb))
        acc["pat"].append(extract_patrimonio_tecnico(wb))

        wb.close()
        print("OK")

    # Merge por tipo, luego fusionar variantes de nombre entre secciones
    print("\nCombinando datos...", flush=True)
    merged = {k: merge_data(v) for k, v in acc.items()}
    rename_map = build_rename_map(merged)
    if rename_map:
        print(f"  Fusionando variantes: {list(rename_map.keys())}")
        merged = apply_rename(merged, rename_map)

    # Construir listas maestras de coops y meses
    all_coops_set = set()
    all_months_set = set()
    for section in merged.values():
        for coop_dict in section.values():
            for coop, ym_dict in coop_dict.items():
                all_coops_set.add(coop)
                all_months_set.update(ym_dict.keys())

    all_coops = sorted(all_coops_set)
    all_months = sorted(all_months_set)

    print(f"  Cooperativas únicas: {len(all_coops)}")
    print(f"  Meses: {all_months[0]} → {all_months[-1]} ({len(all_months)} meses)")
    for k, v in merged.items():
        print(f"  [{k}] {len(v)} series")

    # Convertir a arrays indexados
    print("\nIndexando...", flush=True)
    indexed = {k: to_indexed(v, all_coops, all_months) for k, v in merged.items()}

    # Aplicar redondeo compacto
    for section in indexed.values():
        for label, matrix in section.items():
            for ci, row in enumerate(matrix):
                matrix[ci] = [round_compact(v) for v in row]

    # JSON final
    output = {
        "meta": {
            "generated": datetime.now().strftime("%Y-%m-%d"),
            "months": all_months,
            "coops": all_coops,
        },
        "ind": indexed["ind"],
        "ar":  indexed["ar"],
        "bal": indexed["bal"],
        "res": indexed["res"],
        "clf": indexed["clf"],
        "pat": indexed["pat"],
    }

    # Escribir
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, separators=(",", ":"))

    size_mb = output_file.stat().st_size / 1_048_576
    print(f"\nJSON generado: {output_file}")
    print(f"Tamaño: {size_mb:.2f} MB")
    print("\nListo.")


# ── Entry point ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="ETL Monitor Cooperativas Ecuador")
    parser.add_argument(
        "--source",
        type=Path,
        default=Path(__file__).parent.parent.parent / "Monitor-Finanzas",
        help="Carpeta con los archivos Excel (default: ../Monitor-Finanzas relativo al repo)",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path(__file__).parent.parent / "docs" / "data" / "cooperativas.json",
        help="Ruta del JSON de salida",
    )
    args = parser.parse_args()
    run(args.source, args.output)
