#!/usr/bin/env python3
"""
actualizar_dashboard.py
=======================
Lee el Excel en data/datos.xlsx, procesa los datos y regenera index.html
con la información embebida lista para GitHub Pages.

Uso:
    python actualizar_dashboard.py
    python actualizar_dashboard.py --excel data/mi_archivo.xlsx

El script también es ejecutado automáticamente por GitHub Actions
cada vez que se sube un nuevo archivo .xlsx a la carpeta data/.
"""

import json
import argparse
import calendar as cal_lib
import sys
from pathlib import Path

# ── Dependencias ──────────────────────────────────────────────────────────────
try:
    import pandas as pd
except ImportError:
    print("ERROR: pandas no está instalado. Ejecuta: pip install pandas openpyxl")
    sys.exit(1)

# ── Argumentos ────────────────────────────────────────────────────────────────
parser = argparse.ArgumentParser(description="Actualiza el dashboard de inventario")
parser.add_argument("--excel",  default="data/datos.xlsx", help="Ruta al archivo Excel")
parser.add_argument("--output", default="index.html",      help="Archivo HTML de salida")
args = parser.parse_args()

EXCEL_PATH  = Path(args.excel)
OUTPUT_PATH = Path(args.output)
TEMPLATE    = Path("dashboard_template.html")

print(f"📂 Leyendo Excel: {EXCEL_PATH}")
if not EXCEL_PATH.exists():
    print(f"ERROR: No se encontró el archivo {EXCEL_PATH}")
    sys.exit(1)

# ── Leer Excel ────────────────────────────────────────────────────────────────
xl = pd.ExcelFile(EXCEL_PATH)
sheet = next(
    (s for s in xl.sheet_names if any(k in s.lower() for k in ["historico", "histórico", "faltante"])),
    xl.sheet_names[0],
)
print(f"📋 Usando hoja: '{sheet}'")
df = pd.read_excel(EXCEL_PATH, sheet_name=sheet)

# Detectar columnas de forma flexible
def find_col(df, hints):
    for col in df.columns:
        if any(h in col.lower() for h in hints):
            return col
    return None

col_fecha  = find_col(df, ["día", "dia", "fecha"]) or df.columns[0]
col_clave  = find_col(df, ["clave"])
col_nombre = find_col(df, ["nombre"])
col_marca  = find_col(df, ["marca"])
col_comp   = find_col(df, ["comprador"])
col_stat   = find_col(df, ["clasificacion 2", "clasificación 2"])
col_cl     = next((c for c in df.columns if "clasificaci" in c.lower()
                   and "2" not in c and "punto" not in c.lower()
                   and "rotaci" not in c.lower()), None)
col_cr     = next((c for c in df.columns if "rotaci" in c.lower()
                   or ("clasificaci" in c.lower() and "2" not in c
                       and "punto" not in c.lower() and c != col_cl)), None)

missing = [n for n, c in [("clave", col_clave), ("clasificacion 2", col_stat)] if not c]
if missing:
    print(f"ERROR: No se encontraron columnas: {missing}")
    sys.exit(1)

# ── Filtrar NUEVOS ────────────────────────────────────────────────────────────
EXCLUIR = ["NUEVOS"]
if col_cr:
    antes = len(df)
    df = df[~df[col_cr].isin(EXCLUIR)]
    print(f"🚫 Excluidas {antes - len(df)} filas con clasificación NUEVOS")

# ── Preparar fechas ───────────────────────────────────────────────────────────
df["fecha_str"] = pd.to_datetime(df[col_fecha]).dt.strftime("%Y-%m-%d")
df["mes"]       = pd.to_datetime(df[col_fecha]).dt.strftime("%Y-%m")

# ── Meta por clave ────────────────────────────────────────────────────────────
agg = {
    "Nombre": (col_nombre, "first") if col_nombre else None,
    "Marca":  (col_marca,  "first") if col_marca  else None,
    "Comp":   (col_comp,   "first") if col_comp   else None,
    "CL":     (col_cl,     "first") if col_cl     else None,
    "CR":     (col_cr,     "first") if col_cr     else None,
}
agg = {k: v for k, v in agg.items() if v}
meta = df.groupby(col_clave).agg(**{k: v for k, v in agg.items()}).reset_index()

# ── Calcular días por clave/mes/status ───────────────────────────────────────
print("⚙️  Calculando días faltante y riesgo...")
records = []
for _, row in meta.iterrows():
    clave = row[col_clave]
    sub   = df[df[col_clave] == clave]
    months = {}
    for mes, msub in sub.groupby("mes"):
        f = int(msub[msub[col_stat] == "FALTANTE"]["fecha_str"].nunique())
        r = int(msub[msub[col_stat] == "RIESGO"]["fecha_str"].nunique())
        if f > 0 or r > 0:
            months[mes] = {"f": f, "r": r}
    if not months:
        continue
    rec = {"c": str(clave), "mo": months}
    if col_nombre: rec["n"] = str(row.get("Nombre", ""))[:60]
    if col_marca:  rec["m"] = str(row.get("Marca",  ""))
    if col_comp:   rec["cp"]= str(row.get("Comp",   ""))
    if col_cl:     rec["cl"]= str(row.get("CL",     "")).strip()
    if col_cr:     rec["cr"]= str(row.get("CR",     "")).strip()
    records.append(rec)

# ── Estadísticas del período ──────────────────────────────────────────────────
all_dates   = sorted(df["fecha_str"].unique())
total_days  = df["fecha_str"].nunique()
max_days_db = {k: int(v) for k, v in df.groupby("mes")["fecha_str"].nunique().items()}
cal_days    = {m: cal_lib.monthrange(int(m[:4]), int(m[5:]))[1] for m in df["mes"].unique()}

data_out = {
    "data":        records,
    "compradores": sorted(df[col_comp].dropna().unique().tolist()) if col_comp else [],
    "marcas":      sorted(df[col_marca].dropna().unique().tolist()) if col_marca else [],
    "clasifLetra": sorted(df[col_cl].dropna().unique().tolist()) if col_cl else [],
    "clasifRot":   sorted(df[col_cr].dropna().unique().tolist()) if col_cr else [],
    "meses":       sorted(df["mes"].unique().tolist()),
    "calDays":     {k: int(v) for k, v in cal_days.items()},
    "totalDays":   total_days,
}

print(f"✅ {len(records)} claves procesadas | {total_days} días en BD")
print(f"   Período: {all_dates[0]} → {all_dates[-1]}")
print(f"   Meses: {data_out['meses']}")

# ── Inyectar en el template HTML ──────────────────────────────────────────────
print(f"💾 Generando {OUTPUT_PATH}...")
if not TEMPLATE.exists():
    print(f"ERROR: No se encontró {TEMPLATE}. Asegúrate de que esté en el repo.")
    sys.exit(1)

template = TEMPLATE.read_text(encoding="utf-8")

MONTH_LABELS = {
    "2025-01":"Ene 2025","2025-02":"Feb 2025","2025-03":"Mar 2025",
    "2025-04":"Abr 2025","2025-05":"May 2025","2025-06":"Jun 2025",
    "2025-07":"Jul 2025","2025-08":"Ago 2025","2025-09":"Sep 2025",
    "2025-10":"Oct 2025","2025-11":"Nov 2025","2025-12":"Dic 2025",
    "2026-01":"Ene 2026","2026-02":"Feb 2026","2026-03":"Mar 2026",
    "2026-04":"Abr 2026","2026-05":"May 2026","2026-06":"Jun 2026",
    "2026-07":"Jul 2026","2026-08":"Ago 2026","2026-09":"Sep 2026",
    "2026-10":"Oct 2026","2026-11":"Nov 2026","2026-12":"Dic 2026",
}

html = template
html = html.replace("__DB__",      json.dumps(data_out, ensure_ascii=False, separators=(",", ":")))
html = html.replace("__MAXDAYS__", json.dumps(max_days_db, ensure_ascii=False))
html = html.replace("__LABELS__",  json.dumps(MONTH_LABELS, ensure_ascii=False))
html = html.replace("__CALDAYS__", json.dumps({k: int(v) for k, v in cal_days.items()}, ensure_ascii=False))

OUTPUT_PATH.write_text(html, encoding="utf-8")
size_kb = OUTPUT_PATH.stat().st_size // 1024
print(f"✅ {OUTPUT_PATH} generado ({size_kb} KB)")
print("🚀 Listo para publicar en GitHub Pages")
