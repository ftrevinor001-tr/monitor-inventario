#!/usr/bin/env python3
"""
actualizar_dashboard.py
=======================
Lee el Excel en data/datos.xlsx, procesa los datos y regenera index.html
con la información embebida lista para GitHub Pages.

Uso:
    python actualizar_dashboard.py
    python actualizar_dashboard.py --excel data/mi_archivo.xlsx
"""

import json, argparse, sys, calendar as cal_lib
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    print("ERROR: pandas no instalado. Ejecuta: pip install pandas openpyxl")
    sys.exit(1)

parser = argparse.ArgumentParser()
parser.add_argument("--excel",  default="data/datos.xlsx")
parser.add_argument("--output", default="index.html")
args = parser.parse_args()

EXCEL_PATH = Path(args.excel)
OUTPUT     = Path(args.output)
TEMPLATE   = Path("dashboard_template.html")

print(f"📂 Leyendo: {EXCEL_PATH}")
if not EXCEL_PATH.exists():
    print(f"ERROR: No se encontró {EXCEL_PATH}"); sys.exit(1)
if not TEMPLATE.exists():
    print(f"ERROR: No se encontró {TEMPLATE}"); sys.exit(1)

# ── Leer hoja correcta ────────────────────────────────────────────────────────
xl = pd.ExcelFile(EXCEL_PATH)
sheet = next(
    (s for s in xl.sheet_names if any(k in s.lower() for k in ["historico","histórico","faltante"])),
    xl.sheet_names[0]
)
print(f"📋 Hoja: '{sheet}'")
df = pd.read_excel(EXCEL_PATH, sheet_name=sheet)

# ── Detectar columnas ─────────────────────────────────────────────────────────
def find(hints):
    for c in df.columns:
        if any(h in c.lower() for h in hints):
            return c
    return None

col_fecha  = find(["día","dia","fecha"]) or df.columns[0]
col_clave  = find(["clave"])
col_nombre = find(["nombre"])
col_marca  = find(["marca"])
col_comp   = find(["comprador"])
col_stat   = find(["clasificacion 2","clasificación 2"])
cl_cols    = [c for c in df.columns if "clasificaci" in c.lower()
              and "2" not in c and "punto" not in c.lower()]
col_cl     = next((c for c in cl_cols if "rotaci" not in c.lower()), None)
col_cr     = next((c for c in cl_cols if "rotaci" in c.lower()), None)

if not col_clave or not col_stat:
    print("ERROR: Columnas 'Clave' o 'Clasificacion 2' no encontradas"); sys.exit(1)

# ── Preparar datos ────────────────────────────────────────────────────────────
EXCLUIR = ["NUEVOS"]
if col_cr:
    antes = len(df)
    df = df[~df[col_cr].isin(EXCLUIR)]
    print(f"🚫 Excluidas {antes-len(df)} filas NUEVOS")

df["fecha_str"] = pd.to_datetime(df[col_fecha]).dt.strftime("%Y-%m-%d")
df["mes"]       = pd.to_datetime(df[col_fecha]).dt.strftime("%Y-%m")

# Meses de inventario: solo del último día disponible
last_date = df["fecha_str"].max()
col_mi = next((c for c in df.columns if "meses" in c.lower() and "inventario" in c.lower()), None)
mi_map = {}
if col_mi:
    mi_map = (df[df["fecha_str"] == last_date]
              .groupby(col_clave)[col_mi]
              .first()
              .dropna()
              .to_dict())
    print(f"📦 Meses inventario: {len(mi_map)} claves al {last_date}")

# ── Algoritmo de rachas ───────────────────────────────────────────────────────
def calc_streaks(f_dates, r_dates):
    """
    Retorna:
      gir  : primer día del RIESGO justo antes del último FALTANTE
      glr  : último día de ese RIESGO
      gif  : primer día del último FALTANTE
      glf  : último día del último FALTANTE (None = sigue activo)
      drfa : días en ese RIESGO previo
      dfa  : días en ese último FALTANTE
    """
    if not f_dates:
        return None, None, None, None, 0, 0

    all_ev = sorted([(d,"F") for d in set(f_dates)] + [(d,"R") for d in set(r_dates)])
    i = len(all_ev) - 1

    while i >= 0 and all_ev[i][1] == "R": i -= 1        # saltar riesgo al final

    glf = all_ev[i][0] if i >= 0 else None              # último día faltante
    dfa = 0
    while i >= 0 and all_ev[i][1] == "F": dfa += 1; i -= 1
    gif = all_ev[i+1][0] if dfa > 0 else None           # primer día faltante

    glr = all_ev[i][0] if i >= 0 and all_ev[i][1]=="R" else None
    drfa = 0
    while i >= 0 and all_ev[i][1] == "R": drfa += 1; i -= 1
    gir = all_ev[i+1][0] if drfa > 0 else None          # primer día riesgo previo

    return gir, glr, gif, glf, drfa, dfa

# ── Construir registros ───────────────────────────────────────────────────────
print("⚙️  Procesando claves...")
meta = df.groupby(col_clave).agg(
    Nombre=(col_nombre, "first") if col_nombre else (col_clave,"first"),
    Marca =(col_marca,  "first") if col_marca  else (col_clave,"first"),
    Comp  =(col_comp,   "first") if col_comp   else (col_clave,"first"),
    CL    =(col_cl,     "first") if col_cl     else (col_clave,"first"),
    CR    =(col_cr,     "first") if col_cr     else (col_clave,"first"),
).reset_index()

records = []
for _, row in meta.iterrows():
    clave   = row[col_clave]
    sub     = df[df[col_clave] == clave]
    months  = {}

    for mes, msub in sub.groupby("mes"):
        f_d = sorted(msub[msub[col_stat]=="FALTANTE"]["fecha_str"].unique())
        r_d = sorted(msub[msub[col_stat]=="RIESGO"]["fecha_str"].unique())
        f, r = len(f_d), len(r_d)
        if f > 0 or r > 0:
            months[mes] = {
                "f":  f,  "r":  r,
                "if": f_d[0]  if f_d else None,   # inicio faltante en el mes
                "lf": f_d[-1] if f_d else None,   # fin faltante en el mes
                "ir": r_d[0]  if r_d else None,   # inicio riesgo en el mes
                "lr": r_d[-1] if r_d else None,   # fin riesgo en el mes
            }
    if not months:
        continue

    # Rachas globales
    f_dates = sorted(sub[sub[col_stat]=="FALTANTE"]["fecha_str"].unique())
    r_dates = sorted(sub[sub[col_stat]=="RIESGO"]["fecha_str"].unique())
    gir, glr, gif, glf, drfa, dfa = calc_streaks(f_dates, r_dates)

    # Meses de inventario del último día
    mi_val = mi_map.get(clave, None)
    mi = round(float(mi_val), 2) if mi_val is not None and str(mi_val) != 'nan' else None

    rec = {
        "c":    str(clave),
        "mo":   months,
        "gir":  gir,   "glr": glr,
        "gif":  gif,   "glf": glf,
        "drfa": drfa,  "dfa": dfa,
        "mi":   mi,
    }
    if col_nombre: rec["n"]  = str(row.get("Nombre",""))[:60]
    if col_marca:  rec["m"]  = str(row.get("Marca", ""))
    if col_comp:   rec["cp"] = str(row.get("Comp",  ""))
    if col_cl:     rec["cl"] = str(row.get("CL",    "")).strip()
    if col_cr:     rec["cr"] = str(row.get("CR",    "")).strip()
    records.append(rec)

# ── Estadísticas ──────────────────────────────────────────────────────────────
all_dates   = sorted(df["fecha_str"].unique())
total_days  = int(df["fecha_str"].nunique())
max_days_db = {k: int(v) for k,v in df.groupby("mes")["fecha_str"].nunique().items()}
cal_days    = {m: cal_lib.monthrange(int(m[:4]), int(m[5:]))[1] for m in df["mes"].unique()}

MONTH_LABELS = {
    f"20{y:02d}-{m:02d}": lbl
    for y in range(25, 28)
    for m, lbl in enumerate(["Ene","Feb","Mar","Abr","May","Jun",
                              "Jul","Ago","Sep","Oct","Nov","Dic"], 1)
    for lbl in [f"{lbl} 20{y:02d}"]
}

data_out = {
    "data":        records,
    "compradores": sorted(df[col_comp].dropna().unique().tolist())  if col_comp else [],
    "marcas":      sorted(df[col_marca].dropna().unique().tolist()) if col_marca else [],
    "clasifLetra": sorted(df[col_cl].dropna().unique().tolist())    if col_cl else [],
    "clasifRot":   sorted(df[col_cr].dropna().unique().tolist())    if col_cr else [],
    "meses":       sorted(df["mes"].unique().tolist()),
    "calDays":     {k: int(v) for k,v in cal_days.items()},
    "totalDays":   total_days,
    "firstDate":   all_dates[0],
    "lastDate":    all_dates[-1],
    "miDate":      last_date,
}

print(f"✅ {len(records)} claves | {total_days} días | {all_dates[0]} → {all_dates[-1]}")

# ── Generar HTML ──────────────────────────────────────────────────────────────
print(f"💾 Generando {OUTPUT}...")
template = TEMPLATE.read_text(encoding="utf-8")
html = template
html = html.replace("__DB__",      json.dumps(data_out, ensure_ascii=False, separators=(",",":")))
html = html.replace("__MAXDAYS__", json.dumps(max_days_db, ensure_ascii=False))
html = html.replace("__LABELS__",  json.dumps(MONTH_LABELS, ensure_ascii=False))
html = html.replace("__CALDAYS__", json.dumps({k:int(v) for k,v in cal_days.items()}, ensure_ascii=False))
OUTPUT.write_text(html, encoding="utf-8")
print(f"✅ {OUTPUT} listo ({OUTPUT.stat().st_size//1024} KB)")
