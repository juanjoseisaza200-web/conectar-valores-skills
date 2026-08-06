# -*- coding: utf-8 -*-
"""Fase 4 desde cero: balance de prueba histórico (2019-2025) vs proyección v3, cuenta por cuenta.
Calcula empalme 2025->2026, CAGR hist vs proy, y flags de desfase. Genera anexo xlsx + resumen."""
import pandas as pd, csv, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

BASE = r"C:\Users\Administrador\OneDrive - CONECTAR VALORES SAS\Archivos de Ricardo Palacios Sarmiento - Modelos Interaseo\Mapa de Variables Interaseo SAS ESP"
TMP = r"C:\Users\Administrador\AppData\Local\Temp\interaseo_mv"

eeff = {}
with open(TMP + r"\eeff_v3_val.tsv", encoding="utf-8") as f:
    for row in csv.reader(f, delimiter="\t"):
        r = int(row[0])
        eeff[r] = row[1:]

def m(r, year):
    idx = 9 + (year - 2019)
    try: return float(eeff[r][idx])
    except Exception: return 0.0

def mserie(r, y1, y2):
    return [m(r, y) for y in range(y1, y2 + 1)]

df = pd.read_pickle(TMP + r"\consolidado\detalle_terceros.pkl")
yrs_h = [str(y) for y in range(2019, 2026)]
sub = df.groupby("cod_sub")[yrs_h].sum() / 1e3   # COP -> COP miles
cls = df.groupby("cod_clase2")[yrs_h].sum() / 1e3
cta = df.groupby("cod_cuenta")[yrs_h].sum() / 1e3

def h(code):
    if code in sub.index:
        return [-float(sub.loc[code, y]) for y in yrs_h]
    return [0.0] * 7

def cagr(serie):
    serie = [abs(v) for v in serie]
    nz = [(i, v) for i, v in enumerate(serie) if v > 1]
    if len(nz) < 2: return None
    (i0, v0), (i1, v1) = nz[0], nz[-1]
    if i1 == i0: return None
    return (v1 / v0) ** (1 / (i1 - i0)) - 1

MAPA = [
 ("Recolección (432307)", "432307", 291), ("Transporte (432308)", "432308", 292),
 ("Barrido y limpieza (432309)", "432309", 293), ("Transferencia (432310)", "432310", 294),
 ("Aprovechamiento (432311)", "432311", 295), ("Tratamiento (432312)", "432312", 296),
 ("Disposición final (432313)", "432313", 297), ("Limpieza y lavado (432315)", "432315", 299),
 ("Otros especiales (432316)", "432316", 300), ("Comercialización (432317)", "432317", 301),
 ("Comisiones (439008)", "439008", 303), ("Otros servicios (439090)", "439090", 305),
 ("Facilities (439091)", "439091", 306), ("Convenios prest. integral (439093)", "439093", 307),
 ("Devol. venta aseo (439516)", "439516", 309),
]
rows = []
for nombre, puc, fila in MAPA:
    sh = h(puc)
    p25, p26 = sh[-1], m(fila, 2026)
    proy = mserie(fila, 2026, 2039)
    c_h, c_p = cagr(sh), cagr(proy)
    emp = (p26 / p25 - 1) if abs(p25) > 1000 else None
    flags = []
    if emp is not None and abs(emp) > 0.10: flags.append(f"empalme {emp:+.0%}")
    if c_h is not None and c_p is not None and (c_p - c_h) > 0.03: flags.append(f"CAGR proy>{c_h:.0%}+300pb")
    rows.append({"Línea": nombre, "2019": round(sh[0]/1e6,1), "2025": round(p25/1e6,1),
                 "2026 mod": round(p26/1e6,1), "2039 mod": round(proy[-1]/1e6,1),
                 "Empalme": f"{emp:+.1%}" if emp is not None else "n/a",
                 "CAGR h": f"{c_h:.1%}" if c_h is not None else "n/a",
                 "CAGR p": f"{c_p:.1%}" if c_p is not None else "n/a",
                 "Flag": "; ".join(flags) if flags else "ok"})

def clase(c):
    return [float(cls.loc[c, y]) for y in yrs_h] if c in cls.index else [0.0]*7

ing_h = [-(a + b) for a, b in zip(clase("42"), clase("43"))]
costo_h = clase("75"); admin_h = clase("51")
ing_m = mserie(289, 2026, 2039)
costo_m = [abs(v) for v in mserie(315, 2026, 2039)]
admin_m = [abs(v) for v in mserie(332, 2026, 2039)]
agg = []
def fila_agg(nombre, hist, proy):
    hist = [abs(x) for x in hist]; proy = [abs(x) for x in proy]
    c_h, c_p = cagr(hist), cagr(proy)
    emp = proy[0]/hist[-1]-1 if hist[-1] else None
    agg.append({"Agregado": nombre, "2019": round(hist[0]/1e6,1), "2025": round(hist[-1]/1e6,1),
                "2026 mod": round(proy[0]/1e6,1), "2039 mod": round(proy[-1]/1e6,1),
                "Empalme": f"{emp:+.1%}" if emp is not None else "n/a",
                "CAGR h": f"{c_h:.1%}" if c_h else "n/a", "CAGR p": f"{c_p:.1%}" if c_p else "n/a"})
fila_agg("Ingresos operacionales (42+43)", ing_h, ing_m)
fila_agg("Costo de ventas (75)", costo_h, costo_m)
fila_agg("Gasto administración (51)", admin_h, admin_m)
marg_h = [(abs(i)-abs(c)-abs(a))/abs(i) for i, c, a in zip(ing_h, costo_h, admin_h)]
marg_m = [(i-c-a)/i for i, c, a in zip(ing_m, costo_m, admin_m)]
agg.append({"Agregado": "Margen operativo", "2019": f"{marg_h[0]:.1%}", "2025": f"{marg_h[-1]:.1%}",
            "2026 mod": f"{marg_m[0]:.1%}", "2039 mod": f"{marg_m[-1]:.1%}",
            "Empalme": f"{(marg_m[0]-marg_h[-1])*10000:+.0f} pb", "CAGR h": "", "CAGR p": ""})

bal = []
def gv(cuenta, y):
    return abs(float(cta.loc[cuenta, y])) if cuenta in cta.index else 0.0
dso = lambda saldo, ing: saldo/ing*365 if ing else 0
bal.append({"Métrica": "DSO cartera comercial bruta (días)",
            "2019": round(dso(gv("1408","2019"), abs(ing_h[0]))),
            "2025": round(dso(gv("1408","2025"), abs(ing_h[-1]))),
            "2026 mod": round(dso(m(30,2026), ing_m[0])),
            "2033 mod": round(dso(m(30,2033), m(289,2033))),
            "2039 mod": round(dso(m(30,2039), m(289,2039))),
            "Lectura": "La caída proyectada es mezcla contable del castigo, no recaudo"})
bal.append({"Métrica": "Provisión / cartera bruta",
            "2019": f"{gv('1480','2019')/gv('1408','2019'):.0%}",
            "2025": f"{gv('1480','2025')/gv('1408','2025'):.0%}",
            "2026 mod": f"{abs(m(67,2026))/m(30,2026):.0%}",
            "2033 mod": f"{abs(m(67,2033))/m(30,2033):.0%}",
            "2039 mod": f"{abs(m(67,2039))/m(30,2039):.0%}",
            "Lectura": "Converge a ~20% vs 35-67% histórico — supuesto de morosidad bajo evidencia"})
bal.append({"Métrica": "CxC partes relacionadas (1470, mil MM)",
            "2019": round(gv("1470","2019")/1e6,1), "2025": round(gv("1470","2025")/1e6,1),
            "2026 mod": round(m(59,2026)/1e6,1), "2033 mod": round(m(59,2033)/1e6,1),
            "2039 mod": round(m(59,2039)/1e6,1),
            "Lectura": "Creció 2019-2025 (préstamos grupo); modelo congela (subordinación Tx)"})
bal.append({"Métrica": "CxP totales clase 24 vs EEFF!158 (mil MM)",
            "2019": round(abs(float(cls.loc['24','2019']))/1e6,1) if '24' in cls.index else 0,
            "2025": round(abs(float(cls.loc['24','2025']))/1e6,1) if '24' in cls.index else 0,
            "2026 mod": round(m(158,2026)/1e6,1), "2033 mod": round(m(158,2033)/1e6,1),
            "2039 mod": round(m(158,2039)/1e6,1), "Lectura": "Crece con costos (rotaciones)"})

t1 = pd.DataFrame(rows); t2 = pd.DataFrame(agg); t3 = pd.DataFrame(bal)
out = BASE + r"\20260611_Auditoria_Anexo_HistVsProy.xlsx"
with pd.ExcelWriter(out, engine="openpyxl") as w:
    t1.to_excel(w, sheet_name="Ingresos por línea", index=False)
    t2.to_excel(w, sheet_name="Agregados P&G", index=False)
    t3.to_excel(w, sheet_name="Balance y ratios", index=False)
print("ANEXO:", out)
print("\n=== INGRESOS POR LÍNEA (libro 4323/4390 vs EEFF v3) — COP mil MM ===")
print(t1.to_string(index=False))
print("\n=== AGREGADOS P&G ===")
print(t2.to_string(index=False))
print("\n=== BALANCE Y RATIOS ===")
print(t3.to_string(index=False))
