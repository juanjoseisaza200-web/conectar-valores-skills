# -*- coding: utf-8 -*-
"""scan_dashboard.py — FASE 0: escanea el modelo hoja por hoja para construir el mapa de fuentes.
Lista hojas; por cada hoja clave dumpea etiqueta(colE o colB)+F(escalar)+Q(2026)+AD(2039) de cada
fila con etiqueta; detecta selector de escenario, Check ESF, unidades.

Uso:
    python scan_dashboard.py "modelo.xlsm"
    python scan_dashboard.py "modelo.xlsm" --hojas EEFF,Deuda,WK --rows 1-600
CERO alucinacion: lo que imprime es lo que hay; usa esto para anclar cada metrica a su celda.
"""
import sys, os, argparse
sys.path.insert(0, os.path.expanduser(r"~\.claude\skills\modelaje-conectar\tools"))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import cv_model as cv

KEY_SHEETS=["EEFF","EEFF Hist.","Deuda","WK","Waterfall","Ingresos","OPEX","CAPEX & PPE","TX","Inputs_C","Inputs_Years"]

def fmt(x):
    try: return f"{float(x):,.3f}"
    except: return "" if x is None else str(x)[:30]

def dump_sheet(W, name, r1, r2, labcols=(5,2)):
    if name not in W:
        print(f"  (hoja '{name}' no existe)"); return
    sh=W[name]
    print(f"\n===== {name} (filas {r1}-{r2}) =====")
    for r in range(r1, r2+1):
        lab=None
        for lc in labcols:
            v=sh.Cells(r,lc).Value
            if v not in (None,""): lab=v; break
        if lab is None: continue
        F=sh.Cells(r,6).Value; Q=sh.Cells(r,17).Value; AD=sh.Cells(r,30).Value
        # solo filas con algun numero (señal de metrica)
        if not any(isinstance(x,(int,float)) for x in (F,Q,AD)): continue
        print(f"  r{r:<4} | {str(lab)[:46]:<46} | F={fmt(F):>14} Q={fmt(Q):>16} AD={fmt(AD):>16}")

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument("modelo")
    ap.add_argument("--hojas", default=None, help="lista separada por comas; default = hojas clave")
    ap.add_argument("--rows", default="1-600")
    a=ap.parse_args()
    r1,r2=[int(x) for x in a.rows.split("-")]
    hojas=a.hojas.split(",") if a.hojas else KEY_SHEETS
    xl=wb=None
    try:
        xl,wb,W=cv.abrir(a.modelo, read_only=True, kill=False); cv.calc_wait(xl, full=True)
        print("HOJAS:", list(W.keys()))
        # escenario / unidades / ESF
        try:
            sel=W["Inputs_C"].Range("F6").Value
            print("\nSelector escenario Inputs_C!F6 =", sel, "| I5(MATCH) =", W["Inputs_C"].Range("I5").Value)
        except: print("\n(no se hallo Inputs_C!F6)")
        for name in hojas:
            if name in W:
                dump_sheet(W, name, r1, r2)
        print("\nNOTA: unidades tipicas = COP miles -> dashboard COP mil MM = modelo/1e6. Confirmar con un KPI conocido.")
        print("Check ESF: buscar fila 'Check ESF' en EEFF (suma |.| en 21 cols debe ser 0).")
    finally:
        try:
            if wb: wb.Close(False)
            if xl: xl.Quit()
        except: pass
    print("DONE")

if __name__=="__main__":
    main()
