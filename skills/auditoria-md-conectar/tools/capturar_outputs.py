# -*- coding: utf-8 -*-
r"""Captura de outputs de un modelo (SOLO LECTURA) para auditoría: DSCR reportado vs RECALCULADO
con servicio total, covenants, liquidez en meses de OPEX, checks.
USO: python capturar_outputs.py "ruta\modelo.xlsm"
NOTA: las filas están parametrizadas para un modelo de referencia — EDITAR el bloque FILAS para otro modelo.
Patrón clave reutilizable: nunca confiar en el ratio del modelo; reconstruir numerador y denominador
desde los componentes y comparar (la diferencia es hallazgo)."""
import sys, os, time, json
# resuelve la librería de modelaje-conectar relativa a esta skill (portable entre usuarios)
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "modelaje-conectar", "tools"))
import cv_model as cv

RUTA = sys.argv[1]
FILAS = {
    "dscr_reportado": ("Deuda", 556), "dscr_min": ("Deuda", "F557"), "cumplimiento": ("Deuda", "F558"),
    "numerador_covenant": ("Waterfall", 156),
    "sd_actual": ("Deuda", 549), "sd_lp": ("Deuda", 550),
    "sd_cp_amort": ("Deuda", 450), "sd_cp_int": ("Deuda", 454),
    "dn_ebitda": ("Dashboard", 49), "liquidez_meses_opex": ("Dashboard", 60),
    "dso": ("Dashboard", 61), "caja_eeff": ("EEFF", 17), "dsra": ("Deuda", 416), "ira": ("Deuda", 396),
}
COLS = ["Q","R","S","T","U","V","W","X","Y","Z","AA","AB","AC","AD"]
YRS = list(range(2026, 2040))

xl, wb, W = cv.abrir(RUTA, read_only=True)
try:
    cv.calc_wait(xl, full=True)
    out = {}
    def serie(hoja, fila):
        sh = W[hoja]
        return {y: (sh.Range(f"{c}{fila}").Value or 0) for y, c in zip(YRS, COLS)}
    for k, (hoja, fila) in FILAS.items():
        if isinstance(fila, str):  # celda única tipo "F557"
            out[k] = W[hoja].Range(fila).Value
        else:
            out[k] = serie(hoja, fila)
    # DSCR recalculado con servicio TOTAL
    num = out["numerador_covenant"]
    dscr_recalc = {}
    for y in YRS:
        sd = abs((out["sd_actual"][y] or 0) + (out["sd_lp"][y] or 0)
                 + (out["sd_cp_amort"][y] or 0) + (out["sd_cp_int"][y] or 0))
        dscr_recalc[y] = round(num[y] / sd, 2) if sd else None
    out["DSCR_RECALCULADO_servicio_total"] = dscr_recalc
    out["DSCR_reportado_vs_recalc"] = {
        y: (round(out["dscr_reportado"][y], 2), dscr_recalc[y]) for y in YRS
    }
    # checks
    out["checks"] = {}
    chk = W.get("Checks")
    if chk:
        for r in range(1, 16):
            vals = [chk.Cells(r, c).Value for c in range(1, 8)]
            vals = [v for v in vals if v not in (None, "")]
            if vals: out["checks"][r] = vals
    print(json.dumps(out, indent=1, default=str))
finally:
    wb.Close(False)
    xl.Quit()
