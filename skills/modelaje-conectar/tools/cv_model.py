# -*- coding: utf-8 -*-
"""Librería de modelaje Conectar (estándar FAST) — código probado en producción jun-2026.
Uso: import cv_model as cv  (añadir esta carpeta a sys.path).
Todas las funciones asumen pywin32. Constantes COM incluidas."""
import os, time, json

import win32com.client

XL_MANUAL = -4135
XL_AUTO = -4105
PASTE_FORMULAS = -4123
PASTE_FORMATS = -4122
PASTE_VALUES = -4163
END_DOWN = -4121
END_RIGHT = -4161
CELLTYPE_FORMULAS = -4123
ERRORS_FLAG = 16
AZUL = 0xFF0000   # BGR -> RGB(0,0,255)
NEGRO = 0x000000
ROJO_RGB = 255    # Font.Color de una celda roja
COLS_JAD = ["J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA","AB","AC","AD"]


def abrir(ruta, visible=False, read_only=False, kill=True):
    """Instancia limpia de Excel + workbook. Devuelve (xl, wb, dict_hojas)."""
    if kill:
        os.system('taskkill /f /im EXCEL.EXE >nul 2>&1')
        time.sleep(2)
    xl = win32com.client.DispatchEx("Excel.Application")  # DispatchEx: NO entra al ROT, no usar GetActiveObject después
    xl.Visible = visible
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    xl.AskToUpdateLinks = False
    xl.ScreenUpdating = False
    wb = xl.Workbooks.Open(ruta, UpdateLinks=0, ReadOnly=read_only)
    xl.Calculation = XL_MANUAL
    return xl, wb, {s.Name: s for s in wb.Worksheets}


def calc_wait(xl, full=False, timeout=300):
    """Calcula y ESPERA a que termine. OBLIGATORIO antes de leer cualquier celda/check."""
    if full:
        xl.CalculateFullRebuild()
    else:
        xl.Calculate()
    t0 = time.time()
    while True:
        try:
            if xl.CalculationState == 0:  # xlDone
                break
        except Exception:
            break
        if time.time() - t0 > timeout:
            print("AVISO: timeout de calculo")
            break
        time.sleep(0.5)


def circulares(wb):
    """Lista de referencias circulares reales (ignora los $A$1 benignos con CELL)."""
    out = []
    for sh in wb.Worksheets:
        try:
            cr = sh.CircularReference
            if cr is not None and cr.Address != "$A$1":
                out.append((sh.Name, cr.Address))
        except Exception:
            pass
    return out


def errores(wb):
    """(total, detalle) de celdas con error en todo el libro."""
    total, detalle = 0, []
    for sh in wb.Worksheets:
        try:
            e = sh.UsedRange.SpecialCells(CELLTYPE_FORMULAS, ERRORS_FLAG)
            total += e.Count
            detalle.append((sh.Name, e.Address[:140]))
        except Exception:
            pass
    return total, detalle


def cerrar_modelo(xl, wb, max_iter=40):
    """Réplica de la macro CopyBloque (cierre de circularidad + sizing) para el ESCENARIO ACTIVO.
    Requiere nombres: Scenario, Variables, Macros_Copy/Paste, Repayment_Copy/Paste,
    Check_Macros, Check_TotalMacros, Check_Repayment, Run_sizing (workbook o en hoja Macros).
    Devuelve dict con los checks finales."""
    mac = wb.Worksheets("Macros")
    nm = lambda n: wb.Names(n).RefersToRange
    val = lambda n: nm(n).Value
    esc = int(val("Scenario")); nvar = int(val("Variables"))
    rs = None
    for g in (lambda: wb.Names("Run_sizing").RefersToRange, lambda: mac.Range("Run_sizing")):
        try:
            rs = g(); rs.Value = 1; break
        except Exception:
            pass
    xl.Calculation = XL_AUTO
    calc_wait(xl, full=True)
    calc_wait(xl)
    it = 0
    while (abs(val("Check_TotalMacros") or 0) > 0 or abs(val("Check_Macros") or 0) > 0) and it < max_iter:
        ii = 0
        while abs(val("Check_Macros") or 0) > 0 and ii < 60:
            a = nm("Macros_Copy"); c0 = mac.Cells(a.Row + 1, a.Column + 1)
            src = mac.Range(c0, c0.End(END_RIGHT)); src = mac.Range(src, src.End(END_DOWN))
            b = nm("Macros_Paste"); dst = mac.Cells(b.Row + 2 * esc + nvar * (esc - 1), b.Column + 1)
            src.Copy(); dst.PasteSpecial(PASTE_VALUES); xl.CutCopyMode = False
            calc_wait(xl); ii += 1
        rc = nm("Repayment_Copy"); c0 = mac.Cells(rc.Row + 1, rc.Column + 1)
        src = mac.Range(c0, c0.End(END_DOWN)); src = mac.Range(src, src.End(END_RIGHT))
        rp = nm("Repayment_Paste"); dst = mac.Cells(rp.Row + 2 * esc + 10 * (esc - 1), rc.Column + 1)
        src.Copy(); dst.PasteSpecial(PASTE_VALUES); xl.CutCopyMode = False
        calc_wait(xl); it += 1
    if rs is not None:
        rs.Value = 0
    calc_wait(xl); calc_wait(xl)
    return {"iteraciones": it, "Check_Macros": val("Check_Macros"),
            "Check_TotalMacros": val("Check_TotalMacros"), "Check_Repayment": val("Check_Repayment")}


def fast_fill(sh, fila, desde="Q"):
    """Propaga la fórmula de {desde}{fila} idéntica (relativa) a J:AD. Garantiza FAST."""
    c = sh.Range(f"{desde}{fila}")
    if not str(c.Formula).startswith("="):
        return False
    c.Copy()
    sh.Range(f"J{fila}:AD{fila}").PasteSpecial(PASTE_FORMULAS)
    sh.Application.CutCopyMode = False
    return True


def auditar_fast(sh, filas):
    """Devuelve filas cuyo FormulaR1C1 NO es uniforme J:AD (deben ser cero)."""
    malas = []
    for r in filas:
        s = set()
        for c in range(10, 31):
            try:
                s.add(sh.Cells(r, c).FormulaR1C1)
            except Exception:
                s.add("?")
        if len(s) > 1:
            malas.append((sh.Name, r))
    return malas


def observados_cero(sh, filas, cols=("J", "M", "P"), tol=0.01):
    """Filas nuevas deben dar 0 en años observados. Devuelve violaciones."""
    out = []
    for r in filas:
        for c in cols:
            v = sh.Range(f"{c}{r}").Value
            if v not in (None, 0) and abs(v or 0) > tol:
                out.append((sh.Name, r, c, v))
    return out


def scan_rojos(sh, r1, r2, cmax=31):
    """Celdas con fuente roja en un bloque (solo válido para celdas exportadas)."""
    hits = []
    for r in range(r1, r2 + 1):
        for c in list(range(2, 8)) + list(range(10, cmax)):
            try:
                cell = sh.Cells(r, c)
                if cell.Font.Color == ROJO_RGB and (cell.Value is not None or cell.HasFormula):
                    hits.append(str(cell.Address))
            except Exception:
                pass
    return hits


def copy_row_format(sh, fila_origen, fila_destino):
    sh.Rows(fila_origen).Copy()
    sh.Rows(fila_destino).PasteSpecial(PASTE_FORMATS)
    sh.Application.CutCopyMode = False


def check_esf(eeff, fila_check=285):
    """Valores del check de balance J:AD. Todos deben ser ~0."""
    return [eeff.Range(f"{c}{fila_check}").Value for c in COLS_JAD]


def lbl(sh, r):
    for c in (5, 4, 3, 2, 1):
        v = sh.Cells(r, c).Value
        if v not in (None, ""):
            return str(v)
    return ""


def assert_lbl(sh, r, contiene):
    if contiene.lower() not in lbl(sh, r).lower():
        raise AssertionError(f"Ancla fallida {sh.Name}!r{r}: '{lbl(sh, r)}' no contiene '{contiene}'")
