# -*- coding: utf-8 -*-
"""MOCK TEST de la skill modelaje-conectar sobre un SANDBOX (copia descartable del v3).
Test A (detección): sembrar 4 errores deliberados -> el auditor DEBE atraparlos todos.
Test B (construcción): crear una fila nueva con cv.copy_row_format + cv.fast_fill -> uniforme y limpia.
Test C (cierre): invocar cv.cerrar_modelo() (función refactorizada) -> checks = 0.
El sandbox se elimina al final. El modelo real NO se toca."""
import sys, os, shutil, time, subprocess, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.path.insert(0, os.path.expanduser(r"~\.claude\skills\modelaje-conectar\tools"))
import cv_model as cv

BASE = r"C:\Users\<usuario>\OneDrive - CONECTAR VALORES SAS\<carpeta del modelo>"
REAL = BASE + r"\<modelo real>.xlsm"
SBX = r"C:\Users\<usuario>\AppData\Local\Temp\cv_modelo\_MOCKTEST_SANDBOX.xlsm"
AUD = os.path.expanduser(r"~\.claude\skills\modelaje-conectar\tools\auditar_modelo.py")

resultados = []
def res(nombre, ok, detalle=""):
    resultados.append((nombre, ok, detalle))
    print(f"  [{'OK ' if ok else 'FALLO'}] {nombre}" + (f" -> {detalle}" if detalle else ""))

shutil.copy2(REAL, SBX)
print("Sandbox creado (copia descartable).")

# ============ SEMBRAR ERRORES + CONSTRUIR ============
xl, wb, W = cv.abrir(SBX)
wk = W["WK"]; tx = W["TX"]
try:
    # Error 1: romper FAST — fórmula distinta en una celda del rango
    wk.Range("U95").Formula = "=U64*0.123"
    # Error 2: etiqueta en rojo (formato indebido)
    wk.Cells(98, 5).Font.Color = 255
    # Error 3: valor pegado en columna observada de una fila de flujo
    wk.Range("M412").Formula = None; wk.Range("M412").Value = 999999
    # Error 4: error de cálculo
    tx.Range("Q227").Formula = "=1/0"

    # Test B: construir fila nueva con la librería (zona libre WK 760)
    cv.copy_row_format(wk, 110, 760)
    wk.Cells(760, 5).Value = "MOCKTEST fila de prueba"
    wk.Range("J760").Formula = "=J64*0.01*J$7"
    ok_ff = cv.fast_fill(wk, 760, desde="J")
    unif = cv.auditar_fast(wk, [760])
    res("B1 fast_fill ejecuta", ok_ff)
    res("B2 fila nueva uniforme R1C1 J:AD", not unif, str(unif))

    cv.calc_wait(xl, full=True)

    # ¿La librería detecta lo sembrado?
    res("A1 auditar_fast detecta FAST roto (WK!U95)", ("WK", 95) in cv.auditar_fast(wk, [95]))
    res("A2 scan_rojos detecta etiqueta roja (WK!98)", any("98" in h for h in cv.scan_rojos(wk, 98, 98)))
    res("A3 observados_cero detecta valor pegado (WK!M412)", any(r == 412 for _, r, _, _ in cv.observados_cero(wk, [412])))
    n_err, _ = cv.errores(wb)
    res("A4 errores() detecta #DIV/0 (TX!Q227)", n_err >= 1, f"errores={n_err}")
    wb.Save()
finally:
    wb.Close(False); xl.Quit()

# ============ AUDITOR CLI DEBE DAR FAIL ============
p = subprocess.run([sys.executable, "-X", "utf8", AUD, SBX,
                    "--filas", "WK:93-134,398-426", "--filas-cero", "WK:412-425"],
                   capture_output=True, text=True)
salida = (p.stdout or "") + (p.stderr or "")
res("A5 auditor CLI devuelve FAIL con errores sembrados", p.returncode != 0 and "FAIL" in salida)
for tag, frase in [("A5a detecta FAST", "5. FAST"), ("A5b detecta rojos", "7. Sin rojos"),
                   ("A5c detecta observados", "6. Observados"), ("A5d detecta celdas error", "3. Cero celdas")]:
    linea = next((l for l in salida.splitlines() if frase in l), "")
    res(tag, "[FAIL]" in linea, linea.strip()[:90])

# ============ TEST C: corregir lo sembrado y cerrar con cv.cerrar_modelo ============
xl, wb, W = cv.abrir(SBX)
wk = W["WK"]; tx = W["TX"]
try:
    wk.Range("T95").Copy(); wk.Range("U95").PasteSpecial(-4123); xl.CutCopyMode = False
    wk.Cells(98, 5).Font.Color = 0
    wk.Range("L412").Copy(); wk.Range("M412").PasteSpecial(-4123); xl.CutCopyMode = False
    tx.Range("Q227").Formula = None
    wk.Rows(760).Delete()  # quitar fila de prueba del Test B (es sandbox)
    cierre = cv.cerrar_modelo(xl, wb)   # función refactorizada — primera invocación real
    res("C1 cerrar_modelo() converge", abs(cierre["Check_Macros"] or 0) < 0.5 and abs(cierre["Check_TotalMacros"] or 0) < 0.5, str(cierre))
    esf = cv.check_esf(W["EEFF"])
    res("C2 ESF = 0 tras reparar y cerrar", all(abs(v or 0) < 1 for v in esf))
    n_err, _ = cv.errores(wb)
    res("C3 0 errores tras reparar", n_err == 0)
finally:
    wb.Close(False); xl.Quit()

os.remove(SBX)
print("\nSandbox eliminado.")
total = len(resultados); ok = sum(1 for _, o, _ in resultados if o)
print(f"\nRESULTADO MOCK TEST: {ok}/{total} pruebas superadas")
sys.exit(0 if ok == total else 1)
