# -*- coding: utf-8 -*-
r"""Auditor automático del estándar Conectar — checklist ejecutable.
USO:  python auditar_modelo.py "ruta\modelo.xlsm"
        [--filas "WK:93-134;TX:204-227"]        <- filas nuevas/modificadas: FAST + rojos
        [--filas-cero "WK:412-425"]              <- subconjunto de FLUJO: ademas observados=0
        [--esf-fila 285] [--cerrar]
Notas: filas ANCLADAS (bases fiscales, neteos, saldos con historico) llevan valores
observados POR DISENO -> van en --filas pero NO en --filas-cero.
Devuelve PASS/FAIL por item y exit code 0 solo si TODO pasa.
El modelo que ejecuta la skill NO puede declarar terminado un cambio sin PASS de este auditor."""
import sys, os, argparse, json

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import cv_model as cv


def parse_filas(s):
    out = {}
    if not s:
        return out
    for parte in s.split(";"):
        if ":" not in parte:
            continue
        hoja, rangos = parte.split(":", 1)
        filas = []
        for rg in rangos.split(","):
            if "-" in rg:
                a, b = rg.split("-")
                filas += list(range(int(a), int(b) + 1))
            elif rg.strip():
                filas.append(int(rg))
        out[hoja.strip()] = filas
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("ruta")
    ap.add_argument("--filas", default="", help='Filas nuevas/modificadas, ej "WK:93-134,398-426;TX:204-227"')
    ap.add_argument("--filas-cero", default="", help='Filas de FLUJO que deben dar 0 en observados')
    ap.add_argument("--esf-fila", type=int, default=285)
    ap.add_argument("--cerrar", action="store_true", help="Ejecutar el cierre (CopyBloque) antes de auditar")
    args = ap.parse_args()

    filas = parse_filas(args.filas)
    resultados = []
    ok_total = True

    def chk(nombre, ok, detalle=""):
        nonlocal ok_total
        ok_total = ok_total and ok
        resultados.append((nombre, "PASS" if ok else "FAIL", detalle))

    xl, wb, W = cv.abrir(args.ruta)
    try:
        cv.calc_wait(xl, full=True)
        cv.calc_wait(xl)
        if args.cerrar:
            cierre = cv.cerrar_modelo(xl, wb)
            chk("0. Cierre ejecutado", abs(cierre["Check_Macros"] or 0) < 0.5, json.dumps(cierre, default=str))

        # 1. ESF
        esf = cv.check_esf(W["EEFF"], args.esf_fila)
        chk("1. Check ESF = 0 (J:AD)", all(abs(v or 0) < 1 for v in esf), str([round(v or 0) for v in esf]))

        # 2. Checks de macro
        try:
            val = lambda n: wb.Names(n).RefersToRange.Value
            m1, m2, m3 = val("Check_Macros"), val("Check_TotalMacros"), val("Check_Repayment")
            chk("2. Checks macro = 0", all(abs(x or 0) < 0.5 for x in (m1, m2, m3)), f"Macros={m1} Total={m2} Repay={m3}")
        except Exception as e:
            chk("2. Checks macro", False, f"nombres no encontrados: {e}")

        # 3. Errores
        n_err, det = cv.errores(wb)
        chk("3. Cero celdas con error", n_err == 0, str(det[:5]))

        # 4. Circulares
        circ = cv.circulares(wb)
        chk("4. Cero referencias circulares", not circ, str(circ[:5]))

        # 5 y 7. Filas declaradas: FAST + rojos (SOLO sobre las filas listadas, no el rango completo)
        for hoja, rs in filas.items():
            sh = W[hoja]
            malas = cv.auditar_fast(sh, rs)
            chk(f"5. FAST uniforme R1C1 [{hoja}]", not malas, str(malas[:8]))
            rojos = []
            for r in rs:
                rojos += cv.scan_rojos(sh, r, r)
            chk(f"7. Sin rojos indebidos [{hoja}]", not rojos, str(rojos[:8]))
        # 6. Filas de FLUJO: observados = 0
        for hoja, rs in parse_filas(args.filas_cero).items():
            sh = W[hoja]
            viol = cv.observados_cero(sh, rs)
            chk(f"6. Observados = 0 (flujo) [{hoja}]", not viol, str(viol[:8]))

        print("\n" + "=" * 78)
        print(f"AUDITORIA ESTANDAR CONECTAR — {os.path.basename(args.ruta)}")
        print("=" * 78)
        for n, s, d in resultados:
            print(f"  [{s}] {n}" + (f"  -> {d}" if (s == "FAIL" and d) else ""))
        print("-" * 78)
        print("RESULTADO:", "PASS — modelo conforme" if ok_total else "FAIL — NO entregar; corregir y re-auditar")
        if args.cerrar and ok_total:
            wb.Save()
            print("(guardado tras cierre)")
    finally:
        wb.Close(False) if not (args.cerrar and ok_total) else None
        xl.Quit()
    sys.exit(0 if ok_total else 1)


if __name__ == "__main__":
    main()
