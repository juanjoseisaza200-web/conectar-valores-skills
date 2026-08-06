# -*- coding: utf-8 -*-
"""test_skill.py — autotest de dash_lib: ejerce los builders via COM sobre una copia temporal
(zona scratch), verifica que construyen sin error, y DESCARTA (no guarda). kill=False.
Uso: python test_skill.py "cualquier_dashboard_o_modelo.xlsm"
"""
import sys, os, shutil, tempfile
sys.path.insert(0, os.path.dirname(__file__))
import dash_lib as dl

def main():
    src=sys.argv[1] if len(sys.argv)>1 else None
    if not src or not os.path.exists(src):
        print("Uso: python test_skill.py <archivo.xlsm>"); return
    tmp=os.path.join(tempfile.gettempdir(), "dl_selftest.xlsm")
    if os.path.exists(tmp): os.remove(tmp)
    shutil.copy2(src, tmp)
    xl=wb=None; ok=True
    try:
        xl,wb,W=dl.cv.abrir(tmp, kill=False); dl.cv.calc_wait(xl, full=True)
        # hoja scratch
        sh = W["Dashboard"] if "Dashboard" in W else wb.Worksheets(1)
        hv = W["_DashViz"] if "_DashViz" in W else sh
        R=460  # zona scratch lejana
        n0=sh.ChartObjects().Count
        # 1) banda + sub
        dl.band(sh, R, "TEST BANDA")
        dl.band_sub(sh, R+1, "test sub-header")
        # 2) años + tabla compacta (espejo de una fila cualquiera con datos: usar la propia fila 7 de años)
        dl.years_header(sh, R+2, src_year_row=7)
        # mirror de una metrica: si hay EEFF usa ingresos, si no, una constante
        if "EEFF" in W:
            dl.compact_mirror(sh, R+3, "Ingresos test", "EEFF", 289, scale=1e6)
        else:
            for j in range(14): sh.Cells(R+3,3+j).Value=j+1
            sh.Cells(R+3,2).Value="serie test"
        # 3) chart que lee la compacta
        ch=dl.make_chart(sh, 3, R+5, 12, R+18, dl.COLCLUST,
                         [("Serie", f"=Dashboard!$C${R+3}:$P${R+3}", dl.NAVY)],
                         "Chart test", legend=False, xr=f"=Dashboard!$C${R+2}:$P${R+2}")
        # 4) kpi band (2 tarjetas)
        dl.kpi_band(sh, R+20, [("KPI A","=1.23","0.00",""),("KPI B","=4.56","0.00","sub")], ncols=4)
        # 5) waterfall bridge minimal
        dl.waterfall_bridge(sh, hv, 14, R+24, 24, R+36,
                            [("A","0","=10","0"),("B","=4","0","=6"),("C","0","=4","0")], "Cascada test")
        # 6) anclar + capturar/aplicar colores
        nmoved=dl.anchor_all_charts(sh)
        ref=dl.capture_colors(sh); applied,miss=dl.apply_colors(sh, ref)
        dl.cv.calc_wait(xl, full=True)
        n1=sh.ChartObjects().Count
        circ=dl.cv.circulares(wb); tot,_=dl.cv.errores(wb)
        print(f"charts {n0}->{n1} (+{n1-n0})  anclados={nmoved}  colores aplicados={applied}/miss={miss}")
        print(f"circulares={circ}  errores={tot}")
        ok = (n1-n0)>=2 and not circ and tot==0 and miss==0
        print("RESULTADO:", "PASS" if ok else "REVISAR")
    except Exception as e:
        print("EXCEPCION:", repr(e)); ok=False; print("RESULTADO: FAIL")
    finally:
        try:
            if wb: wb.Close(SaveChanges=False)   # DESCARTA
            if xl: xl.Quit()
        except: pass
        try:
            if os.path.exists(tmp): os.remove(tmp)
        except: pass
    print("DONE")

if __name__=="__main__":
    main()
