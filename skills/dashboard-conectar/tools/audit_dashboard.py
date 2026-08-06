# -*- coding: utf-8 -*-
"""audit_dashboard.py — FASE 5: auditor PASS/FAIL de un dashboard.
Verifica (un dashboard es PURA PRESENTACION):
  - KPIs del modelo IDENTICOS entre el dashboard y la base (no se rompio el modelo)
  - ESF=0 en 21 cols, 0 errores, 0 circulares
  - 0 sparklines (el usuario los detesta)
  - 0 series en color default del tema (serie sin colorear)
  - 0 solapes de charts
  - filtro de periodo en rango completo por defecto (si existe)
Opcional: --diff-colors otra.xlsm  (compara color de cada serie vs un archivo de referencia)

Uso:
    python audit_dashboard.py "dashboard vN.xlsm" --base "modelo.xlsm"
    python audit_dashboard.py "dashboard vN.xlsm" --base "modelo.xlsm" --kpis "EBITDA=Waterfall!60,17;Util=EEFF!452,17"
    python audit_dashboard.py "dashboard vN.xlsm" --diff-colors "dashboard vM.xlsm"
"""
import sys, os, argparse
sys.path.insert(0, os.path.expanduser(r"~\.claude\skills\modelaje-conectar\tools"))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import cv_model as cv
import dash_lib as dl

DEFAULT_KPIS={  # nombre -> (hoja, fila, col)
 "EBITDA_2026":("Waterfall",60,17), "UtilNeta_2026":("EEFF",452,17), "Patrim_2026":("EEFF",241,17),
 "Caja_2026":("EEFF",17,17), "DeudaBruta_2026":("Deuda",540,17),
}
def read_kpis(W, spec):
    out={}
    for k,(sh,r,c) in spec.items():
        try: out[k]=W[sh].Cells(r,c).Value
        except: out[k]=None
    return out

def parse_kpis(s):
    spec={}
    for part in s.split(";"):
        if not part.strip(): continue
        nm, ref = part.split("="); sh, rc = ref.split("!"); r,c = rc.split(",")
        spec[nm.strip()]=(sh.strip(), int(r), int(c))
    return spec

def chart_series_colors(W):
    d=W["Dashboard"]; out={}
    for ci,co in enumerate(d.ChartObjects(),1):
        ch=co.Chart
        for i in range(1, ch.SeriesCollection().Count+1):
            s=ch.SeriesCollection(i)
            try: nm=str(s.Name)
            except: nm=f"s{i}"
            try: fc=int(s.Format.Fill.ForeColor.RGB)
            except: fc=None
            try: lc=int(s.Format.Line.ForeColor.RGB)
            except: lc=None
            out[(ci,nm)]=(fc,lc)
    return out

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument("dashboard")
    ap.add_argument("--base", default=None)
    ap.add_argument("--kpis", default=None)
    ap.add_argument("--diff-colors", dest="diffcolors", default=None)
    ap.add_argument("--esf-row", type=int, default=285)
    a=ap.parse_args()
    DIR=os.path.dirname(os.path.abspath(a.dashboard))
    spec=parse_kpis(a.kpis) if a.kpis else DEFAULT_KPIS

    # modo diff-colors
    if a.diffcolors:
        xl,wb,W=cv.abrir(a.dashboard, read_only=True, kill=False); cv.calc_wait(xl,full=True)
        cur=chart_series_colors(W); wb.Close(False); xl.Quit()
        xl,wb,W=cv.abrir(a.diffcolors, read_only=True, kill=False); cv.calc_wait(xl,full=True)
        ref=chart_series_colors(W); wb.Close(False); xl.Quit()
        diffs=0
        for k in sorted(set(cur)&set(ref)):
            if cur[k]!=ref[k]:
                diffs+=1; print(f"  DIFF {k}: {ref[k]} -> {cur[k]}")
        print("TOTAL diffs de color:", diffs)
        print("RESULTADO:", "PASS" if diffs==0 else "REVISAR"); return

    # snapshot base
    base=None
    if a.base:
        xlb,wbb,Wb=cv.abrir(a.base, read_only=True, kill=False); cv.calc_wait(xlb,full=True)
        base=read_kpis(Wb, spec); wbb.Close(False); xlb.Quit()

    xl,wb,W=cv.abrir(a.dashboard, kill=False); cv.calc_wait(xl,full=True)
    cierre=cv.cerrar_modelo(xl,wb); print("CIERRE:", cierre)
    cv.calc_wait(xl,full=True)
    d=W["Dashboard"]

    ok=True
    if base is not None:
        cur=read_kpis(W, spec)
        print("\n=== KPIs base vs dashboard (deben ser IDENTICOS) ===")
        for k in base:
            b=base[k] or 0; v=cur.get(k) or 0; good=abs(b-v)<max(1.0,abs(b)*1e-6); ok=ok and good
            print(f"  {k:<16} base={b:>16.4f} dash={v:>16.4f} {'OK' if good else 'DIFF'}")

    esf=cv.check_esf(W["EEFF"], a.esf_row) if "EEFF" in W else [0]
    esfok=all(abs(x or 0)<1 for x in esf); ok=ok and esfok
    circ=cv.circulares(wb); ok=ok and not circ
    tot,_=cv.errores(wb); ok=ok and tot==0
    # sparklines
    spk=0
    try: spk=d.Cells.SparklineGroups.Count
    except: pass
    ok = ok and spk==0
    # series sin colorear (tema default)
    themed=[]
    for ci,co in enumerate(d.ChartObjects(),1):
        ch=co.Chart
        for i in range(1,ch.SeriesCollection().Count+1):
            try: fc=int(ch.SeriesCollection(i).Format.Fill.ForeColor.RGB)
            except: fc=None
            if fc in dl.THEME_DEFAULTS: themed.append((ci,i,fc))
    ok = ok and not themed
    # solapes
    chs=[]
    for co in d.ChartObjects():
        chs.append((round(co.Left),round(co.Top),round(co.Width),round(co.Height)))
    ov=0
    for i in range(len(chs)):
        for j in range(i+1,len(chs)):
            ax,ay,aw,ah=chs[i]; bx,by,bw,bh=chs[j]
            if ax<bx+bw and bx<ax+aw and ay<by+bh and by<ay+ah: ov+=1
    ok = ok and ov==0
    # filtro default
    filtro=""
    try:
        st=W["_DashViz"].Range("A52").Value; cnt=W["_DashViz"].Range("A53").Value
        nyears=int(d.Cells(7,30).Value)-int(d.Cells(7,17).Value)+1
        filtro=f"START={st} CNT={cnt} (completo={nyears})";
        if st not in (1,1.0) or cnt!=nyears: ok=False
    except: filtro="(sin filtro)"

    print("\n=== CHECKLIST ===")
    print("ESF=0 (21 cols):", esfok)
    print("Circulares:", circ if circ else "[]")
    print("Errores celdas:", tot)
    print("Sparklines:", spk, "(debe 0)")
    print("Series en color default del tema:", themed if themed else "ninguna")
    print("Solapes de charts:", ov)
    print("Filtro de periodo:", filtro)
    print("Charts:", d.ChartObjects().Count)
    wb.Save()
    print("\nRESULTADO:", "PASS" if ok else "REVISAR")
    wb.Close(False); xl.Quit(); print("DONE")

if __name__=="__main__":
    main()
