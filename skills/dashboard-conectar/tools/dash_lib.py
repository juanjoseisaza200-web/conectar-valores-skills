# -*- coding: utf-8 -*-
"""dash_lib — libreria de construccion de dashboards (estandar Conectar / one-pager).
Codigo PROBADO en produccion (v2..v11). Depende de cv_model (modelaje-conectar).
NO reescribir COM a mano: usar estos helpers. Convenciones en reference/*.md.

Uso:
    import sys, os
    sys.path.insert(0, os.path.expanduser(r"~\\.claude\\skills\\dashboard-conectar\\tools"))
    import dash_lib as dl
    xl, wb, W = dl.cv.abrir(ruta, kill=False)   # NUNCA kill=True
"""
import sys, os, re, time
# import cv_model: LOCAL primero (copia incluida -> 100% portable), fallback a modelaje-conectar
sys.path.insert(0, os.path.expanduser(r"~\.claude\skills\modelaje-conectar\tools"))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import cv_model as cv   # abrir, calc_wait, cerrar_modelo, circulares, errores, check_esf, COLS_JAD

# ---------------- PALETA (valores .RGB de COM = BGR empaquetado) ----------------
NAVY=6174487; TEAL=8882990; GOLD=4826313; MINT=13434828; RED=2960895
SLATE=11507320; LGRAY=13157310; DGRAY=6710886; PURP=9123550; WHITE=16777215; GRID=15066597
THEME_DEFAULTS={8544277, 3305961, 2386713, 13999631, 9644960}  # colores del tema = serie SIN colorear
def rgb(r,g,b): return int(b)*65536+int(g)*256+int(r)   # construir color corporativo nuevo
PALETTE=[NAVY,GOLD,TEAL,SLATE,LGRAY]   # orden por defecto de series

# ---------------- tipos / constantes ----------------
COLCLUST=51; COLSTACK=52; LINE=4; AREASTACK=76; BARCLUST=57
CEN=-4108; LEFT_=-4131; VC=-4108; XLMOVE=2; XLVALIDATE_LIST=3
COLS_JAD=cv.COLS_JAD
def mcol(j): return COLS_JAD[7+j]   # j 0..13 -> Q..AD (años de proyeccion en el modelo)
def dcol(j): return chr(67+j)       # j 0..13 -> C..P (tabla compacta del dashboard)

# ---------------- busqueda por etiqueta (NUNCA por numero fijo) ----------------
def fB(sh, t, r1=1, r2=400):
    """Primera fila cuya col B contiene el texto t (case-insensitive)."""
    t=t.lower()
    for r in range(r1, r2+1):
        v=sh.Cells(r,2).Value
        if v and t in str(v).lower(): return r
    return None
def find_colE(sh, t, r1=1, r2=600, col=5):
    t=t.lower()
    for r in range(r1, r2+1):
        v=sh.Cells(r,col).Value
        if v and t in str(v).lower(): return r
    return None

# ---------------- bandas / sub-headers ----------------
def band(sh, row, text, c1=2, c2=30, fs=12):
    sh.Range(sh.Cells(row,c1), sh.Cells(row,c2)).Merge()
    c=sh.Cells(row,c1); c.Value=text; c.Font.Name="Arial"; c.Font.Bold=True; c.Font.Size=fs
    c.Font.Color=WHITE; c.Interior.Color=NAVY; c.HorizontalAlignment=LEFT_; c.VerticalAlignment=VC
    sh.Rows(row).RowHeight=20
def band_sub(sh, row, text, c2=12):
    sh.Range(sh.Cells(row,2), sh.Cells(row,c2)).Merge()
    c=sh.Cells(row,2); c.Value=text; c.Font.Bold=True; c.Font.Size=10; c.Font.Color=NAVY; c.Font.Name="Arial"

# ---------------- años + tabla compacta ----------------
def years_header(sh, row=7, src_year_row=7):
    """Escribe la cabecera de años compacta C:P = =Q$row..=AD$row (si la fila de años global esta en Q:AD)."""
    for j in range(14):
        cc=sh.Cells(row, 3+j); cc.Formula=f"=${mcol(j)}${src_year_row}"
def compact_mirror(sh, row, label, sheet, mrow, sign=1, scale=1e6, fmt='#,##0.0', fill=WHITE, bold=False):
    """Fila compacta C:P = +/- sheet!{Q..AD}{mrow} / scale  (la etiqueta en col B)."""
    lc=sh.Cells(row,2); lc.Value=label; lc.Font.Name="Arial"; lc.Font.Size=8; lc.Font.Color=NAVY; lc.Font.Bold=bold
    sh.Range(sh.Cells(row,2), sh.Cells(row,16)).Interior.Color=fill
    pre="-" if sign<0 else ""
    for j in range(14):
        cc=sh.Cells(row, 3+j)
        cc.Formula=(f"={pre}{sheet}!{mcol(j)}{mrow}/{scale}" if scale!=1 else f"={pre}{sheet}!{mcol(j)}{mrow}")
        cc.NumberFormat=fmt; cc.Font.Size=8; cc.Font.Color=NAVY; cc.HorizontalAlignment=CEN
        cc.Interior.Color=fill; cc.Font.Bold=bold

# ---------------- tarjetas KPI ----------------
def kpi_band(sh, start_row, cards, ncols=4, navy_title=True):
    """cards = lista de (titulo, formula, fmt, sub). 2 filas x ncols. Celdas fusionadas por tarjeta.
    OJO: titulo NO debe empezar con '=' (#NAME?). El numero (formula) si es formula."""
    width=30//ncols  # columnas por tarjeta sobre B:AD
    per_row=ncols
    for i,(titulo, formula, fmt, sub) in enumerate(cards):
        rr=start_row + (i//per_row)*3
        c1=2 + (i % per_row)*width; c2=c1+width-1
        # titulo
        sh.Range(sh.Cells(rr,c1), sh.Cells(rr,c2)).Merge()
        t=sh.Cells(rr,c1); t.Value=str(titulo); t.Font.Name="Arial"; t.Font.Bold=True; t.Font.Size=8
        t.Font.Color=WHITE; t.Interior.Color=NAVY; t.HorizontalAlignment=CEN
        # numero
        sh.Range(sh.Cells(rr+1,c1), sh.Cells(rr+1,c2)).Merge()
        n=sh.Cells(rr+1,c1)
        if isinstance(formula,str) and formula.startswith("="): n.Formula=formula
        else: n.Value=formula
        n.NumberFormat=fmt; n.Font.Name="Arial"; n.Font.Bold=True; n.Font.Size=18; n.Font.Color=NAVY; n.HorizontalAlignment=CEN
        # sub
        sh.Range(sh.Cells(rr+2,c1), sh.Cells(rr+2,c2)).Merge()
        s=sh.Cells(rr+2,c1); s.Value=str(sub) if sub is not None else ""; s.Font.Size=7; s.Font.Color=DGRAY; s.HorizontalAlignment=CEN

# ---------------- charts ----------------
def style_chart(ch, title, legend=True, axis2_fmt=None):
    try: ch.ChartArea.Format.Line.Visible=False; ch.ChartArea.Format.Fill.ForeColor.RGB=WHITE; ch.ChartArea.Format.Fill.Visible=True
    except: pass
    try: ch.ChartArea.Font.Name="Arial"; ch.ChartArea.Font.Size=9; ch.ChartArea.Font.Color=DGRAY
    except: pass
    try: ch.HasTitle=True; ch.ChartTitle.Text=title; ch.ChartTitle.Font.Name="Arial"; ch.ChartTitle.Font.Size=11; ch.ChartTitle.Font.Bold=True; ch.ChartTitle.Font.Color=NAVY
    except: pass
    try:
        ch.HasLegend=legend
        if legend: ch.Legend.Position=-4107; ch.Legend.Font.Size=8
    except: pass
    try:
        ch.Axes(2).MajorGridlines.Border.Color=GRID; ch.Axes(2).TickLabels.Font.Size=8
        if axis2_fmt: ch.Axes(2).TickLabels.NumberFormat=axis2_fmt; ch.Axes(2).TickLabels.NumberFormatLinked=False
    except: pass
    try: ch.Axes(1).TickLabels.Font.Size=8
    except: pass
    try: ch.PlotArea.Format.Fill.Visible=False
    except: pass
def place_chart(co, sh, c1, r1, c2, r2):
    co.Left=sh.Columns(c1).Left+2; co.Top=sh.Rows(r1).Top+1
    co.Width=sh.Columns(c2).Left - sh.Columns(c1).Left - 4
    co.Height=sh.Rows(r2).Top - sh.Rows(r1).Top - 2
    try: co.Placement=XLMOVE
    except: pass
def add_series(ch, specs, xr="=Dashboard!$Q$7:$AD$7"):
    """specs = lista de (nombre, values_ref, color, *opt). opt: 'line2' (linea eje 2), 'line' (linea eje 1)."""
    while ch.SeriesCollection().Count>0: ch.SeriesCollection(1).Delete()
    for spec in specs:
        nm, vr, col = spec[0], spec[1], spec[2]; opt = spec[3] if len(spec)>3 else None
        s=ch.SeriesCollection().NewSeries(); s.Name=nm; s.Values=vr; s.XValues=xr
        try: s.Format.Fill.Solid(); s.Format.Fill.ForeColor.RGB=col; s.Format.Line.ForeColor.RGB=col
        except: pass
        if opt=="line2":
            try: s.ChartType=LINE; s.AxisGroup=2; s.Format.Line.Weight=2.0
            except: pass
        elif opt=="line":
            try: s.ChartType=LINE; s.Format.Line.Weight=2.0
            except: pass
def make_chart(sh, c1, r1, c2, r2, ctype, specs, title, legend=True, xr="=Dashboard!$Q$7:$AD$7", axis2_fmt=None):
    co=sh.ChartObjects().Add(sh.Columns(c1).Left+2, sh.Rows(r1).Top+1, 200, 120)
    place_chart(co, sh, c1, r1, c2, r2)
    ch=co.Chart; ch.ChartType=ctype; add_series(ch, specs, xr); style_chart(ch, title, legend, axis2_fmt)
    return ch

def waterfall_bridge(sh, hv, c1, r1, c2, r2, steps, title, hv_anchor="AB1"):
    """Cascada = bridge apilado clasico (base invisible + subtotal NAVY + var GOLD).
    steps = lista de (etiqueta, base_formula, sub_formula, var_formula) ya en terminos de SUM(...).
    Escribe helper en hv desde hv_anchor (4 cols: Paso/base/sub/var) y grafica apilado."""
    col0=ord(hv_anchor[0:2].rstrip("0123456789")[0]) if False else None  # simplificado abajo
    # escribir cabecera + filas en hv (cols AB..AE = 28..31 por defecto)
    base_col=28
    hv.Cells(1,base_col).Value="Paso"; hv.Cells(1,base_col+1).Value="base"; hv.Cells(1,base_col+2).Value="sub"; hv.Cells(1,base_col+3).Value="var"
    for i,(lab,b,s,v) in enumerate(steps):
        hv.Cells(2+i,base_col).Value=lab
        hv.Cells(2+i,base_col+1).Formula=b if str(b).startswith("=") else "="+str(b)
        hv.Cells(2+i,base_col+2).Formula=s if str(s).startswith("=") else "="+str(s)
        hv.Cells(2+i,base_col+3).Formula=v if str(v).startswith("=") else "="+str(v)
    co=sh.ChartObjects().Add(sh.Columns(c1).Left+2, sh.Rows(r1).Top+1, 200, 120)
    place_chart(co, sh, c1, r1, c2, r2)
    ch=co.Chart; ch.ChartType=COLSTACK
    ch.SetSourceData(hv.Range(hv.Cells(1,base_col), hv.Cells(1+len(steps),base_col+3))); ch.PlotVisibleOnly=False
    try:
        ch.SeriesCollection(1).Format.Fill.Visible=False          # base invisible
        ch.SeriesCollection(2).Format.Fill.ForeColor.RGB=NAVY     # subtotal
        ch.SeriesCollection(3).Format.Fill.ForeColor.RGB=GOLD     # variacion
    except: pass
    style_chart(ch, title, legend=False)
    return ch

# ---------------- colores: capturar / re-aplicar (tras repuntar series) ----------------
def capture_colors(sh):
    """{(chart_idx, series_name): (fill_rgb, fill_visible, line_rgb)} — para re-aplicar tras repuntar."""
    ref={}
    for ci,co in enumerate(sh.ChartObjects(),1):
        ch=co.Chart
        for i in range(1, ch.SeriesCollection().Count+1):
            s=ch.SeriesCollection(i)
            try: nm=str(s.Name)
            except: nm=f"s{i}"
            try: fc=int(s.Format.Fill.ForeColor.RGB)
            except: fc=None
            try: fv=bool(s.Format.Fill.Visible)
            except: fv=True
            try: lc=int(s.Format.Line.ForeColor.RGB)
            except: lc=None
            ref[(ci,nm)]=(fc,fv,lc)
    return ref
def apply_colors(sh, ref):
    """Re-aplica colores capturados (capture_colors) por (chart_idx, series_name). Probado: persiste."""
    applied=miss=0
    for ci,co in enumerate(sh.ChartObjects(),1):
        ch=co.Chart
        for i in range(1, ch.SeriesCollection().Count+1):
            s=ch.SeriesCollection(i)
            try: nm=str(s.Name)
            except: nm=f"s{i}"
            key=(ci,nm)
            if key not in ref: miss+=1; continue
            fc,fv,lc=ref[key]
            if fv and fc is not None:
                try: s.Format.Fill.Solid(); s.Format.Fill.ForeColor.RGB=fc
                except: pass
            if lc is not None:
                try: s.Format.Line.ForeColor.RGB=lc
                except: pass
            applied+=1
    return applied, miss

def anchor_all_charts(sh):
    """Placement=xlMove en todos los charts (se mueven con su celda si la hoja se compacta)."""
    n=0
    for co in sh.ChartObjects():
        try: co.Placement=XLMOVE; n+=1
        except: pass
    return n

# ---------------- filtro de periodo (interactividad sin VBA) ----------------
def add_period_filter(wb, sh, hv, desde_cell="C9", hasta_cell="G9", years_row=7,
                      idx_cells=("A50","A51","A52","A53")):
    """Crea el filtro de periodo: validacion en celda + nombres OFFSET + re-apunta series de tiempo.
    OJO: tras esto, RE-APLICAR COLORES (apply_colors) porque setear Values resetea el color."""
    fname=os.path.basename(wb.FullName)
    A_des,A_has,A_start,A_cnt = idx_cells
    # 1) validacion en celda
    def vcell(addr, year):
        c=sh.Range(addr); c.Validation.Delete()
        c.Validation.Add(Type=XLVALIDATE_LIST, AlertStyle=1, Operator=1, Formula1=f"=Dashboard!${chr(64+17)}${years_row}:${'AD'}${years_row}")
        c.Value=year; c.Font.Name="Arial"; c.Font.Size=10; c.Font.Bold=True; c.Font.Color=NAVY
        c.HorizontalAlignment=CEN; c.Borders.Color=GOLD; c.Borders.Weight=2
    y0=int(sh.Cells(years_row,17).Value); y1=int(sh.Cells(years_row,30).Value)
    sh.Cells(9,2).Value="Desde"; sh.Cells(9,2).Font.Size=8; sh.Cells(9,2).Font.Color=DGRAY
    sh.Cells(9,5).Value="Hasta"; sh.Cells(9,5).Font.Size=8; sh.Cells(9,5).Font.Color=DGRAY
    try: sh.Range("C9:D9").Merge(); sh.Range("G9:H9").Merge()
    except: pass
    vcell("C9", y0); vcell("G9", y1)
    sh.Cells(8,2).Value="FILTRO DE PERIODO:"; sh.Cells(8,2).Font.Bold=True; sh.Cells(8,2).Font.Size=9; sh.Cells(8,2).Font.Color=NAVY
    # 2) indices en _DashViz
    hv.Range(A_des).Formula=f"=MATCH(Dashboard!$C$9,Dashboard!$Q${years_row}:$AD${years_row},0)"
    hv.Range(A_has).Formula=f"=MATCH(Dashboard!$G$9,Dashboard!$Q${years_row}:$AD${years_row},0)"
    hv.Range(A_start).Formula=f"={A_des}"
    hv.Range(A_cnt).Formula=f"=MAX(1,{A_has}-{A_des}+1)"
    # 3) nombre del eje
    def addname(nm, refers):
        try: wb.Names(nm).Delete()
        except: pass
        wb.Names.Add(Name=nm, RefersTo=refers)
    addname("f_x", f"=OFFSET(Dashboard!$Q${years_row},0,_DashViz!${A_start}-1,1,_DashViz!${A_cnt})")
    made=set()
    def dynname(sheet, r):
        key=("g_" if sheet=="_DashViz" else "f_")+str(r)
        if key not in made:
            addname(key, f"=OFFSET({sheet}!$C${r},0,_DashViz!${A_start}-1,1,_DashViz!${A_cnt})")
            made.add(key)
        return key
    # 4) re-apuntar series de tiempo (Values via setter; NO tocar SERIES())
    pat=re.compile(r"(Dashboard|_DashViz)!\$C\$(\d+):\$P\$(\d+)")
    patx=re.compile(r"Dashboard!\$Q\$"+str(years_row)+r":\$AD\$"+str(years_row))
    rep=0
    for co in sh.ChartObjects():
        ch=co.Chart
        try: ns=ch.SeriesCollection().Count
        except: ns=0
        plan=[]
        for i in range(1, ns+1):
            try: ser=str(ch.SeriesCollection(i).Formula)
            except: ser=""
            mm=re.match(r"=SERIES\((.*),(.*),(.*),(\d+)\)$", ser)
            if not mm: plan.append((i,None,False)); continue
            xref,vref=mm.group(2).strip(),mm.group(3).strip()
            mv=pat.search(vref)
            plan.append((i, (dynname(mv.group(1),mv.group(2)) if mv else None), bool(patx.search(xref))))
        for i,dn,rx in plan:
            if not dn: continue
            try:
                ch.SeriesCollection(i).Values=f"='{fname}'!{dn}"
                if rx: ch.SeriesCollection(i).XValues=f"='{fname}'!f_x"
                rep+=1
            except: pass
    return rep, len(made)

# ---------------- render (revisar con los ojos) ----------------
def render_range(sh, xl, a1, out_png, zoom=85):
    """Captura un rango (con charts flotantes) a PNG. Requiere PIL. visible=True al abrir."""
    from PIL import ImageGrab
    sh.Activate(); xl.ScreenUpdating=True; xl.ActiveWindow.Zoom=zoom
    cv.calc_wait(xl, full=True); time.sleep(1)
    for _ in range(6):
        try:
            sh.Range(a1).CopyPicture(1,2); time.sleep(0.8)
            im=ImageGrab.grabclipboard()
            if im:
                if im.width>1900: im=im.resize((1900,int(im.height*1900/im.width)))
                im.save(out_png); return out_png
        except: time.sleep(0.8)
    return None
