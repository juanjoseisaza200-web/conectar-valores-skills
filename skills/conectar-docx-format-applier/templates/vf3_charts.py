"""
vf3_charts.py — Generador de gráficas matplotlib con paleta y estilo Conectar Valores VF3.

Paleta verificada vs image1.png del VF3 real:
- Teal     #5D9E9D  (1ª serie / acento)
- Navy     #17375E  (2ª serie / dominante / títulos)
- Slate    #4A5568  (3ª serie / texto auxiliar)
- Hairline #BFBFBF  (4ª serie / "Otros" / gridlines)
- Dorado   #C9A449  (acento sutil — usar moderado)

Uso típico:
    from vf3_charts import chart_double_bar, chart_single_bar, chart_line, chart_pie
    chart_double_bar(
        title="Estructura del Mercado",
        left=("Prestadores por Naturaleza",  ["Oficial","Privada","Mixta","Otros"], [569,259,165,11]),
        right=("Suscriptores por Naturaleza",["Oficial","Privada","Mixta","Otros"], [4.6,14.4,1.6,0.05]),
        left_ylabel="Cantidad",
        right_ylabel="Millones de suscriptores",
        source="Fuente: SUI CONSOLIDADO, abril 2026",
        output_path="figura1.png"
    )
"""
import os
import matplotlib
matplotlib.use('Agg')  # backend no-GUI
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator
import numpy as np

# Paleta hex CV (idéntica a vf3_styles.py)
NAVY     = "#17375E"
TEAL     = "#5D9E9D"
SLATE    = "#4A5568"
HAIRLINE = "#BFBFBF"
DORADO   = "#C9A449"
WHITE    = "#FFFFFF"

# Paleta ordenada para series (1ª, 2ª, 3ª, 4ª, 5ª)
SERIES_PALETTE = [TEAL, NAVY, SLATE, HAIRLINE, DORADO]

# Configuración matplotlib estilo CV
def _apply_style(ax, title=None, ylabel=None):
    """Aplica el estilo CV a un Axes."""
    if title:
        ax.set_title(title, fontfamily='Arial', fontsize=12, fontweight='bold',
                     color=NAVY, pad=15)
    if ylabel:
        ax.set_ylabel(ylabel, fontfamily='Arial', fontsize=9, color=SLATE)
    # Quitar bordes externos (top, right, left)
    for spine in ('top', 'right', 'left'):
        ax.spines[spine].set_visible(False)
    ax.spines['bottom'].set_color(SLATE)
    ax.spines['bottom'].set_linewidth(0.8)
    # Gridlines horizontales muy sutiles
    ax.yaxis.grid(True, color='#E5E5E5', linewidth=0.5, zorder=0)
    ax.xaxis.grid(False)
    ax.set_axisbelow(True)
    # Tick labels
    ax.tick_params(axis='x', colors=SLATE, length=0, labelsize=9)
    ax.tick_params(axis='y', colors=SLATE, length=0, labelsize=8)
    for lbl in ax.get_xticklabels():
        lbl.set_fontfamily('Arial')
    for lbl in ax.get_yticklabels():
        lbl.set_fontfamily('Arial')
    # Fondo blanco
    ax.set_facecolor(WHITE)


def _add_data_labels(ax, x_pos, values, total=None, fmt_value=None, fmt_pct=True):
    """Añade labels Navy bold encima de cada barra: valor + (%)."""
    if fmt_value is None:
        fmt_value = lambda v: f"{int(v):,}" if v >= 1 else f"{v:.2f}"
    total = total if total is not None else sum(values)
    for x, v in zip(x_pos, values):
        label_v = fmt_value(v)
        if fmt_pct and total > 0:
            pct = v / total * 100
            label = f"{label_v}\n({pct:.0f}%)" if pct >= 1 else f"{label_v}\n({pct:.1f}%)"
        else:
            label = label_v
        ax.text(x, v, label, ha='center', va='bottom',
                fontfamily='Arial', fontsize=8.5, fontweight='bold',
                color=NAVY)


def chart_single_bar(
    *, title, categories, values,
    ylabel=None, source=None, output_path,
    fmt_value=None, fmt_pct=True,
    fig_width=7.0, fig_height=4.5,
    colors=None,
):
    """Gráfica de barras simple estilo CV.

    Args:
        title: título principal Navy bold arriba
        categories: lista de labels x
        values: lista de valores y
        ylabel: label eje y
        source: texto "Fuente: ..." italic Slate centrado abajo
        output_path: ruta PNG donde guardar
        fmt_value: función para formatear valores (default: int con separador miles)
        fmt_pct: si True, añade % de total al label
        colors: paleta override (default: SERIES_PALETTE)
    """
    fig, ax = plt.subplots(figsize=(fig_width, fig_height), dpi=150)
    fig.patch.set_facecolor(WHITE)

    colors = colors or [SERIES_PALETTE[i % len(SERIES_PALETTE)] for i in range(len(categories))]
    x = np.arange(len(categories))
    bars = ax.bar(x, values, color=colors, edgecolor='none', width=0.65, zorder=2)
    ax.set_xticks(x)
    ax.set_xticklabels(categories)

    _apply_style(ax, title=title, ylabel=ylabel)
    _add_data_labels(ax, x, values, fmt_value=fmt_value, fmt_pct=fmt_pct)

    # Padding superior para labels
    ymax = max(values) * 1.25
    ax.set_ylim(0, ymax)

    if source:
        fig.text(0.5, 0.02, source, ha='center', fontfamily='Arial',
                 fontsize=7.5, fontstyle='italic', color=SLATE)

    plt.tight_layout(rect=[0, 0.05, 1, 1])
    plt.savefig(output_path, dpi=150, bbox_inches='tight', facecolor=WHITE)
    plt.close()
    return output_path


def chart_double_bar(
    *, title, left, right,
    left_ylabel=None, right_ylabel=None,
    source=None, output_path,
    fmt_left=None, fmt_right=None,
    fig_width=10.5, fig_height=5.0,
):
    """Dos gráficas de barras side-by-side bajo un título compartido.

    left = (subtitle, categories, values)
    right = (subtitle, categories, values)
    """
    fig, (ax_l, ax_r) = plt.subplots(1, 2, figsize=(fig_width, fig_height), dpi=150)
    fig.patch.set_facecolor(WHITE)

    # Título compartido arriba
    fig.suptitle(title, fontfamily='Arial', fontsize=13, fontweight='bold',
                 color=NAVY, y=0.97)

    for ax, (subtitle, cats, vals), ylabel, fmt in (
        (ax_l, left, left_ylabel, fmt_left),
        (ax_r, right, right_ylabel, fmt_right),
    ):
        x = np.arange(len(cats))
        colors = [SERIES_PALETTE[i % len(SERIES_PALETTE)] for i in range(len(cats))]
        ax.bar(x, vals, color=colors, edgecolor='none', width=0.65, zorder=2)
        ax.set_xticks(x)
        ax.set_xticklabels(cats)
        _apply_style(ax, title=subtitle, ylabel=ylabel)
        _add_data_labels(ax, x, vals, fmt_value=fmt, fmt_pct=True)
        ax.set_ylim(0, max(vals) * 1.25)

    if source:
        fig.text(0.5, 0.02, source, ha='center', fontfamily='Arial',
                 fontsize=7.5, fontstyle='italic', color=SLATE)

    plt.tight_layout(rect=[0, 0.05, 1, 0.93])
    plt.savefig(output_path, dpi=150, bbox_inches='tight', facecolor=WHITE)
    plt.close()
    return output_path


def chart_line(
    *, title, x_categories, series,
    ylabel=None, source=None, output_path,
    fig_width=8.0, fig_height=4.5,
    markers=True,
):
    """Gráfica de líneas estilo CV.

    series = {"Nombre serie": [valores], ...}
    """
    fig, ax = plt.subplots(figsize=(fig_width, fig_height), dpi=150)
    fig.patch.set_facecolor(WHITE)

    x = np.arange(len(x_categories))
    for i, (name, vals) in enumerate(series.items()):
        color = SERIES_PALETTE[i % len(SERIES_PALETTE)]
        ax.plot(x, vals, color=color, linewidth=2.2, label=name,
                marker='o' if markers else None, markersize=6,
                markerfacecolor=color, markeredgecolor=WHITE, markeredgewidth=1.5,
                zorder=3)
    ax.set_xticks(x)
    ax.set_xticklabels(x_categories)
    _apply_style(ax, title=title, ylabel=ylabel)

    # Leyenda abajo horizontal si > 1 serie
    if len(series) > 1:
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.10),
                  ncol=len(series), frameon=False, fontsize=8.5,
                  labelcolor=SLATE)

    if source:
        fig.text(0.5, 0.02, source, ha='center', fontfamily='Arial',
                 fontsize=7.5, fontstyle='italic', color=SLATE)

    plt.tight_layout(rect=[0, 0.08, 1, 1])
    plt.savefig(output_path, dpi=150, bbox_inches='tight', facecolor=WHITE)
    plt.close()
    return output_path


def chart_pie(
    *, title, categories, values, source=None, output_path,
    fig_width=6.5, fig_height=5.0,
):
    """Donut estilo CV — paleta navy/teal/slate/hairline."""
    fig, ax = plt.subplots(figsize=(fig_width, fig_height), dpi=150)
    fig.patch.set_facecolor(WHITE)

    colors = [SERIES_PALETTE[i % len(SERIES_PALETTE)] for i in range(len(categories))]
    wedges, texts, autotexts = ax.pie(values, labels=categories, colors=colors,
        autopct='%.1f%%', startangle=90, pctdistance=0.78,
        wedgeprops=dict(width=0.42, edgecolor=WHITE, linewidth=2),
        textprops=dict(fontfamily='Arial', fontsize=9, color=NAVY))
    for at in autotexts:
        at.set_fontweight('bold')
        at.set_color(WHITE)
        at.set_fontsize(9)

    ax.set_title(title, fontfamily='Arial', fontsize=12, fontweight='bold',
                 color=NAVY, pad=15)

    if source:
        fig.text(0.5, 0.02, source, ha='center', fontfamily='Arial',
                 fontsize=7.5, fontstyle='italic', color=SLATE)

    plt.tight_layout(rect=[0, 0.05, 1, 1])
    plt.savefig(output_path, dpi=150, bbox_inches='tight', facecolor=WHITE)
    plt.close()
    return output_path


# Smoke test
if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    out_dir = os.path.normpath(os.path.join(here, '..', 'assets', 'chart_smoke'))
    os.makedirs(out_dir, exist_ok=True)

    chart_double_bar(
        title="Estructura del Mercado de Aseo — Colombia (smoke test)",
        left=("Prestadores por Naturaleza",
              ["Oficial", "Privada", "Mixta", "Otros"],
              [569, 259, 165, 11]),
        right=("Suscriptores por Naturaleza (millones)",
               ["Oficial", "Privada", "Mixta", "Otros"],
               [4.6, 14.4, 1.6, 0.05]),
        left_ylabel="Cantidad",
        right_ylabel="Millones de suscriptores",
        source="Fuente: SUI CONSOLIDADO, abril 2026 — smoke test",
        output_path=os.path.join(out_dir, "smoke_double_bar.png"),
        fmt_right=lambda v: f"{v:.1f}M" if v >= 0.1 else f"{int(v*1000)}K",
    )
    chart_line(
        title="Evolución 2020-2024 — smoke",
        x_categories=["2020","2021","2022","2023","2024"],
        series={"Ingresos": [100,108,118,128,140], "EBITDA": [32,35,38,42,45]},
        ylabel="COP MM", source="Fuente: smoke test",
        output_path=os.path.join(out_dir, "smoke_line.png"),
    )
    chart_pie(
        title="Mix tecnológico — smoke",
        categories=["Hidro","Térmica","Solar"],
        values=[62.8, 29.6, 7.6],
        source="Fuente: smoke test",
        output_path=os.path.join(out_dir, "smoke_pie.png"),
    )
    print(f"[OK] Smoke charts en: {out_dir}")
