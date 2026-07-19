# -*- coding: utf-8 -*-
"""Figuras del informe de internado: organigrama y cartas Gantt (alta resolucion)."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch
import matplotlib.patheffects as pe

plt.rcParams["font.family"] = "DejaVu Sans"
plt.rcParams["figure.dpi"] = 300

NAVY = "#1f3b5c"
BLUE = "#2f6db0"
LBLUE = "#cfe0f2"
GREEN = "#4b8b5b"
LGREEN = "#d9ead0"
GREY = "#5b6570"

# =====================================================================
# FIGURA 1: ORGANIGRAMA DE LA UNIDAD DE FARMACIA (alta resolucion)
# =====================================================================
def rbox(ax, x, y, w, h, title, subtitle="", fc="#ffffff", ec=NAVY,
         tcolor=NAVY, tsize=12, ssize=8.5, dashed=False):
    box = FancyBboxPatch((x - w/2, y - h/2), w, h,
                         boxstyle="round,pad=0.015,rounding_size=0.12",
                         linewidth=1.8, edgecolor=ec, facecolor=fc,
                         linestyle=("--" if dashed else "-"))
    box.set_path_effects([pe.withSimplePatchShadow(offset=(2.0, -2.0),
                          shadow_rgbFace="#9aa4b0", alpha=0.30)])
    ax.add_patch(box)
    if subtitle:
        ax.text(x, y + h*0.17, title, ha="center", va="center",
                fontsize=tsize, weight="bold", color=tcolor)
        ax.text(x, y - h*0.22, subtitle, ha="center", va="center",
                fontsize=ssize, color=GREY, style="italic", wrap=True)
    else:
        ax.text(x, y, title, ha="center", va="center", fontsize=tsize,
                weight="bold", color=tcolor)

def conn(ax, x1, y1, x2, y2, color=NAVY):
    ax.plot([x1, x2], [y1, y2], color=color, linewidth=1.6, zorder=0,
            solid_capstyle="round")

fig, ax = plt.subplots(figsize=(11.5, 7.6))
ax.set_xlim(0, 16); ax.set_ylim(0, 11); ax.axis("off")

ax.text(8, 10.5, "Organigrama de la Unidad de Farmacia",
        ha="center", va="center", fontsize=15, weight="bold", color=NAVY)
ax.text(8, 10.0, "CESFAM Villa Nonguén", ha="center", va="center",
        fontsize=11, color=GREY, style="italic")

rbox(ax, 8, 9.0, 6.6, 1.0, "Dirección del CESFAM Villa Nonguén",
     "Administración delegada, COSADES", fc=LBLUE, tsize=12)
rbox(ax, 8, 7.1, 7.2, 1.15, "Q.F. Directora Técnica de Farmacia  |  Jennifer Vallejos",
     "Dirección técnica, arsenal, control legal y gestión de la unidad",
     fc="#eaf1fa", tsize=11)
rbox(ax, 8, 5.2, 7.2, 1.15, "Q.F. de Atención Farmacéutica  |  Fernanda Torres",
     "Atención farmacéutica en box y visitas domiciliarias (PSCV)",
     fc="#eaf1fa", tsize=11)

tens = [
    ("TENS", "Briyit Ortiz"),
    ("TENS", "Paola Romero"),
    ("TENS", "Génesis Cartes"),
    ("TENS", "Dominic Cartes"),
]
xs = [2.6, 6.2, 9.8, 13.4]
for x, (r, n) in zip(xs, tens):
    rbox(ax, x, 3.0, 3.2, 1.05, r, n, fc=LGREEN, ec=GREEN, tcolor=GREEN, tsize=11, ssize=9.5)
ax.text(8, 3.75, "Técnicos en Enfermería de Nivel Superior (TENS)",
        ha="center", va="center", fontsize=9.5, color=GREEN, weight="bold",
        bbox=dict(boxstyle="round,pad=0.2", fc="white", ec="none"))
ax.text(8, 2.15, "Recepción, almacenamiento, fraccionamiento, reenvasado y dispensación de medicamentos",
        ha="center", va="center", fontsize=8.8, color=GREY, style="italic")

rbox(ax, 8, 0.9, 6.2, 0.9, "Interno de Química y Farmacia",
     "Integrado a la unidad durante la rotación", fc="#fbf6e5",
     ec=GREY, tcolor=GREY, tsize=10.5, ssize=8.5, dashed=True)

conn(ax, 8, 8.5, 8, 7.68)
conn(ax, 8, 6.52, 8, 5.78)
conn(ax, 8, 4.62, 8, 4.05)
conn(ax, xs[0], 4.05, xs[-1], 4.05)
for x in xs:
    conn(ax, x, 4.05, x, 3.53)
conn(ax, 8, 2.47, 8, 1.36, color=GREY)

plt.tight_layout()
plt.savefig("recursos-informe/fig_organigrama.png", dpi=300, bbox_inches="tight")
plt.close()

# =====================================================================
# FIGURA 2: CARTA GANTT PLANIFICADA (semanal, horizontal, alta calidad)
# =====================================================================
def gantt_semanal(activities, filename, title, color):
    weeks = [f"S{i+1}" for i in range(9)]
    fig, ax = plt.subplots(figsize=(15, 8.5))
    n = len(activities)
    for i, (name, spans) in enumerate(activities):
        y = n - i - 1
        ax.axhspan(y - 0.5, y + 0.5, color=("#f4f7fb" if i % 2 == 0 else "#ffffff"),
                   zorder=0)
        for (start, end) in spans:
            ax.barh(y, end - start + 1, left=start - 1, height=0.55, color=color,
                    edgecolor=NAVY, linewidth=0.9, zorder=3, alpha=0.92)
    ax.set_yticks(range(n))
    ax.set_yticklabels([a[0] for a in reversed(activities)], fontsize=12.5)
    ax.set_xticks([i + 0.5 for i in range(9)])
    ax.set_xticklabels(weeks, fontsize=13, weight="bold")
    ax.set_xlim(0, 9); ax.set_ylim(-0.5, n - 0.5)
    for x in range(10):
        ax.axvline(x, color="#c7d0da", linewidth=0.7, zorder=1)
    ax.set_xlabel("Semanas de internado (11 de mayo al 10 de julio de 2026)",
                  fontsize=12.5)
    ax.set_title(title, fontsize=16, weight="bold", color=NAVY, pad=16)
    ax.tick_params(length=0)
    for s in ["top", "right", "left"]:
        ax.spines[s].set_visible(False)
    plt.tight_layout()
    plt.savefig(filename, dpi=400, bbox_inches="tight")
    plt.savefig(filename.replace(".png", ".pdf"), bbox_inches="tight")
    plt.close()

planificadas = [
    ("Inducción y conocimiento del centro", [(1, 1)]),
    ("Recepción, almacenamiento y gestión de stock", [(1, 2)]),
    ("Fraccionamiento y reenvasado", [(2, 4)]),
    ("Dispensación informada de medicamentos", [(3, 7)]),
    ("Control de indicadores y registros", [(3, 8)]),
    ("Atención farmacéutica y visitas domiciliarias", [(2, 9)]),
    ("Educación y uso racional de medicamentos", [(6, 7)]),
    ("Inventario y ordenamiento de bodega", [(5, 5)]),
    ("Apoyo en supervisiones del Servicio de Salud", [(4, 5)]),
    ("Seminario de título (protocolo CMO)", [(1, 9)]),
    ("Reuniones de supervisión y elaboración del informe", [(6, 6), (9, 9)]),
]
gantt_semanal(planificadas, "recursos-informe/fig_gantt_planificada.png",
              "Carta Gantt de actividades planificadas", BLUE)

# =====================================================================
# FIGURA 3: CARTA GANTT DESARROLLADA (detalle diario, vertical, grande)
# =====================================================================
activities = [
    "Inducción y conocimiento del centro",
    "Recepción y gestión de stock",
    "Fraccionamiento y reenvasado",
    "Dispensación de medicamentos",
    "Control de indicadores y registros",
    "Atención farmacéutica y visitas domiciliarias",
    "Educación y uso racional (taller)",
    "Inventario y ordenamiento de bodega",
    "Supervisiones del Servicio de Salud",
    "Vinculación con el medio",
    "Seminario de título (protocolo CMO)",
    "Reuniones de supervisión e informe",
]
acol = ["#3f6fa3", "#4a7fb5", "#5aa06a", "#6bb07a", "#c99a3b", "#c0504d",
        "#8e6fb0", "#6f9c8a", "#b5651d", "#7a8b99", "#2f4b7c", "#9c3f6a"]

# schedule[week][day] = list de indices de actividad (day 0=Lu ... 4=Vi)
schedule = {
 0: {0:[0], 1:[2,1], 2:[4,2], 3:[2,5,0], 4:[10]},
 1: {0:[10,5], 1:[10,2,1], 2:[2], 3:[3], 4:[1,3]},
 2: {0:[5,9], 1:[5,10], 2:[10], 3:[3], 4:[3]},
 3: {0:[1], 1:[4], 2:[1,4], 3:[8,0,10], 4:[4,5]},
 4: {0:[3,1], 1:[7], 2:[8,10], 3:[3,2], 4:[2,5]},
 5: {0:[11,10], 1:[5,3], 2:[5,3], 3:[5,3], 4:[5]},
 6: {0:[1,3], 1:[1,3], 2:[2,3], 3:[3], 4:[6]},
 7: {0:[3,1], 1:[3,5], 2:[1,3], 3:[5,9], 4:[10]},
 8: {0:[5], 1:[5], 2:[5], 3:[2,1], 4:[5]},
}
week_dates = ["S1  11-15 may", "S2  18-22 may", "S3  25-29 may", "S4  01-05 jun",
              "S5  08-12 jun", "S6  15-19 jun", "S7  22-26 jun", "S8  29 jun-03 jul",
              "S9  06-10 jul"]
days = ["Lu", "Ma", "Mi", "Ju", "Vi"]

nrows = 9 * 5
ncols = len(activities)
fig, ax = plt.subplots(figsize=(12.5, 18.5))
ax.set_xlim(-0.5, ncols); ax.set_ylim(0, nrows)
ax.invert_yaxis()

# encabezados de columna: numeros 1..12 (la leyenda mapea numero -> actividad)
for c in range(ncols):
    ax.text(c + 0.5, -0.35, str(c + 1), ha="center", va="bottom",
            fontsize=15, color=NAVY, weight="bold")

# grilla y celdas
for w in range(9):
    for d in range(5):
        row = w * 5 + d
        # sombreado alterno por semana
        if w % 2 == 0:
            ax.axhspan(row, row + 1, xmin=0, xmax=1, color="#f5f8fc", zorder=0)
        acts = schedule[w].get(d, [])
        for a in acts:
            ax.add_patch(plt.Rectangle((a + 0.12, row + 0.15), 0.76, 0.7,
                         facecolor=acol[a], edgecolor="white", linewidth=0.6, zorder=3))
    # separador de semana
    ax.axhline(w * 5, color="#9aa4b0", linewidth=1.1, zorder=2)
ax.axhline(nrows, color="#9aa4b0", linewidth=1.1)

for c in range(ncols + 1):
    ax.axvline(c, color="#dbe2ea", linewidth=0.6, zorder=1)

# etiquetas de dia (eje Y) y semana
yt, yl = [], []
for w in range(9):
    for d in range(5):
        row = w * 5 + d
        yt.append(row + 0.5); yl.append(days[d])
ax.set_yticks(yt); ax.set_yticklabels(yl, fontsize=10.5)
ax.set_xticks([])
ax.tick_params(length=0)
for s in ax.spines.values():
    s.set_visible(False)

# etiquetas de semana a la izquierda
for w in range(9):
    ax.text(-1.9, w * 5 + 2.5, week_dates[w], rotation=90, ha="center", va="center",
            fontsize=11, weight="bold", color=NAVY, clip_on=False)

# leyenda: numero de columna -> actividad
import matplotlib.patches as mpatches
handles = [mpatches.Patch(facecolor=acol[i], edgecolor="#ffffff",
           label="%d.  %s" % (i + 1, activities[i])) for i in range(ncols)]
fig.legend(handles=handles, loc="lower center", ncol=2, fontsize=12,
           frameon=True, bbox_to_anchor=(0.5, 0.006), handlelength=1.6,
           columnspacing=1.6, borderpad=1.0,
           title="Referencia de actividades (número de columna)", title_fontsize=13)

fig.suptitle("Carta Gantt de actividades desarrolladas (detalle por día y semana)",
             fontsize=16, weight="bold", color=NAVY, y=0.995)
fig.subplots_adjust(top=0.94, left=0.14, right=0.99, bottom=0.17)
plt.savefig("recursos-informe/fig_gantt_desarrollada.png", dpi=400, bbox_inches="tight")
plt.savefig("recursos-informe/fig_gantt_desarrollada.pdf", bbox_inches="tight")
plt.close()

import os
print("Figuras generadas:")
for f in ["fig_organigrama.png", "fig_gantt_planificada.png", "fig_gantt_desarrollada.png"]:
    print(" ", f, os.path.getsize("recursos-informe/" + f), "bytes")
