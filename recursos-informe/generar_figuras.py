# -*- coding: utf-8 -*-
"""Figuras del informe de internado: organigrama y cartas Gantt SEMANALES.
Se exportan en PNG (400 dpi) y PDF (vectorial) para insertar manualmente."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch
import matplotlib.patheffects as pe

plt.rcParams["font.family"] = "DejaVu Sans"

NAVY = "#1f3b5c"
BLUE = "#2f6db0"
LBLUE = "#cfe0f2"
GREEN = "#4b8b5b"
LGREEN = "#d9ead0"
GREY = "#5b6570"
R = "recursos-informe/"

# =====================================================================
# FIGURA 1: ORGANIGRAMA DE LA UNIDAD DE FARMACIA
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
                fontsize=ssize, color=GREY, style="italic")
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
     "Atención Farmacéutica en box y visitas domiciliarias (PSCV)",
     fc="#eaf1fa", tsize=11)
tens = [("TENS", "Briyit Ortiz"), ("TENS", "Paola Romero"),
        ("TENS", "Génesis Cartes"), ("TENS", "Dominic Cartes")]
xs = [2.6, 6.2, 9.8, 13.4]
for x, (r, n) in zip(xs, tens):
    rbox(ax, x, 3.0, 3.2, 1.05, r, n, fc=LGREEN, ec=GREEN, tcolor=GREEN, tsize=11, ssize=9.5)
ax.text(8, 3.75, "Técnicos en Enfermería de Nivel Superior (TENS)",
        ha="center", va="center", fontsize=9.5, color=GREEN, weight="bold",
        bbox=dict(boxstyle="round,pad=0.2", fc="white", ec="none"))
ax.text(8, 2.15, "Recepción, almacenamiento, fraccionamiento, reenvasado y despacho de medicamentos",
        ha="center", va="center", fontsize=8.8, color=GREY, style="italic")
rbox(ax, 8, 0.9, 6.2, 0.9, "Interno de Química y Farmacia",
     "Integrado a la unidad durante la rotación", fc="#fbf6e5",
     ec=GREY, tcolor=GREY, tsize=10.5, ssize=8.5, dashed=True)
conn(ax, 8, 8.5, 8, 7.68); conn(ax, 8, 6.52, 8, 5.78); conn(ax, 8, 4.62, 8, 4.05)
conn(ax, xs[0], 4.05, xs[-1], 4.05)
for x in xs:
    conn(ax, x, 4.05, x, 3.53)
conn(ax, 8, 2.47, 8, 1.36, color=GREY)
plt.tight_layout()
plt.savefig(R+"fig_organigrama.png", dpi=400, bbox_inches="tight")
plt.savefig(R+"fig_organigrama.pdf", bbox_inches="tight")
plt.close()

# =====================================================================
# CARTAS GANTT SEMANALES (9 semanas)
# =====================================================================
def gantt_semanal(activities, filename, title, color):
    weeks = ["S1", "S2", "S3", "S4", "S5", "S6", "S7", "S8", "S9"]
    fig, ax = plt.subplots(figsize=(15, 8.5))
    n = len(activities)
    for i, (name, spans) in enumerate(activities):
        y = n - i - 1
        ax.axhspan(y - 0.5, y + 0.5, color=("#f4f7fb" if i % 2 == 0 else "#ffffff"),
                   zorder=0)
        for (start, end) in spans:
            ax.barh(y, end - start + 1, left=start - 1, height=0.56, color=color,
                    edgecolor=NAVY, linewidth=1.0, zorder=3, alpha=0.92)
    ax.set_yticks(range(n))
    ax.set_yticklabels([a[0] for a in reversed(activities)], fontsize=12.5)
    ax.set_xticks([i + 0.5 for i in range(9)])
    ax.set_xticklabels(weeks, fontsize=13, weight="bold")
    ax.set_xlim(0, 9); ax.set_ylim(-0.5, n - 0.5)
    for x in range(10):
        ax.axvline(x, color="#c7d0da", linewidth=0.8, zorder=1)
    # segunda fila de fechas bajo las semanas
    fechas = ["11-15\nmay", "18-22\nmay", "25-29\nmay", "01-05\njun", "08-12\njun",
              "15-19\njun", "22-26\njun", "29 jun\n03 jul", "06-10\njul"]
    secax = ax.secondary_xaxis("bottom")
    secax.set_xticks([i + 0.5 for i in range(9)])
    secax.set_xticklabels(fechas, fontsize=9, color=GREY)
    secax.tick_params(length=0, pad=22)
    secax.spines["bottom"].set_visible(False)
    ax.set_title(title, fontsize=16, weight="bold", color=NAVY, pad=16)
    ax.tick_params(length=0)
    for s in ["top", "right", "left"]:
        ax.spines[s].set_visible(False)
    plt.tight_layout()
    plt.savefig(filename, dpi=400, bbox_inches="tight")
    plt.savefig(filename.replace(".png", ".pdf"), bbox_inches="tight")
    plt.close()

# Actividades PLANIFICADAS (plan acordado al inicio)
planificadas = [
    ("Inducción y conocimiento del centro y la farmacia", [(1, 1)]),
    ("Recepción, almacenamiento y gestión de stock", [(1, 2)]),
    ("Fraccionamiento y reenvasado de medicamentos", [(2, 9)]),
    ("Rotulación y despacho de medicamentos", [(2, 9)]),
    ("Atención Farmacéutica y visitas domiciliarias", [(1, 9)]),
    ("Control de indicadores y registros de despacho", [(3, 8)]),
    ("Manejo de medicamentos de control legal", [(4, 9)]),
    ("Educación sanitaria y uso racional (taller)", [(7, 7)]),
    ("Inventario y ordenamiento de bodega", [(5, 5)]),
    ("Eliminación de medicamentos vencidos", [(8, 9)]),
    ("Apoyo en supervisiones del Servicio de Salud", [(4, 5)]),
    ("Seminario de título (protocolo CMO)", [(1, 9)]),
]
gantt_semanal(planificadas, R+"fig_gantt_planificada.png",
              "Carta Gantt de actividades planificadas (9 semanas)", BLUE)

# Actividades DESARROLLADAS (lo realmente ejecutado)
desarrolladas = [
    ("Inducción y conocimiento del centro y la farmacia", [(1, 1)]),
    ("Recepción, almacenamiento y gestión de stock", [(1, 2), (4, 4)]),
    ("Fraccionamiento y reenvasado de medicamentos", [(1, 9)]),
    ("Rotulación y despacho de medicamentos", [(2, 9)]),
    ("Atención Farmacéutica y visitas domiciliarias", [(1, 9)]),
    ("Control de indicadores y registros de despacho", [(3, 3), (4, 5), (8, 8)]),
    ("Manejo de medicamentos de control legal", [(4, 9)]),
    ("Educación sanitaria y uso racional (taller)", [(7, 7)]),
    ("Inventario y ordenamiento de bodega", [(5, 5)]),
    ("Eliminación de medicamentos vencidos", [(8, 9)]),
    ("Apoyo en supervisiones del Servicio de Salud", [(4, 5)]),
    ("Seminario de título (protocolo CMO)", [(1, 9)]),
]
gantt_semanal(desarrolladas, R+"fig_gantt_desarrollada.png",
              "Carta Gantt de actividades desarrolladas (9 semanas)", GREEN)

import os
from PIL import Image
print("Figuras generadas:")
for f in ["fig_organigrama.png", "fig_gantt_planificada.png", "fig_gantt_desarrollada.png"]:
    im = Image.open(R + f)
    print(f"  {f}  {im.size}  {os.path.getsize(R+f)//1024} KB")
