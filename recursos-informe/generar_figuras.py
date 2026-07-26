# -*- coding: utf-8 -*-
"""Organigrama y TRES diseños distintos de carta Gantt SEMANAL (9 semanas),
cada actividad con un color propio. PNG 400 dpi + PDF vectorial."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Rectangle
import matplotlib.patheffects as pe

plt.rcParams["font.family"] = "DejaVu Sans"
NAVY = "#1f3b5c"; GREEN = "#4b8b5b"; LGREEN = "#d9ead0"; LBLUE = "#cfe0f2"; GREY = "#5b6570"
R = "recursos-informe/"

# =====================================================================
# ORGANIGRAMA
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
        ax.text(x, y + h*0.17, title, ha="center", va="center", fontsize=tsize, weight="bold", color=tcolor)
        ax.text(x, y - h*0.22, subtitle, ha="center", va="center", fontsize=ssize, color=GREY, style="italic")
    else:
        ax.text(x, y, title, ha="center", va="center", fontsize=tsize, weight="bold", color=tcolor)

def conn(ax, x1, y1, x2, y2, color=NAVY):
    ax.plot([x1, x2], [y1, y2], color=color, linewidth=1.6, zorder=0, solid_capstyle="round")

fig, ax = plt.subplots(figsize=(11.5, 7.6))
ax.set_xlim(0, 16); ax.set_ylim(0, 11); ax.axis("off")
ax.text(8, 10.5, "Organigrama de la Unidad de Farmacia", ha="center", va="center", fontsize=15, weight="bold", color=NAVY)
ax.text(8, 10.0, "CESFAM Villa Nonguén", ha="center", va="center", fontsize=11, color=GREY, style="italic")
rbox(ax, 8, 9.0, 6.6, 1.0, "Dirección del CESFAM Villa Nonguén", "Administración delegada, COSADES", fc=LBLUE, tsize=12)
rbox(ax, 8, 7.1, 7.2, 1.15, "Q.F. Directora Técnica de Farmacia  |  Yeniffer Vallejos",
     "Dirección técnica, arsenal, control legal y gestión de la unidad", fc="#eaf1fa", tsize=11)
rbox(ax, 8, 5.2, 7.2, 1.15, "Q.F. de Atención Farmacéutica  |  Fernanda Torres",
     "Atención Farmacéutica en box y visitas domiciliarias (PSCV)", fc="#eaf1fa", tsize=11)
tens = [("TENS", "Briyit Ortiz"), ("TENS", "Paola Romero"), ("TENS", "Génesis Cartes"), ("TENS", "Dominic Cartes")]
xs = [2.6, 6.2, 9.8, 13.4]
for x, (r, n) in zip(xs, tens):
    rbox(ax, x, 3.0, 3.2, 1.05, r, n, fc=LGREEN, ec=GREEN, tcolor=GREEN, tsize=11, ssize=9.5)
ax.text(8, 3.75, "Técnicos en Enfermería de Nivel Superior (TENS)", ha="center", va="center",
        fontsize=9.5, color=GREEN, weight="bold", bbox=dict(boxstyle="round,pad=0.2", fc="white", ec="none"))
ax.text(8, 2.15, "Recepción, almacenamiento, fraccionamiento, reenvasado y despacho de medicamentos",
        ha="center", va="center", fontsize=8.8, color=GREY, style="italic")
rbox(ax, 8, 0.9, 6.2, 0.9, "Interno de Química y Farmacia", "Integrado a la unidad durante la rotación",
     fc="#fbf6e5", ec=GREY, tcolor=GREY, tsize=10.5, ssize=8.5, dashed=True)
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
# DATOS DE LA CARTA GANTT (actividades desarrolladas, 9 semanas)
# =====================================================================
ACTS = [
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
# 12 colores distintos (paleta cualitativa)
COLORS = ["#4e79a7", "#f28e2b", "#59a14f", "#e15759", "#76b7b2", "#edc948",
          "#b07aa1", "#ff7f9e", "#9c755f", "#8c6bb1", "#17a589", "#c0504d"]
WEEKS = [f"S{i+1}" for i in range(9)]
FECHAS = ["11-15\nmay", "18-22\nmay", "25-29\nmay", "01-05\njun", "08-12\njun",
          "15-19\njun", "22-26\njun", "29 jun\n03 jul", "06-10\njul"]
TITLE = "Carta Gantt de actividades desarrolladas (9 semanas)"

def base_axes(ax, n):
    ax.set_yticks(range(n))
    ax.set_yticklabels([a[0] for a in reversed(ACTS)], fontsize=12)
    ax.set_xticks([i + 0.5 for i in range(9)])
    ax.set_xticklabels(WEEKS, fontsize=13, weight="bold")
    ax.set_xlim(0, 9); ax.set_ylim(-0.5, n - 0.5)
    secax = ax.secondary_xaxis("bottom")
    secax.set_xticks([i + 0.5 for i in range(9)])
    secax.set_xticklabels(FECHAS, fontsize=9, color=GREY)
    secax.tick_params(length=0, pad=24); secax.spines["bottom"].set_visible(False)
    ax.tick_params(length=0)
    for s in ["top", "right", "left"]:
        ax.spines[s].set_visible(False)

def save(fig, name):
    plt.tight_layout()
    fig.savefig(R + name + ".png", dpi=400, bbox_inches="tight")
    fig.savefig(R + name + ".pdf", bbox_inches="tight")
    plt.close(fig)

# ---------------------------------------------------------------------
# OPCION 1: barras horizontales, un color por actividad, filas separadas
# ---------------------------------------------------------------------
fig, ax = plt.subplots(figsize=(15, 9))
n = len(ACTS)
for i, (name, spans) in enumerate(ACTS):
    y = n - i - 1
    ax.axhspan(y - 0.5, y + 0.5, color=("#f6f8fb" if i % 2 == 0 else "#ffffff"), zorder=0)
    for (s, e) in spans:
        ax.barh(y, e - s + 1, left=s - 1, height=0.46, color=COLORS[i],
                edgecolor="white", linewidth=2.2, zorder=3)
for x in range(10):
    ax.axvline(x, color="#d7dee6", linewidth=0.8, zorder=1)
base_axes(ax, n)
ax.set_title(TITLE + "  ·  Opción 1", fontsize=16, weight="bold", color=NAVY, pad=16)
save(fig, "fig_gantt_opcion1")

# ---------------------------------------------------------------------
# OPCION 2: cuadrícula de celdas separadas + etiquetas de actividad en
# un recuadro con el color correspondiente (a la izquierda).
# ---------------------------------------------------------------------
import textwrap
LABELW = 5.0  # ancho (en unidades de dato) reservado para los recuadros de etiqueta
fig, ax = plt.subplots(figsize=(17, 9.2))
# grilla ligera de fondo (solo la zona de semanas)
for w in range(9):
    for i in range(n):
        ax.add_patch(Rectangle((w + 0.16, i - 0.32), 0.68, 0.64, facecolor="#f2f5f9",
                     edgecolor="#e3e8ee", linewidth=0.6, zorder=0))
for i, (name, spans) in enumerate(ACTS):
    y = n - i - 1
    # recuadro de etiqueta con el color de la actividad
    ax.add_patch(FancyBboxPatch((-LABELW + 0.15, y - 0.36), LABELW - 0.45, 0.72,
                 boxstyle="round,pad=0.02,rounding_size=0.10",
                 facecolor=COLORS[i], edgecolor="white", linewidth=1.6, zorder=3,
                 clip_on=False))
    ax.text(-LABELW + 0.15 + (LABELW - 0.45) / 2, y, textwrap.fill(name, 30),
            ha="center", va="center", fontsize=9.5, color="white", weight="bold",
            zorder=4, clip_on=False)
    # celdas de las semanas en que se realizó la actividad
    for (s, e) in spans:
        for w in range(s, e + 1):
            ax.add_patch(FancyBboxPatch((w - 1 + 0.16, y - 0.32), 0.68, 0.64,
                         boxstyle="round,pad=0.01,rounding_size=0.08",
                         facecolor=COLORS[i], edgecolor="white", linewidth=1.5, zorder=3))
# ejes: semanas arriba en x, sin etiquetas de texto en y
ax.set_yticks([])
ax.set_xticks([i + 0.5 for i in range(9)])
ax.set_xticklabels(WEEKS, fontsize=13, weight="bold")
ax.set_xlim(-LABELW, 9); ax.set_ylim(-0.5, n - 0.5)
secax = ax.secondary_xaxis("bottom")
secax.set_xticks([i + 0.5 for i in range(9)])
secax.set_xticklabels(FECHAS, fontsize=9, color=GREY)
secax.tick_params(length=0, pad=24); secax.spines["bottom"].set_visible(False)
ax.tick_params(length=0)
for spine in ["top", "right", "left"]:
    ax.spines[spine].set_visible(False)
ax.set_title(TITLE, fontsize=16, weight="bold", color=NAVY, pad=16)
save(fig, "fig_gantt_opcion2")

# ---------------------------------------------------------------------
# OPCION 3: barras redondeadas con sombra y fondo de semanas alternado
# ---------------------------------------------------------------------
fig, ax = plt.subplots(figsize=(15, 9.2))
# fondo alternado por semana (columnas)
for w in range(9):
    if w % 2 == 0:
        ax.axvspan(w, w + 1, color="#f4f7fb", zorder=0)
for i, (name, spans) in enumerate(ACTS):
    y = n - i - 1
    for (s, e) in spans:
        bar = FancyBboxPatch((s - 1 + 0.08, y - 0.28), (e - s + 1) - 0.16, 0.56,
                             boxstyle="round,pad=0.01,rounding_size=0.18",
                             facecolor=COLORS[i], edgecolor="white", linewidth=1.6, zorder=3)
        bar.set_path_effects([pe.withSimplePatchShadow(offset=(1.5, -1.5),
                              shadow_rgbFace="#b7c0ca", alpha=0.5)])
        ax.add_patch(bar)
for x in range(10):
    ax.axvline(x, color="#cfd8e0", linewidth=0.7, zorder=1)
base_axes(ax, n)
ax.set_title(TITLE + "  ·  Opción 3", fontsize=16, weight="bold", color=NAVY, pad=16)
save(fig, "fig_gantt_opcion3")

# =====================================================================
# CARTA GANTT DE ACTIVIDADES PLANIFICADAS (estilo Opción 2)
# Refleja la planificación inicial de la rotación; incluye actividades
# que se planificaron pero no se concretaron (Comité de Farmacia y
# farmacovigilancia), lo que permite justificarlas en el texto.
# =====================================================================
PLAN_ACTS = [
    ("Inducción y conocimiento del centro y la farmacia", [(1, 1)]),
    ("Recepción, almacenamiento y gestión de stock", [(1, 3)]),
    ("Fraccionamiento y reenvasado de medicamentos", [(2, 8)]),
    ("Rotulación y despacho de medicamentos", [(2, 9)]),
    ("Atención Farmacéutica y visitas domiciliarias", [(3, 9)]),
    ("Control de indicadores y registros de despacho", [(3, 6)]),
    ("Manejo de medicamentos de control legal", [(4, 6)]),
    ("Participación en Comité de Farmacia y Terapéutica", [(5, 5)]),
    ("Educación sanitaria y uso racional (taller)", [(6, 7)]),
    ("Inventario y ordenamiento de bodega", [(5, 6)]),
    ("Farmacovigilancia y notificación de RAM", [(3, 9)]),
    ("Eliminación de medicamentos vencidos", [(8, 9)]),
    ("Seminario de título (protocolo CMO)", [(1, 9)]),
]
COLORS_PLAN = ["#4e79a7", "#f28e2b", "#59a14f", "#e15759", "#76b7b2", "#edc948",
               "#b07aa1", "#ff7f9e", "#9c755f", "#8c6bb1", "#17a589", "#c0504d", "#6b7b8c"]
TITLE_PLAN = "Carta Gantt de actividades planificadas (9 semanas)"

nP = len(PLAN_ACTS)
fig, ax = plt.subplots(figsize=(17, 9.8))
for w in range(9):
    for i in range(nP):
        ax.add_patch(Rectangle((w + 0.16, i - 0.32), 0.68, 0.64, facecolor="#f2f5f9",
                     edgecolor="#e3e8ee", linewidth=0.6, zorder=0))
for i, (name, spans) in enumerate(PLAN_ACTS):
    y = nP - i - 1
    ax.add_patch(FancyBboxPatch((-LABELW + 0.15, y - 0.36), LABELW - 0.45, 0.72,
                 boxstyle="round,pad=0.02,rounding_size=0.10",
                 facecolor=COLORS_PLAN[i], edgecolor="white", linewidth=1.6, zorder=3,
                 clip_on=False))
    ax.text(-LABELW + 0.15 + (LABELW - 0.45) / 2, y, textwrap.fill(name, 30),
            ha="center", va="center", fontsize=9.5, color="white", weight="bold",
            zorder=4, clip_on=False)
    for (s, e) in spans:
        for w in range(s, e + 1):
            ax.add_patch(FancyBboxPatch((w - 1 + 0.16, y - 0.32), 0.68, 0.64,
                         boxstyle="round,pad=0.01,rounding_size=0.08",
                         facecolor=COLORS_PLAN[i], edgecolor="white", linewidth=1.5, zorder=3))
ax.set_yticks([])
ax.set_xticks([i + 0.5 for i in range(9)])
ax.set_xticklabels(WEEKS, fontsize=13, weight="bold")
ax.set_xlim(-LABELW, 9); ax.set_ylim(-0.5, nP - 0.5)
secax = ax.secondary_xaxis("bottom")
secax.set_xticks([i + 0.5 for i in range(9)])
secax.set_xticklabels(FECHAS, fontsize=9, color=GREY)
secax.tick_params(length=0, pad=24); secax.spines["bottom"].set_visible(False)
ax.tick_params(length=0)
for spine in ["top", "right", "left"]:
    ax.spines[spine].set_visible(False)
ax.set_title(TITLE_PLAN, fontsize=16, weight="bold", color=NAVY, pad=16)
save(fig, "fig_gantt_planificada")

import os
from PIL import Image
print("Figuras generadas:")
for f in ["fig_organigrama.png", "fig_gantt_planificada.png", "fig_gantt_opcion1.png", "fig_gantt_opcion2.png", "fig_gantt_opcion3.png"]:
    im = Image.open(R + f)
    print(f"  {f}  {im.size}  {os.path.getsize(R+f)//1024} KB")
