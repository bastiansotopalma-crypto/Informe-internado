import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch
import matplotlib.patches as mpatches

plt.rcParams["font.family"] = "DejaVu Sans"

# ----------------------------------------------------------------------------
# FIGURA 1: Organigrama de la Unidad de Farmacia del CESFAM Villa Nonguen
# ----------------------------------------------------------------------------
def box(ax, x, y, w, h, text, fc, ec="#33475b", tc="black", fs=11, bold=True):
    p = FancyBboxPatch((x - w / 2, y - h / 2), w, h,
                       boxstyle="round,pad=0.02,rounding_size=0.06",
                       linewidth=1.6, edgecolor=ec, facecolor=fc)
    ax.add_patch(p)
    ax.text(x, y, text, ha="center", va="center", fontsize=fs,
            color=tc, weight="bold" if bold else "normal", wrap=True)

def line(ax, x1, y1, x2, y2):
    ax.plot([x1, x2], [y1, y2], color="#33475b", linewidth=1.4, zorder=0)

fig, ax = plt.subplots(figsize=(9.2, 6.4))
ax.set_xlim(0, 12); ax.set_ylim(0, 10); ax.axis("off")

# Nivel 0: Direccion del CESFAM
box(ax, 6, 9.2, 5.2, 1.0, "Dirección CESFAM Villa Nonguén\n(COSADES)", "#dbe7f3", fs=11)
# Nivel 1: Direccion Tecnica de Farmacia
box(ax, 6, 7.2, 5.6, 1.1, "Q.F. Directora Técnica de Farmacia\nJennifer Vallejos", "#bcd4ea", fs=11)
# Nivel 2: QF Atencion Farmaceutica
box(ax, 6, 5.2, 5.6, 1.1, "Q.F. de Atención Farmacéutica\nFernanda Torres", "#bcd4ea", fs=11)
# Nivel 3: TENS
tens = ["TENS\nDominic", "TENS\nGénesis", "TENS\nFreire", "TENS\nPaola"]
xs = [1.9, 4.6, 7.3, 10.0]
for x, t in zip(xs, tens):
    box(ax, x, 2.6, 2.2, 1.1, t, "#e8f0d8", fs=10)

# Conectores
line(ax, 6, 8.7, 6, 7.75)
line(ax, 6, 6.65, 6, 5.75)
line(ax, 6, 4.65, 6, 3.55)
line(ax, 1.9, 3.55, 10.0, 3.55)   # barra horizontal
for x in xs:
    line(ax, x, 3.55, x, 3.15)

ax.text(6, 0.7, "Interno de Química y Farmacia integrado a la unidad durante la rotación",
        ha="center", va="center", fontsize=9.5, style="italic", color="#33475b")

plt.tight_layout()
plt.savefig("recursos-informe/fig_organigrama.png", dpi=200, bbox_inches="tight")
plt.close()

# ----------------------------------------------------------------------------
# CARTA GANTT (funcion reutilizable)
# ----------------------------------------------------------------------------
def gantt(activities, filename, title, color):
    weeks = ["S1", "S2", "S3", "S4", "S5", "S6", "S7", "S8", "S9"]
    fig, ax = plt.subplots(figsize=(11, 6.6))
    n = len(activities)
    for i, (name, spans) in enumerate(activities):
        y = n - i - 1
        for (start, end) in spans:
            ax.barh(y, end - start + 1, left=start - 1, height=0.55,
                    color=color, edgecolor="#33475b", linewidth=0.8)
    ax.set_yticks(range(n))
    ax.set_yticklabels([a[0] for a in reversed(activities)], fontsize=9.5)
    ax.set_xticks([i + 0.5 for i in range(9)])
    ax.set_xticklabels(weeks, fontsize=10)
    ax.set_xlim(0, 9)
    ax.set_ylim(-0.6, n - 0.4)
    for x in range(10):
        ax.axvline(x, color="#d0d7de", linewidth=0.6, zorder=0)
    ax.set_xlabel("Semanas de internado (11 de mayo al 10 de julio de 2026)", fontsize=10)
    ax.set_title(title, fontsize=12, weight="bold", pad=12)
    ax.tick_params(length=0)
    for s in ["top", "right", "left"]:
        ax.spines[s].set_visible(False)
    plt.tight_layout()
    plt.savefig(filename, dpi=200, bbox_inches="tight")
    plt.close()

# Actividades PLANIFICADAS
planificadas = [
    ("Inducción y conocimiento del centro y la unidad de farmacia", [(1, 1)]),
    ("Recepción, almacenamiento y gestión de stock", [(1, 2)]),
    ("Fraccionamiento y reenvasado de medicamentos", [(2, 4)]),
    ("Dispensación informada de medicamentos", [(3, 7)]),
    ("Control de indicadores y registros de despacho", [(3, 8)]),
    ("Atención farmacéutica y visitas domiciliarias", [(2, 9)]),
    ("Educación sanitaria y uso racional de medicamentos", [(6, 7)]),
    ("Inventario y ordenamiento de bodega", [(5, 5)]),
    ("Participación en Comité de Farmacia y Terapéutica", [(4, 4)]),
    ("Actividades de farmacovigilancia", [(5, 6)]),
    ("Desarrollo del trabajo de investigación (protocolo CMO)", [(1, 9)]),
    ("Reuniones de supervisión y elaboración del informe", [(6, 6), (9, 9)]),
]
gantt(planificadas, "recursos-informe/fig_gantt_planificadas.png",
      "Carta Gantt de actividades planificadas", "#7fa8d0")

# Actividades DESARROLLADAS
desarrolladas = [
    ("Inducción y conocimiento del centro y la unidad de farmacia", [(1, 1)]),
    ("Recepción, almacenamiento y gestión de stock", [(1, 2), (4, 4)]),
    ("Fraccionamiento y reenvasado de medicamentos", [(1, 9)]),
    ("Dispensación informada de medicamentos", [(3, 9)]),
    ("Control de indicadores y registros de despacho", [(3, 3), (4, 5), (8, 8)]),
    ("Atención farmacéutica y visitas domiciliarias", [(2, 9)]),
    ("Educación sanitaria y uso racional de medicamentos", [(7, 7)]),
    ("Inventario y ordenamiento de bodega", [(5, 5)]),
    ("Apoyo en supervisiones del Servicio de Salud", [(4, 5)]),
    ("Vinculación con el medio (apoyo comunitario)", [(3, 3)]),
    ("Desarrollo del trabajo de investigación (protocolo CMO)", [(1, 9)]),
    ("Reuniones de supervisión y elaboración del informe", [(6, 6), (9, 9)]),
]
gantt(desarrolladas, "recursos-informe/fig_gantt_desarrolladas.png",
      "Carta Gantt de actividades desarrolladas", "#5a9367")

print("Figuras generadas:")
import os
for f in ["recursos-informe/fig_organigrama.png", "recursos-informe/fig_gantt_planificadas.png", "recursos-informe/fig_gantt_desarrolladas.png"]:
    print(" ", f, os.path.getsize(f), "bytes")
