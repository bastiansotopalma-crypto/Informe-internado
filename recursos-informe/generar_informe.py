# -*- coding: utf-8 -*-
"""Genera el Informe de Internado en Farmacia Asistencial y APS (CESFAM Villa Nonguen)
segun el formato del Instructivo de Internados USS 2026."""

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FONT = "Arial"

doc = Document()

# ---------------------------------------------------------------- base styles
normal = doc.styles["Normal"]
normal.font.name = FONT
normal.font.size = Pt(12)
normal.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
pf = normal.paragraph_format
pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
pf.space_after = Pt(6)
pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

def style_heading(name, size, bold=True, upper=True, align=WD_ALIGN_PARAGRAPH.CENTER,
                  color=RGBColor(0, 0, 0), space_before=18, space_after=12):
    st = doc.styles[name]
    st.font.name = FONT
    st.font.size = Pt(size)
    st.font.bold = bold
    st.font.color.rgb = color
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    st.paragraph_format.alignment = align
    st.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    st.paragraph_format.space_before = Pt(space_before)
    st.paragraph_format.space_after = Pt(space_after)
    st.paragraph_format.keep_with_next = True

style_heading("Heading 1", 12, upper=True, align=WD_ALIGN_PARAGRAPH.CENTER)
style_heading("Heading 2", 12, upper=False, align=WD_ALIGN_PARAGRAPH.LEFT,
              space_before=12, space_after=6)
style_heading("Heading 3", 12, upper=False, align=WD_ALIGN_PARAGRAPH.LEFT,
              space_before=8, space_after=4)
cap = doc.styles["Caption"]
cap.font.name = FONT
cap.font.size = Pt(10)
cap.font.italic = True
cap.font.bold = False
cap.font.color.rgb = RGBColor(0, 0, 0)
cap.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
cap.paragraph_format.space_before = Pt(4)
cap.paragraph_format.space_after = Pt(10)

# ------------------------------------------------------------- margin helper
def set_margins(section):
    section.left_margin = Cm(4)
    section.right_margin = Cm(2.5)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)

sec0 = doc.sections[0]
set_margins(sec0)

# ------------------------------------------------------------- field helpers
def add_field(paragraph, instr, placeholder=""):
    r = paragraph.add_run()
    fb = OxmlElement("w:fldChar"); fb.set(qn("w:fldCharType"), "begin")
    r._r.append(fb)
    r2 = paragraph.add_run()
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve")
    it.text = instr
    r2._r.append(it)
    r3 = paragraph.add_run()
    fs = OxmlElement("w:fldChar"); fs.set(qn("w:fldCharType"), "separate")
    r3._r.append(fs)
    if placeholder:
        paragraph.add_run(placeholder)
    r5 = paragraph.add_run()
    fe = OxmlElement("w:fldChar"); fe.set(qn("w:fldCharType"), "end")
    r5._r.append(fe)

def add_page_number(paragraph):
    add_field(paragraph, " PAGE ", "1")

def footer_page_number(section):
    section.footer.is_linked_to_previous = False
    p = section.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in list(p.runs):
        run.text = ""
    add_page_number(p)
    for r in p.runs:
        r.font.name = FONT; r.font.size = Pt(11)

def set_pgnum(section, fmt=None, start=None):
    sectPr = section._sectPr
    pg = sectPr.find(qn("w:pgNumType"))
    if pg is None:
        pg = OxmlElement("w:pgNumType"); sectPr.append(pg)
    if fmt:
        pg.set(qn("w:fmt"), fmt)
    if start is not None:
        pg.set(qn("w:start"), str(start))

# ------------------------------------------------------------- content helpers
def para(text, style=None, align=None, bold=False, italic=False, size=None,
         space_after=None, space_before=None):
    p = doc.add_paragraph(style=style)
    run = p.add_run(text)
    run.bold = bold; run.italic = italic
    if size: run.font.size = Pt(size)
    if align is not None: p.alignment = align
    if space_after is not None: p.paragraph_format.space_after = Pt(space_after)
    if space_before is not None: p.paragraph_format.space_before = Pt(space_before)
    return p

def bullet(text):
    p = doc.add_paragraph(style="List Bullet")
    p.add_run(text)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    for r in p.runs:
        r.font.name = FONT; r.font.size = Pt(12)
    return p

def numbered(text):
    p = doc.add_paragraph(style="List Number")
    p.add_run(text)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    for r in p.runs:
        r.font.name = FONT; r.font.size = Pt(12)
    return p

def heading(text, level=1):
    return doc.add_heading(text, level=level)

def caption(label, text):
    p = doc.add_paragraph(style="Caption")
    p.add_run(f"{label} ")
    add_field(p, f" SEQ {label} \\* ARABIC ", "1")
    p.add_run(f". {text}")
    return p

def add_image(path, width_cm):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    p.add_run().add_picture(path, width=Cm(width_cm))
    return p

def blank(n=1):
    for _ in range(n):
        doc.add_paragraph()

def make_table(headers, rows, col_widths=None, font_size=10):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].paragraphs[0].text = ""
        run = hdr[i].paragraphs[0].add_run(h)
        run.bold = True; run.font.name = FONT; run.font.size = Pt(font_size)
        hdr[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        shd = OxmlElement("w:shd"); shd.set(qn("w:fill"), "D9E2F3")
        hdr[i]._tc.get_or_add_tcPr().append(shd)
    for row in rows:
        cells = t.add_row().cells
        for i, val in enumerate(row):
            cells[i].paragraphs[0].text = ""
            run = cells[i].paragraphs[0].add_run(val)
            run.font.name = FONT; run.font.size = Pt(font_size)
            cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
            cells[i].paragraphs[0].paragraph_format.space_after = Pt(2)
    if col_widths:
        for i, w in enumerate(col_widths):
            for row in t.rows:
                row.cells[i].width = Cm(w)
    return t

# =====================================================================
# PORTADA  (Seccion 1, sin numero de pagina)
# =====================================================================
def centered(text, size=12, bold=False, italic=False, space_after=6, space_before=0):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    r = p.add_run(text); r.bold = bold; r.italic = italic
    r.font.name = FONT; r.font.size = Pt(size)
    return p

centered("[Insertar logo institucional actualizado de la USS]", size=11, italic=True, space_after=18, space_before=12)
centered("UNIVERSIDAD SAN SEBASTIÁN", size=14, bold=True)
centered("FACULTAD DE CIENCIAS", size=14, bold=True)
centered("ESCUELA DE QUÍMICA Y FARMACIA", size=14, bold=True)
centered("SEDE CONCEPCIÓN", size=14, bold=True, space_after=48)
centered("Informe de Internado en Farmacia Asistencial y Atención Primaria de Salud",
         size=12, bold=True, space_after=40)
centered("Lugar de Internado", size=12, bold=False, space_after=2)
centered("Centro de Salud Familiar Villa Nonguén, Concepción", size=14, bold=True, space_after=44)
centered("Estudiante", size=12, space_after=2)
centered("Bastián Alonso Espinoza Palma", size=12, bold=True, space_after=44)
centered("Docente Tutor: Q.F. de la Unidad de Farmacia, CESFAM Villa Nonguén", size=11, space_after=2)
centered("Docente Supervisor: Escuela de Química y Farmacia, USS", size=11, space_after=40)
centered("Concepción, Chile", size=12, space_after=2)
centered("Julio de 2026", size=12)

# =====================================================================
# SECCION 2: paginas preliminares  (numeros romanos)
# =====================================================================
doc.add_section(WD_SECTION.NEW_PAGE)
sec1 = doc.sections[1]
set_margins(sec1)
sec1.footer.is_linked_to_previous = False
# portada sin numero
sec0.footer.is_linked_to_previous = False
sec0.footer.paragraphs[0].text = ""
footer_page_number(sec1)
set_pgnum(sec1, fmt="lowerRoman", start=2)

heading("TABLA DE CONTENIDOS")
p = doc.add_paragraph()
add_field(p, ' TOC \\o "1-3" \\h \\z \\u ',
          "Actualice este campo en Word (clic derecho sobre la tabla y Actualizar campos).")

doc.add_page_break()
heading("ÍNDICE DE TABLAS")
p = doc.add_paragraph()
add_field(p, ' TOC \\h \\z \\c "Tabla" ',
          "Actualice este campo en Word (clic derecho y Actualizar campos).")

heading("ÍNDICE DE FIGURAS")
p = doc.add_paragraph()
add_field(p, ' TOC \\h \\z \\c "Figura" ',
          "Actualice este campo en Word (clic derecho y Actualizar campos).")

doc.add_page_break()
heading("ABREVIATURAS")
abrev = [
    ("APS", "Atención Primaria de Salud"),
    ("CESFAM", "Centro de Salud Familiar"),
    ("CENABAST", "Central de Abastecimiento del Sistema Nacional de Servicios de Salud"),
    ("CMO", "Capacidad, Motivación y Oportunidad (modelo de atención farmacéutica)"),
    ("COSADES", "Corporación de Salud y Desarrollo Social"),
    ("DM2", "Diabetes Mellitus tipo 2"),
    ("DLP", "Dislipidemia"),
    ("FOFAR", "Fondo de Farmacia"),
    ("GES", "Garantías Explícitas en Salud"),
    ("HTA", "Hipertensión Arterial"),
    ("MINSAL", "Ministerio de Salud de Chile"),
    ("PRM", "Problemas Relacionados con los Medicamentos"),
    ("PSCV", "Programa de Salud Cardiovascular"),
    ("Q.F.", "Químico Farmacéutico"),
    ("RAM", "Reacción Adversa a Medicamentos"),
    ("RNM", "Resultados Negativos asociados a la Medicación"),
    ("TENS", "Técnico en Enfermería de Nivel Superior"),
    ("URM", "Uso Racional de Medicamentos"),
]
for a, d in abrev:
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(f"{a}: "); r.bold = True; r.font.name = FONT; r.font.size = Pt(12)
    r2 = p.add_run(d); r2.font.name = FONT; r2.font.size = Pt(12)

# =====================================================================
# SECCION 3: cuerpo  (numeros arabigos, inicia en 1)
# =====================================================================
doc.add_section(WD_SECTION.NEW_PAGE)
sec2 = doc.sections[2]
set_margins(sec2)
footer_page_number(sec2)
set_pgnum(sec2, fmt="decimal", start=1)

# --------------------------------------------------------------- INTRODUCCION
heading("INTRODUCCIÓN")

para("El presente informe reúne la experiencia desarrollada durante el Internado en "
     "Farmacia Asistencial y Atención Primaria de Salud, realizado en la unidad de "
     "farmacia del Centro de Salud Familiar (CESFAM) Villa Nonguén, en la comuna de "
     "Concepción, Región del Biobío. El internado corresponde a una actividad "
     "obligatoria del plan de estudios de la carrera de Química y Farmacia, cuyo "
     "propósito es que el estudiante adquiera experiencia directa del quehacer "
     "profesional del químico farmacéutico en un establecimiento de ejercicio real, "
     "bajo la supervisión de un profesional del área. La atención primaria constituye "
     "la puerta de entrada al sistema público de salud y, dentro de ella, la farmacia "
     "cumple un rol central en el acceso a los medicamentos y en el uso seguro y "
     "racional de la farmacoterapia por parte de la población (Ministerio de Salud de "
     "Chile, 2018b).")

heading("Aspectos generales del centro de internado", level=2)
para("El CESFAM Villa Nonguén es un establecimiento del nivel primario de atención, "
     "ubicado en Río Loa 1397, en el sector de Nonguén, en la periferia oriente de "
     "Concepción, e integrado a la red asistencial del Servicio de Salud Concepción. "
     "Su origen se remonta a 1987, cuando profesionales de la entonces Octava Región, "
     "con el apoyo de la organización italiana CESTAS y financiamiento de la "
     "cooperación del Gobierno de Italia y de la Unión Europea, elaboraron un proyecto "
     "de salud comunitaria para el sector. El establecimiento inició sus funciones "
     "asistenciales en 1991 y, en 1993, fue seleccionado por el Ministerio de Salud "
     "como uno de los centros piloto para implementar el Modelo de Salud Familiar y el "
     "financiamiento per cápita en el país (Diario Concepción, 2017). Desde entonces "
     "funciona bajo la modalidad de administración delegada, a cargo de la Corporación "
     "de Salud y Desarrollo Social (COSADES), lo que le permite formar parte de la red "
     "pública manteniendo cierta autonomía en su gestión (COSADES, s.f.).")
para("El centro atiende a una población inscrita validada cercana a las 16.760 "
     "personas, organizada en torno a unas 5.560 familias, con un promedio de 3,1 "
     "integrantes por grupo familiar, y entrega prestaciones a todos los grupos del "
     "ciclo vital, desde la gestación hasta la persona mayor. Para ordenar la atención, "
     "el territorio se divide en cuatro sectores identificados por colores (verde, "
     "azul, café y blanco), lo que facilita la planificación de los controles, las "
     "visitas domiciliarias y las actividades comunitarias según el lugar de residencia "
     "de los usuarios. La dotación total del establecimiento alcanza los 122 "
     "funcionarios distribuidos en distintas profesiones y oficios, y contempla una "
     "jornada diurna y otra vespertina. La Tabla 1 resume los principales antecedentes "
     "del centro y la Tabla 2 detalla su dotación de personal.")

caption("Tabla", "Antecedentes generales del CESFAM Villa Nonguén.")
make_table(
    ["Antecedente", "Detalle"],
    [
        ["Dependencia administrativa", "Administración delegada, COSADES, Servicio de Salud Concepción"],
        ["Nivel de atención", "Primario (Atención Primaria de Salud)"],
        ["Dirección", "Río Loa 1397, Villa Nonguén, Concepción, Región del Biobío"],
        ["Población inscrita validada", "16.760 personas aproximadamente"],
        ["Familias registradas", "5.560 (promedio de 3,1 integrantes por familia)"],
        ["Sectorización", "Cuatro sectores: verde, azul, café y blanco"],
        ["Dotación total", "122 funcionarios (jornada diurna y vespertina)"],
    ],
    col_widths=[5.5, 9.5])

caption("Tabla", "Dotación de personal del CESFAM Villa Nonguén según estamento.")
make_table(
    ["Estamento", "N.º", "Estamento", "N.º"],
    [
        ["Médico", "14", "Enfermera/o", "11"],
        ["Odontólogo", "7", "Matrona", "8"],
        ["Químico farmacéutico", "3", "Kinesiólogo", "3"],
        ["Nutricionista", "3", "Psicólogo", "6"],
        ["Trabajador social", "5", "Podólogo", "1"],
        ["Técnico en enfermería", "25", "Secretaria", "11"],
        ["Auxiliar de servicios", "8", "Otros (administrativos y de apoyo)", "10"],
    ],
    col_widths=[5.6, 1.9, 5.6, 1.9])
para("Nota: La dotación total corresponde a 122 funcionarios. Elaboración a partir de "
     "los antecedentes entregados por el establecimiento.", size=10, italic=True,
     space_after=10)

heading("Organigrama y distribución de la unidad de farmacia", level=2)
para("La unidad de farmacia se organiza bajo la responsabilidad de un químico "
     "farmacéutico que ejerce la dirección técnica y responde por el funcionamiento "
     "general de la unidad, el cumplimiento normativo y la gestión del arsenal "
     "farmacoterapéutico. Un segundo químico farmacéutico se dedica de manera "
     "preferente a la atención farmacéutica, tanto en box como en las visitas "
     "domiciliarias del Programa de Salud Cardiovascular, y un equipo permanente de "
     "cuatro técnicos en enfermería de nivel superior (TENS) apoya las tareas de "
     "recepción, almacenamiento, fraccionamiento, reenvasado y dispensación. El "
     "interno de Química y Farmacia se integró a este equipo durante toda la rotación. "
     "La Figura 1 presenta el organigrama de la unidad.")
add_image("recursos-informe/fig_organigrama.png", 13.5)
caption("Figura", "Organigrama de la unidad de farmacia del CESFAM Villa Nonguén.")

heading("Servicios que brinda el centro", level=2)
para("El CESFAM Villa Nonguén entrega una atención integral y continua a lo largo del "
     "curso de vida, con acciones de promoción, prevención, tratamiento y "
     "rehabilitación. Su cartera de prestaciones incluye, entre otros, el Programa de "
     "Salud de la Mujer, el Programa Nacional de Salud de la Infancia, el Programa de "
     "Salud Integral del Adolescente, el Programa de Salud Cardiovascular, el Programa "
     "de Salud Mental, el Programa Nacional de Salud Integral de Personas Mayores, el "
     "Programa de Salud Familiar y el Programa VIH/ITS. A ello se suman prestaciones de "
     "procedimientos, vacunatorio, toma de muestras, salud dental, controles de salud, "
     "exámenes preventivos y atención domiciliaria, junto con la entrega de alimentos "
     "de los programas PNAC y PACAM en el hall central del establecimiento.")
para("La unidad de farmacia participa de manera transversal en estos programas a "
     "través de la entrega de medicamentos, el fraccionamiento y reenvasado según las "
     "indicaciones, el control del arsenal y de los medicamentos sujetos a control "
     "legal, y la atención farmacéutica orientada al uso seguro y racional de la "
     "terapia. Dentro del Programa de Salud Cardiovascular, el químico farmacéutico "
     "acompaña a los usuarios en box y en su domicilio, con seguimiento "
     "farmacoterapéutico, educación y revisión de la medicación, y participa en la "
     "gestión del Fondo de Farmacia (FOFAR), que financia su presencia en la atención "
     "primaria y el acceso a los medicamentos para la hipertensión, la diabetes y la "
     "dislipidemia (Ministerio de Salud de Chile, 2020).")

heading("Terminología y definiciones", level=2)
para("Para facilitar la lectura del informe se precisan a continuación algunos "
     "términos propios del trabajo desarrollado en la unidad de farmacia.")
defs = [
    ("Atención farmacéutica", "participación del farmacéutico en el cuidado de la "
     "persona para aprovechar mejor sus medicamentos y mejorar sus resultados en "
     "salud, cooperando con el paciente y el equipo en el diseño, la ejecución y el "
     "seguimiento de un plan terapéutico (Bonal et al., 2003)."),
    ("Dispensación informada", "acto profesional de entrega del medicamento "
     "acompañado de la información necesaria para su uso correcto, seguro y efectivo."),
    ("Fraccionamiento y reenvasado", "proceso mediante el cual se acondiciona el "
     "medicamento en la cantidad que corresponde a cada paciente, rotulando el envase "
     "con el nombre, la indicación, la cantidad, la fecha de vencimiento y la serie o "
     "lote, de manera que quede asegurada la trazabilidad."),
    ("Seguimiento farmacoterapéutico", "práctica clínica en la que el farmacéutico "
     "detecta, previene y resuelve los problemas relacionados con los medicamentos de "
     "forma continua y documentada (Ministerio de Salud de Chile, 2018a)."),
    ("Problemas relacionados con los medicamentos (PRM)", "situaciones que, en el "
     "proceso de uso de los medicamentos, pueden causar o causan un resultado negativo "
     "asociado a la medicación."),
    ("Resultados negativos asociados a la medicación (RNM)", "resultados en la salud "
     "del paciente que no se corresponden con los objetivos de la farmacoterapia y que "
     "se clasifican según necesidad, efectividad y seguridad (Comité de Consenso de "
     "Granada, 2007)."),
    ("Modelo CMO", "marco de atención farmacéutica que ordena el trabajo en torno a "
     "tres pilares, Capacidad, Motivación y Oportunidad, orientado a priorizar a los "
     "pacientes más complejos, trabajar la adherencia y dar continuidad al seguimiento "
     "(Calleja Hernández y Morillo Verdugo, 2016)."),
]
for term, d in defs:
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    p.paragraph_format.space_after = Pt(4)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run(f"{term}: "); r.bold = True; r.font.name = FONT; r.font.size = Pt(12)
    r2 = p.add_run(d); r2.font.name = FONT; r2.font.size = Pt(12)

heading("Objetivos", level=2)
heading("Objetivo general", level=3)
para("Desarrollar las competencias profesionales del químico farmacéutico en el "
     "ámbito de la farmacia asistencial y la atención primaria de salud, mediante la "
     "integración al equipo de la unidad de farmacia del CESFAM Villa Nonguén y la "
     "participación en sus procesos técnicos y clínicos durante el periodo de "
     "internado.")
heading("Objetivos específicos", level=3)
numbered("Conocer la estructura organizacional, las funciones del equipo y los "
         "procesos administrativos y técnicos de la unidad de farmacia del CESFAM "
         "Villa Nonguén.")
numbered("Participar en los procesos de recepción, almacenamiento, gestión de stock, "
         "fraccionamiento, reenvasado y dispensación de medicamentos, aplicando las "
         "buenas prácticas y resguardando la trazabilidad.")
numbered("Desarrollar actividades de atención farmacéutica, seguimiento "
         "farmacoterapéutico y educación sanitaria, incluidas las visitas "
         "domiciliarias y la promoción del uso racional de los medicamentos.")
numbered("Diseñar un protocolo de atención farmacéutica domiciliaria basado en el "
         "Modelo CMO para los pacientes del Programa de Salud Cardiovascular del "
         "CESFAM Villa Nonguén, como trabajo de investigación del internado.")

# --------------------------------------------------------- ACTIVIDADES REALIZADAS
heading("ACTIVIDADES REALIZADAS")

heading("Cronograma de actividades planificadas", level=2)
para("Al inicio del internado se acordó con el equipo de farmacia un plan de trabajo "
     "que ordenó la rotación por las distintas tareas de la unidad y reservó tiempo "
     "para el desarrollo del trabajo de investigación. La Figura 2 presenta la carta "
     "Gantt con las actividades planificadas a lo largo de las nueve semanas de "
     "internado, comprendidas entre el 11 de mayo y el 10 de julio de 2026.")
add_image("recursos-informe/fig_gantt_planificadas.png", 15.0)
caption("Figura", "Carta Gantt de actividades planificadas del internado.")

heading("Cronograma de actividades desarrolladas", level=2)
para("La Figura 3 muestra las actividades efectivamente desarrolladas. En términos "
     "generales, la ejecución siguió lo planificado, con algunos ajustes propios de "
     "la dinámica del establecimiento. El fraccionamiento, la dispensación y la "
     "atención farmacéutica se extendieron durante casi todo el periodo, ya que "
     "constituyen tareas habituales de la unidad, y se sumaron actividades no "
     "previstas al inicio, como el apoyo durante las supervisiones del Servicio de "
     "Salud y una situación de vinculación con el medio surgida en terreno.")
add_image("recursos-informe/fig_gantt_desarrolladas.png", 15.0)
caption("Figura", "Carta Gantt de actividades desarrolladas del internado.")
para("Respecto de las actividades no realizadas, la participación en el Comité de "
     "Farmacia y Terapéutica no pudo concretarse porque el comité no sesionó durante "
     "el periodo de la rotación, de modo que su contenido se abordó de manera indirecta "
     "a través de la revisión de los criterios de selección del arsenal y de los "
     "cambios de medicamentos informados por el Servicio de Salud. Las actividades "
     "formales de farmacovigilancia se limitaron a conocer el circuito de notificación "
     "y el algoritmo utilizado, sin que se presentara un caso que ameritara notificar "
     "durante la estadía.")

heading("Descripción de las actividades realizadas", level=2)

heading("Inducción y conocimiento del centro y de la unidad de farmacia", level=3)
para("Durante la primera semana se realizó la presentación al equipo de trabajo y se "
     "recorrieron las dependencias del CESFAM y de la unidad de farmacia. El equipo "
     "explicó la organización de la unidad, la ubicación física de los medicamentos, "
     "el flujo de trabajo diario y el uso de los sistemas de registro. Se revisaron "
     "los protocolos internos del establecimiento, incluidos los de emergencia, como "
     "el código rojo y la reanimación cardiopulmonar ante un paro cardiorrespiratorio, "
     "y el manejo inicial de situaciones críticas como la reacción alérgica, la "
     "hemorragia masiva y la convulsión. También se conocieron los programas "
     "computacionales de apoyo y la ficha clínica electrónica del establecimiento "
     "(SINET Sur), utilizada para consultar antecedentes y dejar constancia de las "
     "atenciones.")

heading("Recepción, almacenamiento y gestión del stock", level=3)
para("Se participó en la recepción y el almacenamiento de los medicamentos que llegan "
     "a la unidad, que se adquieren principalmente a través de la Central de "
     "Abastecimiento (CENABAST) y que se ordenan resguardando las condiciones de "
     "conservación. Se colaboró en la gestión del stock, con especial atención al "
     "control de los medicamentos próximos a vencer, que se marcaban y se controlaban "
     "para evitar pérdidas. Cuando un medicamento estaba por vencer y no alcanzaba a "
     "utilizarse, se gestionaban préstamos o canjes con otros centros de la red, y las "
     "existencias que no se podían aprovechar se reportaban y se redistribuían hacia "
     "otros CESFAM, de manera de dar un uso eficiente a los recursos y reducir las "
     "mermas. Se conoció además el proceso de programación anual de las necesidades de "
     "medicamentos, que se planifica según la demanda y las metas del establecimiento.")

heading("Fraccionamiento y reenvasado de medicamentos", level=3)
para("El fraccionamiento y el reenvasado fueron tareas frecuentes a lo largo del "
     "internado. El trabajo consistió en acondicionar los medicamentos en la cantidad "
     "correspondiente a cada paciente y preparar las bolsas de tratamiento, rotulando "
     "cada envase con el nombre del medicamento, la indicación, la cantidad, la fecha "
     "de vencimiento y la serie o lote, de manera de asegurar la trazabilidad del "
     "producto hasta el usuario. Este proceso permitió comprender la importancia del "
     "rotulado correcto y del resguardo de la información para la seguridad del "
     "paciente, sobre todo en personas con polifarmacia que retiran varios "
     "medicamentos a la vez.")

heading("Dispensación informada de medicamentos", level=3)
para("Se apoyó la dispensación de medicamentos a los usuarios, con lo que se "
     "reforzaron el reconocimiento de los principios activos, las presentaciones y su "
     "ubicación dentro de la unidad, y las buenas prácticas de entrega. La "
     "dispensación se acompañó de la información necesaria para el uso correcto de la "
     "terapia, verificando la indicación y resolviendo dudas de los pacientes. También "
     "se conocieron los criterios de manejo de los medicamentos sujetos a control legal "
     "y la forma de registrar su entrega.")

heading("Atención farmacéutica y visitas domiciliarias", level=3)
para("La atención farmacéutica fue una de las actividades más significativas del "
     "internado. Se acompañó al químico farmacéutico en la atención de los usuarios "
     "del Programa de Salud Cardiovascular, tanto en box como en el domicilio, y se "
     "participó en la entrega de medicamentos a domicilio para pacientes con "
     "dificultades para acudir al centro. En estas visitas se revisaba la medicación "
     "que el paciente realmente utilizaba, se conciliaba con la que tenía indicada, se "
     "resolvían dudas sobre horarios y formas de administración, y se reforzaba la "
     "adherencia. En una de las salidas a terreno se prestó apoyo a un adulto mayor "
     "para regresar a su hogar, situación que reflejó el componente humano y "
     "comunitario del trabajo en la atención primaria. Durante una jornada vespertina "
     "se observó, además, el funcionamiento de la farmacia en el horario de la tarde y "
     "el procedimiento de cierre de la unidad.")

heading("Control de indicadores y registros de despacho", level=3)
para("Se colaboró en el registro y el control de los indicadores de la unidad. Entre "
     "ellos, se trabajó en el indicador de porcentaje de despacho de medicamentos "
     "realizado según el protocolo de atención farmacéutica, desagregado por sector, y "
     "se completó la plantilla mensual de despacho correspondiente al mes de mayo. "
     "Estas tareas permitieron comprender cómo la unidad mide su actividad, justifica "
     "el uso de los recursos y respalda la toma de decisiones y la rendición ante el "
     "Servicio de Salud.")

heading("Inventario y ordenamiento de bodega", level=3)
para("Se participó en las tareas de inventario y en el ordenamiento de la bodega de "
     "la unidad, cuidando que las existencias quedaran correctamente ubicadas y que no "
     "permanecieran cajas en el piso, tanto por seguridad como por buenas prácticas de "
     "almacenamiento. Esta actividad reforzó la importancia del orden y de la "
     "conservación adecuada de los medicamentos.")

heading("Apoyo en supervisiones del Servicio de Salud", level=3)
para("Durante el internado el establecimiento recibió supervisiones e inspecciones "
     "del Servicio de Salud. Se acompañó al equipo en la preparación y atención de "
     "estas instancias, lo que permitió conocer los aspectos que se evalúan en la "
     "unidad de farmacia y la relevancia de mantener los registros y los procesos al "
     "día. En paralelo, en algunas de estas jornadas se dispuso de tiempo para avanzar "
     "en el trabajo de investigación.")

heading("Educación sanitaria y promoción del uso racional de medicamentos", level=3)
para("Se participó en un taller grupal de uso racional de medicamentos dirigido a los "
     "usuarios del CESFAM. La actividad se realizó junto al químico farmacéutico y "
     "consistió en conversar con los pacientes, entregar material educativo en forma "
     "de trípticos y desarrollar una dinámica participativa con paletas de respuesta, "
     "en la que los asistentes respondían situaciones sobre el uso correcto de los "
     "medicamentos y recibían pequeños incentivos por participar. La instancia "
     "favoreció la educación sanitaria y permitió acercar el rol del farmacéutico a la "
     "comunidad.")

heading("Trabajo de investigación del internado", level=2)
para("Como parte del internado, y bajo la supervisión del químico farmacéutico tutor, "
     "se desarrolló un trabajo de investigación consistente en el diseño de un "
     "protocolo de atención farmacéutica domiciliaria basado en el Modelo CMO para los "
     "pacientes del Programa de Salud Cardiovascular del CESFAM Villa Nonguén. El "
     "trabajo se enmarca en la necesidad de ordenar y estandarizar la visita "
     "domiciliaria del farmacéutico, que en el centro se realizaba con criterio "
     "clínico y registro en la ficha, pero sin un procedimiento escrito y común que la "
     "hiciera trazable y comparable entre profesionales.")

heading("Tipo de estudio y periodo", level=3)
para("Correspondió a un trabajo de desarrollo metodológico. No se realizó un "
     "experimento ni se recogieron datos clínicos de pacientes con fines de "
     "investigación, sino que se diseñó un protocolo a partir de la revisión de la "
     "literatura, del análisis de la normativa del Ministerio de Salud y de la "
     "adaptación de herramientas ya utilizadas en la atención primaria, ajustadas a lo "
     "observado en el propio centro. El trabajo se desarrolló durante los meses del "
     "internado, entre mayo y julio de 2026, en el CESFAM Villa Nonguén y su "
     "territorio.")

heading("Diagnóstico, población y criterios", level=3)
para("En una primera etapa se diagnosticaron las brechas del programa de visitas "
     "domiciliarias vigente, comparando la práctica del centro con las orientaciones "
     "técnicas del Programa de Salud Cardiovascular y del Fondo de Farmacia, la Guía "
     "de Atención Farmacéutica y Seguimiento Farmacoterapéutico en APS y los tres "
     "pilares del Modelo CMO. La Tabla 3 resume las brechas detectadas. El protocolo "
     "se orientó a las personas adultas y mayores del Programa de Salud Cardiovascular, "
     "con prioridad en quienes toman muchos medicamentos, dependen de un cuidador, se "
     "descompensan con frecuencia o tienen dificultades para acudir al centro. Se "
     "definieron criterios de inclusión, usuarios del programa mayores de 18 años, "
     "descompensados o con control inestable y en seguimiento activo, y criterios de "
     "exclusión, como las barreras de comunicación que impiden la entrevista, el "
     "deterioro cognitivo importante sin cuidador de apoyo, la residencia fuera del "
     "territorio y la inasistencia a controles por más de seis meses.")

caption("Tabla", "Brechas del programa de visitas domiciliarias según los pilares del Modelo CMO.")
make_table(
    ["Pilar del Modelo CMO", "Lo que se espera", "Lo observado", "Brecha"],
    [
        ["Capacidad", "Priorizar la visita según el riesgo y la complejidad del paciente.",
         "El ingreso depende de la derivación y del tiempo disponible (unas 25 visitas al mes).",
         "No hay un criterio de estratificación de riesgo propio del farmacéutico."],
        ["Motivación", "Abordar la adherencia con entrevista motivacional estructurada.",
         "Muchas derivaciones son por adherencia o desconocimiento del tratamiento.",
         "La educación y el trabajo de barreras dependen de cada visita."],
        ["Oportunidad", "Registro estructurado y continuidad del seguimiento.",
         "Registro en la ficha clínica electrónica (SINET Sur); frecuencia según disponibilidad.",
         "Falta un registro propio de la atención farmacéutica y una pauta común de seguimiento."],
    ],
    col_widths=[3.0, 4.0, 4.0, 4.0], font_size=9)

heading("Intervenciones, instrumentos y fuentes de información", level=3)
para("El protocolo tradujo los tres pilares del Modelo CMO en pasos concretos. La "
     "Capacidad se abordó mediante la priorización de los pacientes y la revisión del "
     "botiquín del hogar para conciliar la medicación; la Motivación, mediante la "
     "entrevista motivacional orientada a trabajar las creencias y las barreras de "
     "adherencia; y la Oportunidad, mediante la continuidad del seguimiento por vía "
     "telefónica y el refuerzo del vínculo con la farmacia. Para su aplicación se "
     "seleccionaron y adaptaron instrumentos sencillos, compatibles con la rutina del "
     "CESFAM, entre ellos el test de Morisky-Green-Levine para la adherencia, la "
     "clasificación de los PRM y RNM según el Consenso de Granada, una ficha de "
     "conciliación, una ficha de seguimiento farmacoterapéutico, una hoja con el "
     "esquema de horarios y un protocolo de llamada telefónica para agendar la visita. "
     "La revisión bibliográfica que sustentó estas decisiones se realizó en bases de "
     "datos como PubMed, SciELO y Google Scholar, junto con la normativa y las guías "
     "del Ministerio de Salud. Como apoyo para la información de medicamentos se "
     "consultaron fuentes de referencia como UpToDate, Medscape y Drugs.com.")

heading("Consideraciones éticas", level=3)
para("Al tratarse del diseño de un protocolo y no de un estudio con pacientes, el "
     "trabajo no contempló intervenciones experimentales ni la recolección de datos "
     "clínicos sensibles con fines de investigación. Su elaboración se basó en la "
     "literatura, en la normativa vigente y en la observación de los procesos "
     "habituales del centro. Se resguardaron los principios de confidencialidad y de "
     "uso responsable de la información clínica, sin registrar datos que permitieran "
     "identificar a pacientes ni modificar tratamientos fuera del trabajo habitual del "
     "equipo de salud. La propuesta quedó a disposición de los químicos farmacéuticos "
     "del establecimiento para su revisión y eventual incorporación a la práctica.")

# --------------------------------------------------- DISCUSION Y CONCLUSIONES
heading("DISCUSIÓN Y CONCLUSIONES")

heading("Alcances de las actividades realizadas", level=2)
para("El internado permitió conocer de manera directa el funcionamiento de una "
     "unidad de farmacia de atención primaria y participar en la mayoría de sus "
     "procesos, desde la gestión del arsenal hasta la atención farmacéutica en el "
     "domicilio. Las actividades realizadas cubrieron los objetivos planteados y "
     "abarcaron tanto tareas de gestión y logística como tareas clínicas y educativas, "
     "lo que entregó una visión completa del rol del químico farmacéutico en este "
     "nivel de atención. La experiencia confirmó que el trabajo del farmacéutico en la "
     "atención primaria no se agota en la entrega del medicamento, sino que se extiende "
     "al uso seguro y racional de la terapia, al seguimiento del paciente y a la "
     "vinculación con la comunidad.")

heading("Cumplimiento de los objetivos", level=2)
para("El primer objetivo específico se cumplió al conocer la estructura, las "
     "funciones del equipo y los procesos de la unidad de farmacia, tal como se "
     "describe en la introducción y en el organigrama. El segundo se alcanzó mediante "
     "la participación sostenida en la recepción, el almacenamiento, la gestión de "
     "stock, el fraccionamiento, el reenvasado y la dispensación, resguardando la "
     "trazabilidad a través del rotulado. El tercero se desarrolló a través de la "
     "atención farmacéutica, las visitas domiciliarias y el taller de uso racional de "
     "medicamentos. El cuarto objetivo se materializó en el diseño del protocolo de "
     "atención farmacéutica domiciliaria basado en el Modelo CMO.")
para("Los resultados observados en terreno concuerdan con lo descrito en la "
     "literatura reciente. La evidencia muestra que las intervenciones farmacéuticas "
     "domiciliarias mejoran la adherencia y el conocimiento de los pacientes sobre su "
     "tratamiento, sobre todo en personas mayores y con polifarmacia (Ahn et al., "
     "2024), y que el acompañamiento sostenido en el tiempo produce mejores resultados "
     "de adherencia en pacientes crónicos (Lambert et al., 2024). En el ámbito "
     "cardiovascular, la falta de adherencia sigue siendo una de las principales "
     "barreras para el control de la hipertensión, la diabetes y la dislipidemia, y "
     "las estrategias más eficaces combinan la comunicación con el paciente, la "
     "continuidad de la atención y la participación del farmacéutico (Espinosa García "
     "et al., 2023). Asimismo, las experiencias de aplicación del Modelo CMO en "
     "distintos grupos de pacientes crónicos han mostrado mejoras en la adherencia, en "
     "los objetivos farmacoterapéuticos y en la experiencia del usuario (Cantillana "
     "Suárez et al., 2021; Sánchez Yáñez et al., 2023), aunque su adopción todavía "
     "enfrenta barreras organizacionales que conviene anticipar (Álvarez Díaz et al., "
     "2025). Estos antecedentes respaldan la pertinencia del protocolo diseñado durante "
     "el internado.")

heading("Fortalezas y debilidades del centro", level=2)
para("Entre las fortalezas del CESFAM Villa Nonguén destacan su condición de centro "
     "pionero en el Modelo de Salud Familiar, un equipo de farmacia consolidado que ya "
     "realiza atención farmacéutica en box y en el domicilio, y una cultura de trabajo "
     "orientada a la comunidad. La sectorización del territorio y la existencia de "
     "indicadores de despacho facilitan la organización de la atención y el uso "
     "eficiente de los recursos. Como aspecto por mejorar, se observó que la visita "
     "domiciliaria del farmacéutico se realizaba sin un procedimiento escrito y "
     "estandarizado, y que el registro se apoyaba en hojas personales además de la "
     "ficha clínica electrónica, lo que dificultaba la comparación entre profesionales "
     "y el seguimiento en el tiempo. La carga de trabajo y la disponibilidad de tiempo "
     "también limitaban la frecuencia de las visitas.")

heading("Sugerencias y propuestas de mejora", level=2)
numbered("Incorporar de manera formal el protocolo de atención farmacéutica "
         "domiciliaria basado en el Modelo CMO, con criterios de priorización propios "
         "del farmacéutico, para ordenar el ingreso de los pacientes al seguimiento.")
numbered("Estandarizar el registro de la atención farmacéutica mediante fichas de "
         "conciliación y de seguimiento comunes, que luego se traspasen a la ficha "
         "clínica electrónica, de modo de mejorar la trazabilidad y la continuidad.")
numbered("Reforzar el trabajo de la adherencia con herramientas simples y validadas, "
         "como el test de Morisky-Green-Levine, y con la entrevista motivacional.")
numbered("Aprovechar la continuidad telefónica y las instancias educativas grupales "
         "para mantener el vínculo con los pacientes entre una visita y otra.")

heading("Conclusiones", level=2)
numbered("El internado en el CESFAM Villa Nonguén permitió desarrollar las "
         "competencias del químico farmacéutico en el ámbito de la farmacia "
         "asistencial y la atención primaria, cumpliendo el objetivo general "
         "planteado.")
numbered("Se conocieron y se aplicaron los procesos técnicos y administrativos de la "
         "unidad de farmacia, con especial énfasis en el fraccionamiento, el "
         "reenvasado, la dispensación y la gestión del stock, resguardando la "
         "trazabilidad y el uso eficiente de los recursos.")
numbered("Se participó de manera activa en la atención farmacéutica, las visitas "
         "domiciliarias y la promoción del uso racional de los medicamentos, lo que "
         "evidenció el aporte clínico y comunitario del farmacéutico en este nivel de "
         "atención.")
numbered("Se diseñó un protocolo de atención farmacéutica domiciliaria basado en el "
         "Modelo CMO para el Programa de Salud Cardiovascular, que aporta una "
         "herramienta concreta para ordenar, estandarizar y hacer trazable la visita "
         "domiciliaria en el establecimiento.")
numbered("La experiencia confirmó la relevancia de la atención primaria como puerta "
         "de entrada al sistema de salud y el valor de un ejercicio profesional basado "
         "en la evidencia, la seguridad del paciente y el trabajo en equipo.")

# --------------------------------------------------------------- BIBLIOGRAFIA
heading("BIBLIOGRAFÍA")
refs = [
    "Ahn, H., Lee, S., Kim, S., Jang, S., & Bae, S. (2024). Effects of pharmacist-led "
    "home visit services and factors influencing medication adherence improvement. "
    "PLoS ONE, 19(11), e0313101.",
    "Álvarez-Díaz, A. M., Morillo-Verdugo, R., & Fernández-Llamazares, C. M. (2025). "
    "Estudio cualitativo sobre la adopción y potenciación del modelo "
    "capacidad-motivación-oportunidad para la atención farmacéutica en consultas "
    "externas de farmacia en España. Farmacia Hospitalaria, 49(6), 384-391.",
    "Bonal, J., Alerany, C., Bassons, T., & Gascón, P. (2003). Farmacia clínica y "
    "atención farmacéutica. Sociedad Española de Farmacia Hospitalaria.",
    "Calleja Hernández, M. Á., & Morillo Verdugo, R. (Eds.). (2016). El modelo CMO en "
    "consultas externas de farmacia hospitalaria. Sociedad Española de Farmacia "
    "Hospitalaria.",
    "Cantillana-Suárez, M. G., Robustillo-Cortés, M. A., Gutiérrez-Pizarraya, A., & "
    "Morillo-Verdugo, R. (2021). Impact and acceptance of pharmacist-led interventions "
    "during HIV care using the Capacity-Motivation-Opportunity model: The IRAFE study. "
    "European Journal of Hospital Pharmacy, 28(e1), e157-e163.",
    "Comité de Consenso de Granada. (2007). Tercer Consenso de Granada sobre problemas "
    "relacionados con medicamentos (PRM) y resultados negativos asociados a la "
    "medicación (RNM). Ars Pharmaceutica, 48(1), 5-17.",
    "COSADES. (s.f.). Historia y desarrollo institucional del CESFAM Villa Nonguén. "
    "Corporación de Salud y Desarrollo Social. https://www.cosades.cl",
    "Delgado, O., Anoz, L., Serrano, A., & Nicolás, J. (2007). Conciliación "
    "farmacoterapéutica como garantía de continuidad asistencial. Sociedad Española de "
    "Farmacia Hospitalaria.",
    "Diario Concepción. (2017, 11 de marzo). La historia del Cesfam que rompió "
    "paradigmas de la salud pública. https://www.diarioconcepcion.cl",
    "Espinosa García, J., Prados Torres, J. D., Leiva Fernández, F., & Barnestein "
    "Fonseca, P. (2023). Adherencia terapéutica de los pacientes con riesgo "
    "cardiovascular en atención primaria. Proyecto REAAP. Medicina de Familia. "
    "SEMERGEN, 49(6), 102016.",
    "Faúndez Navarrete, P. A. (2020). Evaluación del programa de atención farmacéutica "
    "del CESFAM Villa Nonguén [Memoria de título, Universidad de Concepción].",
    "Haynes, R. B., Taylor, D. W., & Sackett, D. L. (1979). Compliance in health care. "
    "Johns Hopkins University Press.",
    "Lambert, K., Bugnon, O., Del Giorno, R., Schneider, M. P., & PANDIA-IRIS Study "
    "Group. (2024). The differential impact of a 6- versus 12-month pharmacist-led "
    "interprofessional medication adherence program on medication adherence in "
    "patients with diabetic kidney disease: The randomized PANDIA-IRIS study. Frontiers "
    "in Pharmacology, 15, 1294436.",
    "Manzano García, M., & Morillo Verdugo, R. (2018). Aprendizaje y aplicación del "
    "modelo de atención farmacéutica CMO para residentes de farmacia hospitalaria. "
    "Sociedad Española de Farmacia Hospitalaria.",
    "Ministerio de Salud de Chile. (2017). Orientación técnica Programa de Salud "
    "Cardiovascular. Gobierno de Chile.",
    "Ministerio de Salud de Chile. (2018a). Guía de atención farmacéutica y "
    "seguimiento farmacoterapéutico en APS. Gobierno de Chile.",
    "Ministerio de Salud de Chile. (2018b). Orientación técnica Fondo de Farmacia. "
    "Gobierno de Chile.",
    "Ministerio de Salud de Chile. (2020). Resolución Exenta N.º 51: Aprueba Programa "
    "Fondo de Farmacia para Enfermedades Crónicas no Transmisibles. Gobierno de Chile.",
    "Morillo-Verdugo, R., Robustillo-Cortés, M. A., Navarro-Ruiz, A., "
    "Sánchez-Rubio Ferrández, J., & Fernández-Espínola, S. (2022). Clinical impact of "
    "the capacity-motivation-opportunity pharmacist-led intervention in people living "
    "with HIV in Spain, 2019-2020. Journal of Multidisciplinary Healthcare, 15, "
    "1203-1211.",
    "Morisky, D. E., Green, L. W., & Levine, D. M. (1986). Concurrent and predictive "
    "validity of a self-reported measure of medication adherence. Medical Care, 24(1), "
    "67-74.",
    "Sánchez-Yáñez, E., Manzano-García, M., & Morillo-Verdugo, R. (2023). Application "
    "of CMO (capacity, motivation, and opportunity) methodology in pharmaceutical care "
    "to optimize the pharmacotherapy in older people living with HIV. DISPIMDINAC "
    "project. Revista Española de Quimioterapia, 36(6), 584-591.",
]
for r in sorted(refs, key=lambda s: s.lower()):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.left_indent = Cm(1)
    p.paragraph_format.first_line_indent = Cm(-1)
    run = p.add_run(r); run.font.name = FONT; run.font.size = Pt(12)

# --------------------------------------------------------------- ANEXOS
heading("ANEXOS")
para("Los anexos reúnen los principales instrumentos de apoyo que acompañan al "
     "protocolo de atención farmacéutica domiciliaria diseñado durante el internado. "
     "Son formatos de trabajo pensados para usarse durante la visita y el seguimiento "
     "en el CESFAM, y pueden ajustarse a los registros propios del establecimiento.")

heading("Anexo 1. Test de adherencia de Morisky-Green-Levine", level=2)
caption("Tabla", "Preguntas del test de Morisky-Green-Levine.")
make_table(
    ["N.º", "Pregunta", "Respuesta esperada"],
    [
        ["1", "¿Olvida alguna vez tomar sus medicamentos?", "No"],
        ["2", "¿Toma los medicamentos a la hora indicada?", "Sí"],
        ["3", "Cuando se siente bien, ¿deja de tomarlos?", "No"],
        ["4", "Si alguna vez le sientan mal, ¿deja de tomarlos?", "No"],
    ],
    col_widths=[1.5, 9.5, 4.0], font_size=11)
para("Se considera adherente al paciente que responde de la forma esperada las cuatro "
     "preguntas. Cualquier otra combinación indica problemas de adherencia que conviene "
     "abordar en la entrevista.", size=11)

heading("Anexo 2. Clasificación de los RNM según el Consenso de Granada", level=2)
caption("Tabla", "Clasificación de los resultados negativos asociados a la medicación.")
make_table(
    ["Dimensión", "Pregunta que orienta", "Ejemplo"],
    [
        ["Necesidad", "¿El paciente usa los medicamentos que necesita?",
         "Toma un fármaco sin indicación o falta uno que requiere."],
        ["Efectividad", "¿El tratamiento logra el objetivo buscado?",
         "No alcanza la meta terapéutica pese a usar el medicamento."],
        ["Seguridad", "¿El tratamiento es seguro para el paciente?",
         "Aparece una reacción adversa o una interacción relevante."],
    ],
    col_widths=[3.0, 6.0, 6.0], font_size=11)

heading("Anexo 3. Otros instrumentos del protocolo", level=2)
bullet("Ficha de conciliación farmacéutica: compara la medicación indicada con la que "
       "el paciente realmente usa y deja constancia de los cambios acordados.")
bullet("Ficha de seguimiento farmacoterapéutico: registra los antecedentes del "
       "paciente, los PRM y RNM detectados, y el resultado de las intervenciones.")
bullet("Hoja de esquema de horarios de medicación: ordena, en una tabla simple, qué "
       "medicamento tomar y a qué hora, como apoyo para el paciente y su cuidador.")
bullet("Protocolo de llamada telefónica: guion breve para contactar y agendar al "
       "paciente antes de la visita domiciliaria.")

doc.save("Informe de Internado - Farmacia Asistencial y APS - CESFAM Villa Nonguen.docx")
print("Documento generado correctamente.")
