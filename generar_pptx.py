from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
import copy

# ── PALETA ────────────────────────────────────────────────────────────────────
AZUL_OSC   = RGBColor(0x1F, 0x38, 0x64)
AZUL_MED   = RGBColor(0x2E, 0x75, 0xB6)
AZUL_CLAR  = RGBColor(0xBD, 0xD7, 0xEE)
ROJO       = RGBColor(0xC0, 0x00, 0x00)
ROJO_CLAR  = RGBColor(0xFF, 0xDC, 0xDC)
VERDE      = RGBColor(0x37, 0x56, 0x23)
VERDE_MED  = RGBColor(0x70, 0xAD, 0x47)
VERDE_CLAR = RGBColor(0xE2, 0xEF, 0xDA)
AMARILLO   = RGBColor(0xFF, 0xC0, 0x00)
AMAR_CLAR  = RGBColor(0xFF, 0xF2, 0xCC)
AZUL_PALE  = RGBColor(0xDE, 0xEA, 0xF1)
GRIS       = RGBColor(0x59, 0x59, 0x59)
BLANCO     = RGBColor(0xFF, 0xFF, 0xFF)
NEGRO      = RGBColor(0x00, 0x00, 0x00)

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)

prs = Presentation()
prs.slide_width  = SLIDE_W
prs.slide_height = SLIDE_H

blank_layout = prs.slide_layouts[6]   # completamente en blanco

# ── HELPERS ──────────────────────────────────────────────────────────────────
def add_rect(slide, l, t, w, h, fill_rgb, line_rgb=None, line_width=Pt(0)):
    from pptx.util import Pt
    shape = slide.shapes.add_shape(1, l, t, w, h)   # MSO_SHAPE_TYPE.RECTANGLE = 1
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_rgb
    if line_rgb:
        shape.line.color.rgb = line_rgb
        shape.line.width = line_width
    else:
        shape.line.fill.background()
    return shape

def add_text_box(slide, text, l, t, w, h,
                 font_size=18, bold=False, color=NEGRO,
                 align=PP_ALIGN.LEFT, italic=False,
                 font_name="Calibri", word_wrap=True):
    txBox = slide.shapes.add_textbox(l, t, w, h)
    tf = txBox.text_frame
    tf.word_wrap = word_wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return txBox

def add_para(tf, text, font_size=14, bold=False, color=NEGRO,
             align=PP_ALIGN.LEFT, italic=False, space_before=Pt(4)):
    from pptx.util import Pt
    p = tf.add_paragraph()
    p.alignment = align
    p.space_before = space_before
    run = p.add_run()
    run.text = text
    run.font.name = "Calibri"
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return p

def slide_header(slide, titulo, subtitulo=None, bg=AZUL_OSC, titulo_color=BLANCO):
    # Barra superior
    add_rect(slide, 0, 0, SLIDE_W, Inches(1.15), bg)
    add_text_box(slide, titulo,
                 Inches(0.35), Inches(0.12), Inches(12.6), Inches(0.7),
                 font_size=28, bold=True, color=titulo_color, align=PP_ALIGN.LEFT)
    if subtitulo:
        add_text_box(slide, subtitulo,
                     Inches(0.35), Inches(0.78), Inches(12.6), Inches(0.35),
                     font_size=14, color=RGBColor(0xBD,0xD7,0xEE), italic=True)
    # Línea divisoria
    add_rect(slide, 0, Inches(1.15), SLIDE_W, Inches(0.04), AMARILLO)

def slide_footer(slide, numero, total=8):
    add_rect(slide, 0, Inches(7.2), SLIDE_W, Inches(0.3), AZUL_OSC)
    add_text_box(slide,
                 "Optimización de Costos de Mantenimiento 2026  |  Confidencial — Uso Interno",
                 Inches(0.3), Inches(7.2), Inches(10), Inches(0.3),
                 font_size=9, color=RGBColor(0xBD,0xD7,0xEE), italic=True)
    add_text_box(slide, f"{numero} / {total}",
                 Inches(12.3), Inches(7.2), Inches(0.9), Inches(0.3),
                 font_size=9, bold=True, color=BLANCO, align=PP_ALIGN.RIGHT)

def kpi_box(slide, l, t, w, h, label, valor, sub, bg, val_color=NEGRO):
    add_rect(slide, l, t, w, h, bg, line_rgb=RGBColor(0xBF,0xBF,0xBF), line_width=Pt(1))
    add_text_box(slide, label, l+Inches(0.1), t+Inches(0.08), w-Inches(0.2), Inches(0.3),
                 font_size=11, color=GRIS, align=PP_ALIGN.CENTER)
    add_text_box(slide, valor, l+Inches(0.05), t+Inches(0.35), w-Inches(0.1), Inches(0.55),
                 font_size=22, bold=True, color=val_color, align=PP_ALIGN.CENTER)
    add_text_box(slide, sub, l+Inches(0.1), t+Inches(0.88), w-Inches(0.2), Inches(0.28),
                 font_size=10, color=GRIS, italic=True, align=PP_ALIGN.CENTER)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 0 — PORTADA
# ══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank_layout)
add_rect(s, 0, 0, SLIDE_W, SLIDE_H, AZUL_OSC)
add_rect(s, 0, Inches(4.5), SLIDE_W, Inches(0.06), AMARILLO)

add_text_box(s, "OPTIMIZACIÓN DE COSTOS",
             Inches(0.6), Inches(1.5), Inches(12), Inches(0.9),
             font_size=40, bold=True, color=BLANCO, align=PP_ALIGN.CENTER)
add_text_box(s, "DE MANTENIMIENTO 2026",
             Inches(0.6), Inches(2.3), Inches(12), Inches(0.9),
             font_size=40, bold=True, color=AMARILLO, align=PP_ALIGN.CENTER)
add_text_box(s, "Plan de Acción · Proyección Financiera · Hoja de Ruta Mayo–Diciembre",
             Inches(0.6), Inches(3.2), Inches(12), Inches(0.5),
             font_size=18, color=AZUL_CLAR, italic=True, align=PP_ALIGN.CENTER)
add_text_box(s, "Presentación para Gerencia General  ·  Mayo 2026  ·  Confidencial",
             Inches(0.6), Inches(4.8), Inches(12), Inches(0.4),
             font_size=13, color=RGBColor(0x8F,0xAA,0xDC), align=PP_ALIGN.CENTER)

# Badges de cifras clave
bw, bh = Inches(2.3), Inches(1.3)
badges = [
    ("$299.7M", "Costo real\nEne-Abr 2026", ROJO),
    ("+33.5%", "Sobrejecución\nvs. presupuesto", AMARILLO),
    ("$113.5M", "Ahorro\nproyectado", VERDE_MED),
    ("297%", "ROI\ndel plan", VERDE_MED),
]
for i, (val, lbl, color) in enumerate(badges):
    bx = Inches(0.6) + i*(bw + Inches(0.22))
    add_rect(s, bx, Inches(5.4), bw, bh, RGBColor(0x0D,0x1F,0x3C))
    add_text_box(s, val,  bx, Inches(5.45), bw, Inches(0.55),
                 font_size=26, bold=True, color=color, align=PP_ALIGN.CENTER)
    add_text_box(s, lbl, bx, Inches(5.98), bw, Inches(0.45),
                 font_size=11, color=RGBColor(0xBD,0xD7,0xEE), align=PP_ALIGN.CENTER)

slide_footer(s, 0, 8)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 1 — EL PROBLEMA
# ══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank_layout)
add_rect(s, 0, Inches(1.19), SLIDE_W, SLIDE_H - Inches(1.19) - Inches(0.3), RGBColor(0xF7,0xF9,0xFC))
slide_header(s, "El Problema: Mantenimiento Fuera de Control",
             "Situación al cierre de Abril 2026", bg=ROJO)
slide_footer(s, 1)

# KPI boxes fila superior
kpis_1 = [
    ("Costo real Ene-Abr", "$299.7M", "Presupuesto: $185M", ROJO_CLAR, ROJO),
    ("Sobrejecución", "+$114.7M", "+33.5% vs. ppto", ROJO_CLAR, ROJO),
    ("KPI Costo / Ventas", "3.89%", "Meta: ≤ 2.5%", ROJO_CLAR, ROJO),
    ("Margen bruto", "-2.8 pp", "Erosión vs. 2025", AMAR_CLAR, RGBColor(0x7F,0x60,0x00)),
]
kw = Inches(2.9)
for i, (lbl, val, sub, bg, vc) in enumerate(kpis_1):
    kpi_box(s, Inches(0.35) + i*(kw+Inches(0.22)), Inches(1.35),
            kw, Inches(1.25), lbl, val, sub, bg, vc)

# Tabla mes a mes
headers = ["Mes", "Costo Real", "Presupuesto", "Desviación", "% Desv."]
rows_data = [
    ("Enero 2026",    "$53.9M", "$40.0M", "+$13.9M", "+35%"),
    ("Febrero 2026",  "$56.9M", "$45.0M", "+$11.9M", "+26%"),
    ("Marzo 2026",    "$85.1M", "$50.0M", "+$35.1M", "+70%"),
    ("Abril 2026",   "$103.8M", "$50.0M", "+$53.8M", "+108%"),
    ("TOTAL",        "$299.7M", "$185.0M","+$114.7M", "+62%"),
]
col_ws = [Inches(2.0), Inches(2.0), Inches(2.0), Inches(2.0), Inches(1.6)]
col_xs = [Inches(0.35)]
for w in col_ws[:-1]:
    col_xs.append(col_xs[-1] + w + Inches(0.07))

row_h = Inches(0.44)
t_top = Inches(2.75)

# header row
for j, (hdr, cx, cw) in enumerate(zip(headers, col_xs, col_ws)):
    add_rect(s, cx, t_top, cw, row_h, AZUL_OSC)
    add_text_box(s, hdr, cx+Inches(0.05), t_top+Inches(0.07), cw-Inches(0.1), row_h-Inches(0.1),
                 font_size=13, bold=True, color=BLANCO, align=PP_ALIGN.CENTER)

for i, (row) in enumerate(rows_data):
    ry = t_top + (i+1)*row_h + Inches(0.05)*i
    is_total = row[0] == "TOTAL"
    bg_row = ROJO_CLAR if is_total else (RGBColor(0xF2,0xF2,0xF2) if i%2==0 else BLANCO)
    for j, (cell, cx, cw) in enumerate(zip(row, col_xs, col_ws)):
        add_rect(s, cx, ry, cw, row_h, bg_row,
                 line_rgb=RGBColor(0xBF,0xBF,0xBF), line_width=Pt(0.5))
        cell_color = ROJO if (j >= 3 and "+" in cell) else (NEGRO if not is_total else ROJO)
        add_text_box(s, cell, cx+Inches(0.05), ry+Inches(0.07),
                     cw-Inches(0.1), row_h-Inches(0.1),
                     font_size=13, bold=is_total,
                     color=cell_color, align=PP_ALIGN.CENTER)

# Destacado lateral
add_rect(s, Inches(10.2), Inches(2.75), Inches(2.85), Inches(4.2), ROJO_CLAR,
         line_rgb=ROJO, line_width=Pt(2))
add_text_box(s, "⚠ ABRIL 2026",
             Inches(10.3), Inches(2.85), Inches(2.65), Inches(0.45),
             font_size=14, bold=True, color=ROJO, align=PP_ALIGN.CENTER)
add_text_box(s, "+108%\nsobre presupuesto",
             Inches(10.3), Inches(3.3), Inches(2.65), Inches(0.8),
             font_size=22, bold=True, color=ROJO, align=PP_ALIGN.CENTER)
add_text_box(s, "El peor mes registrado.\n$68.7M sin descripción\ntécnica documentada.",
             Inches(10.3), Inches(4.15), Inches(2.65), Inches(0.85),
             font_size=12, color=GRIS, align=PP_ALIGN.CENTER)
add_text_box(s, "Sin acción inmediata:\ncosto 2026 proyectado\n→ $993M (+72%)",
             Inches(10.3), Inches(5.1), Inches(2.65), Inches(0.75),
             font_size=12, bold=True, color=ROJO, align=PP_ALIGN.CENTER)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 2 — DIAGNÓSTICO
# ══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank_layout)
add_rect(s, 0, Inches(1.19), SLIDE_W, SLIDE_H - Inches(1.19) - Inches(0.3), RGBColor(0xF7,0xF9,0xFC))
slide_header(s, "Diagnóstico: ¿Dónde Está el Dinero?",
             "6 causas raíz identificadas en el análisis de datos", bg=AZUL_OSC)
slide_footer(s, 2)

causas = [
    ("1", "CIF Mantenimiento\n+80.2% YoY",
     "$179.7M (2025)\n→ $299.7M (2026)", ROJO, ROJO_CLAR),
    ("2", "Personal Temporal\n+65.3% sin aprobación",
     "$54M en 4 meses\n3 cooperativas activas", ROJO, ROJO_CLAR),
    ("3", "Pico Inyección Marzo\nCiclo recurrente",
     "$59M en un solo mes\nNo presupuestado", AMARILLO, AMAR_CLAR),
    ("4", "$68.7M sin trazabilidad\n(55% de Abril)",
     "Sin OT documentada\nNo se puede controlar", ROJO, ROJO_CLAR),
    ("5", "Brecha contable\n$146.9M",
     "2 fuentes distintas\nSin conciliación", AMARILLO, AMAR_CLAR),
    ("6", "70% Correctivo\n30% Preventivo",
     "Correctivo cuesta 3-5x\nmás que preventivo", ROJO, ROJO_CLAR),
]

box_w = Inches(3.85)
box_h = Inches(2.35)
for i, (num, titulo, detalle, num_color, bg) in enumerate(causas):
    col = i % 3
    row = i // 3
    bx = Inches(0.35) + col * (box_w + Inches(0.32))
    by = Inches(1.35) + row * (box_h + Inches(0.22))
    add_rect(s, bx, by, box_w, box_h, bg, line_rgb=num_color, line_width=Pt(2))
    # Número
    add_rect(s, bx, by, Inches(0.48), Inches(0.48), num_color)
    add_text_box(s, num, bx, by, Inches(0.48), Inches(0.48),
                 font_size=16, bold=True, color=BLANCO, align=PP_ALIGN.CENTER)
    add_text_box(s, titulo, bx+Inches(0.55), by+Inches(0.06), box_w-Inches(0.65), Inches(0.75),
                 font_size=13, bold=True, color=num_color)
    add_text_box(s, detalle, bx+Inches(0.15), by+Inches(0.9), box_w-Inches(0.3), Inches(1.1),
                 font_size=12, color=GRIS)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 3 — PLAN EN 3 FASES
# ══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank_layout)
add_rect(s, 0, Inches(1.19), SLIDE_W, SLIDE_H - Inches(1.19) - Inches(0.3), RGBColor(0xF7,0xF9,0xFC))
slide_header(s, "La Solución: Plan de Acción en 3 Fases",
             "Mayo → Diciembre 2026  ·  Meta: reducir KPI de 3.89% a ≤ 2.5%", bg=AZUL_OSC)
slide_footer(s, 3)

fases = [
    ("FASE 1", "ESTABILIZACIÓN", "Mayo – Junio 2026",
     "Parar la hemorragia",
     ["• Control de OT con aprobación escalonada",
      "  (<$2M / <$10M / Gerencia General)",
      "• Congelar personal temporal nuevo",
      "• Auditoría y conciliación $146.9M",
      "• Documentar retroactivo Abril"],
     "Ahorro meta: $51.6M",
     ROJO, ROJO_CLAR),
    ("FASE 2", "CONTROL", "Julio – Septiembre 2026",
     "Reducir correctivo al 45%",
     ["• Contratos anuales 5 proveedores (-12%)",
      "  Tornifer, Maquitec, Optimoldes...",
      "• Stock mínimo repuestos críticos ($45M)",
      "• Plan Maestro Mantenimiento Preventivo",
      "• Gestión anticíclica inyección marzo"],
     "Ahorro meta: $88.7M",
     AZUL_MED, AZUL_CLAR),
    ("FASE 3", "TRANSFORMACIÓN", "Oct – Diciembre 2026",
     "Cultura preventiva ≥ 60%",
     ["• Implementar CMMS / módulo PM en ERP",
      "• Redimensionar nómina fija vs. temporal",
      "• Presupuesto Base Cero 2027",
      "• % Preventivo ≥ 60% consolidado"],
     "Ahorro meta: $25M",
     VERDE, VERDE_CLAR),
]

fw = Inches(3.85)
for i, (f_num, f_nom, f_per, f_obj, f_acc, f_aho, f_col, f_bg) in enumerate(fases):
    fx = Inches(0.35) + i * (fw + Inches(0.30))
    fy = Inches(1.35)
    fh = Inches(5.85)
    add_rect(s, fx, fy, fw, fh, f_bg, line_rgb=f_col, line_width=Pt(2))
    # Header de fase
    add_rect(s, fx, fy, fw, Inches(0.95), f_col)
    add_text_box(s, f_num, fx+Inches(0.1), fy+Inches(0.04), fw-Inches(0.2), Inches(0.35),
                 font_size=12, bold=True, color=BLANCO)
    add_text_box(s, f_nom, fx+Inches(0.1), fy+Inches(0.38), fw-Inches(0.2), Inches(0.35),
                 font_size=16, bold=True, color=BLANCO)
    # Período y objetivo
    add_text_box(s, f_per, fx+Inches(0.12), fy+Inches(1.02), fw-Inches(0.24), Inches(0.28),
                 font_size=11, bold=True, color=f_col)
    add_text_box(s, f_obj, fx+Inches(0.12), fy+Inches(1.28), fw-Inches(0.24), Inches(0.28),
                 font_size=11, color=GRIS, italic=True)
    # Acciones
    for j, acc in enumerate(f_acc):
        add_text_box(s, acc, fx+Inches(0.12), fy+Inches(1.65) + j*Inches(0.58),
                     fw-Inches(0.24), Inches(0.55), font_size=11, color=NEGRO)
    # Badge de ahorro
    add_rect(s, fx+Inches(0.12), fy+Inches(5.2), fw-Inches(0.24), Inches(0.48), f_col)
    add_text_box(s, f_aho, fx+Inches(0.12), fy+Inches(5.22), fw-Inches(0.24), Inches(0.44),
                 font_size=14, bold=True, color=BLANCO, align=PP_ALIGN.CENTER)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 4 — PROYECCIÓN FINANCIERA
# ══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank_layout)
add_rect(s, 0, Inches(1.19), SLIDE_W, SLIDE_H - Inches(1.19) - Inches(0.3), RGBColor(0xF7,0xF9,0xFC))
slide_header(s, "Proyección Financiera: Sin Plan vs. Con Plan",
             "Trayectoria mensual Mayo–Diciembre 2026  ·  COP Millones", bg=AZUL_OSC)
slide_footer(s, 4)

# Tabla de proyección
meses_c = ["Mayo","Junio","Julio","Ago.","Sep.","Oct.","Nov.","Dic."]
sin_p = [81.1, 75.7, 92.7, 89.6, 84.6, 78.5, 94.6, 97.0]
con_p = [74.4, 69.0, 86.0, 79.2, 67.0, 59.1, 71.7, 74.0]
ppto_c = [47.0, 44.0, 52.0, 51.0, 49.0, 46.0, 53.0, 54.0]

tl = Inches(0.35)
tt = Inches(1.42)
tw = Inches(9.3)
col_w0 = Inches(2.0)
col_wn = (tw - col_w0) / 8
row_hdr_h = Inches(0.38)
row_h2 = Inches(0.46)

# Columnas: etiqueta + 8 meses
hdrs_t = ["Escenario"] + meses_c
cw_list = [col_w0] + [col_wn]*8
cx_list = [tl]
for cw in cw_list[:-1]:
    cx_list.append(cx_list[-1] + cw)

for j, (hdr, cx, cw) in enumerate(zip(hdrs_t, cx_list, cw_list)):
    add_rect(s, cx, tt, cw, row_hdr_h, AZUL_OSC)
    add_text_box(s, hdr, cx+Inches(0.04), tt+Inches(0.05),
                 cw-Inches(0.08), row_hdr_h-Inches(0.08),
                 font_size=11, bold=True, color=BLANCO, align=PP_ALIGN.CENTER)

table_rows = [
    ("Sin Plan",   sin_p,  ROJO_CLAR,  ROJO),
    ("Con Plan",   con_p,  VERDE_CLAR, VERDE),
    ("Presupuesto",ppto_c, AMAR_CLAR,  RGBColor(0x7F,0x60,0x00)),
]
for ri, (lbl, vals, bg, col) in enumerate(table_rows):
    ry = tt + row_hdr_h + ri*row_h2 + Inches(0.04)*ri
    for j, (val, cx, cw) in enumerate(zip([lbl]+vals, cx_list, cw_list)):
        add_rect(s, cx, ry, cw, row_h2, bg,
                 line_rgb=RGBColor(0xBF,0xBF,0xBF), line_width=Pt(0.5))
        txt = lbl if j == 0 else f"${val:.1f}M"
        add_text_box(s, txt, cx+Inches(0.04), ry+Inches(0.08),
                     cw-Inches(0.08), row_h2-Inches(0.12),
                     font_size=11, bold=(j==0), color=col if j>0 else col,
                     align=PP_ALIGN.CENTER)

# Fila de ahorro
ry_aho = tt + row_hdr_h + 3*row_h2 + Inches(0.18)
add_rect(s, tl, ry_aho, tw, row_h2, VERDE_CLAR,
         line_rgb=VERDE, line_width=Pt(1.5))
ahorros = [f"${(s-c):.1f}M" for s,c in zip(sin_p, con_p)]
add_text_box(s, "Ahorro mensual", tl+Inches(0.04), ry_aho+Inches(0.08),
             col_w0-Inches(0.08), row_h2-Inches(0.12),
             font_size=11, bold=True, color=VERDE, align=PP_ALIGN.CENTER)
for j, (aho, cx, cw) in enumerate(zip(ahorros, cx_list[1:], cw_list[1:])):
    add_text_box(s, aho, cx+Inches(0.04), ry_aho+Inches(0.08),
                 cw-Inches(0.08), row_h2-Inches(0.12),
                 font_size=11, bold=True, color=VERDE, align=PP_ALIGN.CENTER)

# Cajas de totales
totals_data = [
    ("Sin Plan\nMay-Dic", "$693.8M", ROJO_CLAR, ROJO),
    ("Con Plan\nMay-Dic", "$580.3M", VERDE_CLAR, VERDE),
    ("Ahorro\nTotal",     "$113.5M", VERDE_CLAR, VERDE),
    ("KPI Con Plan\nDic 2026","2.9%",AZUL_CLAR, AZUL_MED),
]
bw2 = Inches(2.65)
by2 = Inches(5.1)
for i, (lbl, val, bg, vc) in enumerate(totals_data):
    bx2 = Inches(9.82) if i == 0 else Inches(9.82)  # right panel
    bx2 = Inches(0.35) + i*(bw2 + Inches(0.24))
    add_rect(s, bx2, by2, bw2, Inches(1.6), bg, line_rgb=vc, line_width=Pt(1.5))
    add_text_box(s, lbl, bx2+Inches(0.1), by2+Inches(0.08), bw2-Inches(0.2), Inches(0.5),
                 font_size=11, color=GRIS, align=PP_ALIGN.CENTER)
    add_text_box(s, val, bx2+Inches(0.05), by2+Inches(0.55), bw2-Inches(0.1), Inches(0.75),
                 font_size=26, bold=True, color=vc, align=PP_ALIGN.CENTER)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 5 — LAS 6 PALANCAS
# ══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank_layout)
add_rect(s, 0, Inches(1.19), SLIDE_W, SLIDE_H - Inches(1.19) - Inches(0.3), RGBColor(0xF7,0xF9,0xFC))
slide_header(s, "Las 6 Palancas de Ahorro",
             "Acciones con mayor impacto financiero  ·  Total May-Dic 2026: $113.5M COP", bg=VERDE)
slide_footer(s, 5)

palancas_d = [
    ("1", "Control de OT",
     "Aprobación escalonada\npor monto. Bloquea\ngastos no autorizados.",
     "$4.0M / mes\nDesde Mayo",
     "Alta", ROJO),
    ("2", "Reducción Temporal",
     "De $15.2M/mes a\n$7.5M/mes en dic.\nAuditoría 3 cooperativas.",
     "$2.7M → $4.2M/mes\nDesde Mayo",
     "Alta", ROJO),
    ("3", "Contratos Proveedores",
     "5 contratos anuales.\nTornifer, Maquitec,\nOptimoldes, IDEAS T.",
     "$2.2M → $4.0M/mes\nDesde Agosto",
     "Media-Alta", AZUL_MED),
    ("4", "Stock Repuestos",
     "Elimina recargo 30%\nen compras emergencia.\n$45M inversión.",
     "$2.25M / mes\nDesde Septiembre",
     "Alta", AZUL_MED),
    ("5", "Plan Preventivo",
     "PMP para 15+ equipos.\nMeta: 60% preventivo\nen diciembre.",
     "$5.0M → $8.5M/mes\nDesde Septiembre",
     "Media", VERDE),
    ("6", "Pico Inyección Mar.",
     "Contador de ciclos.\nDistribuir $59M en\n6 cuotas preventivas.",
     "$35-45M evitados\nen Marzo 2027",
     "Media", VERDE),
]

pw = Inches(3.85)
ph = Inches(2.7)
for i, (num, nom, desc, aho, cert, col) in enumerate(palancas_d):
    col_i = i % 3
    row_i = i // 3
    px = Inches(0.35) + col_i*(pw + Inches(0.30))
    py = Inches(1.38) + row_i*(ph + Inches(0.22))
    add_rect(s, px, py, pw, ph, BLANCO, line_rgb=col, line_width=Pt(2))
    # Badge número
    add_rect(s, px, py, Inches(0.42), Inches(0.42), col)
    add_text_box(s, num, px, py, Inches(0.42), Inches(0.42),
                 font_size=14, bold=True, color=BLANCO, align=PP_ALIGN.CENTER)
    add_text_box(s, nom, px+Inches(0.5), py+Inches(0.06), pw-Inches(0.6), Inches(0.38),
                 font_size=14, bold=True, color=col)
    add_text_box(s, desc, px+Inches(0.12), py+Inches(0.5), pw-Inches(0.24), Inches(0.95),
                 font_size=11, color=GRIS)
    # Badge ahorro
    add_rect(s, px+Inches(0.12), py+Inches(1.55), pw-Inches(0.24), Inches(0.52), col)
    add_text_box(s, aho, px+Inches(0.12), py+Inches(1.57), pw-Inches(0.24), Inches(0.48),
                 font_size=12, bold=True, color=BLANCO, align=PP_ALIGN.CENTER)
    # Certeza
    cert_color = VERDE if cert=="Alta" else (RGBColor(0x7F,0x60,0x00) if "Media" in cert else GRIS)
    add_text_box(s, f"Certeza: {cert}", px+Inches(0.12), py+Inches(2.22), pw-Inches(0.24), Inches(0.3),
                 font_size=9, color=cert_color, italic=True)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 6 — ROI
# ══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank_layout)
add_rect(s, 0, Inches(1.19), SLIDE_W, SLIDE_H - Inches(1.19) - Inches(0.3), RGBColor(0xF7,0xF9,0xFC))
slide_header(s, "ROI del Plan: Inversión vs. Retorno",
             "Análisis financiero completo  ·  La inversión se recupera en Octubre 2026", bg=VERDE)
slide_footer(s, 6)

# Columna izq: Inversión
add_rect(s, Inches(0.35), Inches(1.38), Inches(5.8), Inches(5.55), ROJO_CLAR,
         line_rgb=ROJO, line_width=Pt(2))
add_rect(s, Inches(0.35), Inches(1.38), Inches(5.8), Inches(0.52), ROJO)
add_text_box(s, "INVERSIÓN REQUERIDA", Inches(0.45), Inches(1.4), Inches(5.6), Inches(0.45),
             font_size=15, bold=True, color=BLANCO, align=PP_ALIGN.CENTER)

inv_items = [
    ("Stock repuestos críticos", "$45,000,000", "(activo en balance — no es gasto)"),
    ("CMMS / módulo PM en ERP",  "$5,000,000",  "(SaaS o módulo existente)"),
    ("Capacitación RCM",         "$3,000,000",  "(técnicos de mantenimiento)"),
    ("TOTAL",                    "$53,000,000", ""),
]
for j, (lbl, val, nota) in enumerate(inv_items):
    ry = Inches(2.05) + j*Inches(0.85)
    is_tot = lbl == "TOTAL"
    if is_tot:
        add_rect(s, Inches(0.45), ry, Inches(5.6), Inches(0.7), ROJO)
    add_text_box(s, lbl, Inches(0.55), ry+Inches(0.06), Inches(2.8), Inches(0.35),
                 font_size=13, bold=is_tot, color=BLANCO if is_tot else NEGRO)
    add_text_box(s, val, Inches(3.4), ry+Inches(0.06), Inches(1.8), Inches(0.35),
                 font_size=14, bold=True, color=BLANCO if is_tot else ROJO,
                 align=PP_ALIGN.RIGHT)
    if nota:
        add_text_box(s, nota, Inches(0.55), ry+Inches(0.42), Inches(5.1), Inches(0.3),
                     font_size=9, color=GRIS, italic=True)

# Columna der: Retorno
add_rect(s, Inches(6.5), Inches(1.38), Inches(6.48), Inches(5.55), VERDE_CLAR,
         line_rgb=VERDE, line_width=Pt(2))
add_rect(s, Inches(6.5), Inches(1.38), Inches(6.48), Inches(0.52), VERDE)
add_text_box(s, "RETORNO DEL PLAN", Inches(6.6), Inches(1.4), Inches(6.28), Inches(0.45),
             font_size=15, bold=True, color=BLANCO, align=PP_ALIGN.CENTER)

ret_items = [
    ("Ahorros May-Dic 2026",    "$113,500,000"),
    ("Inversión requerida",     "-$53,000,000"),
    ("Beneficio neto 2026",     "$60,500,000"),
    ("ROI del plan (2026)",     "297%"),
    ("Payback",                 "Octubre 2026"),
    ("ROI acumulado 2026-2027", "590-680%"),
]
for j, (lbl, val) in enumerate(ret_items):
    ry = Inches(2.05) + j*Inches(0.78)
    is_key = lbl in ["Beneficio neto 2026","ROI del plan (2026)","ROI acumulado 2026-2027"]
    bg = VERDE if is_key else VERDE_CLAR
    fc = BLANCO if is_key else VERDE
    if is_key:
        add_rect(s, Inches(6.55), ry-Inches(0.04), Inches(6.35), Inches(0.68), VERDE)
    add_text_box(s, lbl, Inches(6.65), ry+Inches(0.04), Inches(3.5), Inches(0.35),
                 font_size=13, bold=is_key, color=BLANCO if is_key else NEGRO)
    add_text_box(s, val, Inches(10.2), ry+Inches(0.04), Inches(2.5), Inches(0.35),
                 font_size=14, bold=True, color=BLANCO if is_key else VERDE,
                 align=PP_ALIGN.RIGHT)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 7 — KPIs
# ══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank_layout)
add_rect(s, 0, Inches(1.19), SLIDE_W, SLIDE_H - Inches(1.19) - Inches(0.3), RGBColor(0xF7,0xF9,0xFC))
slide_header(s, "Tablero de KPIs — Metas Mensuales",
             "Indicadores de seguimiento del plan  ·  Reporte en reunión mensual de resultados", bg=AZUL_MED)
slide_footer(s, 7)

kpi_track = [
    ("KPI Costo / Ventas",   "3.89%",  "≤ 2.5%",  ROJO,     "Reducción gradual cada mes"),
    ("% Mant. Preventivo",   "~30%",   "≥ 60%",   AMARILLO, "Meta de 60% en Diciembre"),
    ("Personal Temporal/mes","$15.2M", "≤ $7.5M", ROJO,     "Reducción -50% en 8 meses"),
    ("Proveedores contrato", "0",      "5",        ROJO,     "5 contratos firmados en Sept."),
    ("OT con aprobación",    "~30%",   "100%",     ROJO,     "100% desde Junio 2026"),
    ("Paros x falta repuesto","N/D",   "0",        AMARILLO, "Stock mínimo activo en Sept."),
]

kw3 = Inches(3.82)
kh3 = Inches(1.3)
for i, (nombre, actual, meta, alerta, nota) in enumerate(kpi_track):
    col_i = i % 2
    row_i = i // 2
    kx = Inches(0.35) + col_i*(kw3*2 + Inches(0.46)) if col_i == 0 else Inches(0.35) + kw3 + Inches(0.42)
    # full-width row? No, 2 per row but each card takes half
    kx = Inches(0.35) + col_i*(kw3 + Inches(0.42))
    # actually let's do 2 columns of 2 cards each, and a third col for the note
    # Simpler: 2 per row, wider cards
    kw3_full = Inches(6.1)
    kx = Inches(0.35) + col_i*(kw3_full + Inches(0.42))
    ky = Inches(1.42) + row_i*(kh3 + Inches(0.22))

    add_rect(s, kx, ky, kw3_full, kh3, BLANCO,
             line_rgb=alerta, line_width=Pt(2))
    # Barra lateral de color
    add_rect(s, kx, ky, Inches(0.18), kh3, alerta)
    add_text_box(s, nombre, kx+Inches(0.28), ky+Inches(0.06), Inches(2.5), Inches(0.35),
                 font_size=12, bold=True, color=NEGRO)
    add_text_box(s, "Actual:", kx+Inches(0.28), ky+Inches(0.42), Inches(0.7), Inches(0.3),
                 font_size=10, color=GRIS)
    add_text_box(s, actual, kx+Inches(0.95), ky+Inches(0.38), Inches(1.0), Inches(0.38),
                 font_size=14, bold=True, color=ROJO)
    add_text_box(s, "→", kx+Inches(1.95), ky+Inches(0.38), Inches(0.35), Inches(0.38),
                 font_size=14, color=GRIS, align=PP_ALIGN.CENTER)
    add_text_box(s, "Meta:", kx+Inches(2.35), ky+Inches(0.42), Inches(0.7), Inches(0.3),
                 font_size=10, color=GRIS)
    add_text_box(s, meta, kx+Inches(3.0), ky+Inches(0.38), Inches(1.0), Inches(0.38),
                 font_size=14, bold=True, color=VERDE)
    add_text_box(s, nota, kx+Inches(0.28), ky+Inches(0.82), kw3_full-Inches(0.5), Inches(0.3),
                 font_size=10, color=GRIS, italic=True)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 8 — PRÓXIMOS PASOS
# ══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(blank_layout)
add_rect(s, 0, Inches(1.19), SLIDE_W, SLIDE_H - Inches(1.19) - Inches(0.3), RGBColor(0xF7,0xF9,0xFC))
slide_header(s, "Cronograma y Próximos Pasos",
             "Acciones inmediatas — Semana 1 de Mayo 2026", bg=AZUL_OSC)
slide_footer(s, 8)

# Timeline visual
timeline = [
    ("SEMANA 1\nMayo 2026", [
        "Aprobar plan en junta directiva",
        "Implementar OT con aprobación escalonada",
        "Congelar personal temporal nuevo",
        "Documentar retroactivo abril ($68.7M)",
    ], ROJO),
    ("JUNIO –\nJULIO 2026", [
        "Iniciar negociación 5 proveedores",
        "Auditoría contable $146.9M",
        "Reducir temporal a $9M/mes",
        "Primer contrato firmado",
    ], AMARILLO),
    ("AGOSTO –\nSEPT. 2026", [
        "Stock repuestos críticos operativo",
        "Plan Maestro PM activo",
        "5 contratos firmados",
        "% Preventivo ≥ 52%",
    ], AZUL_MED),
    ("OCT. –\nDIC. 2026", [
        "CMMS implementado",
        "Nómina 2027 redimensionada",
        "Presupuesto Base Cero 2027",
        "KPI ≤ 2.9% | Preventivo ≥ 60%",
    ], VERDE),
]

tw2 = Inches(2.92)
for i, (periodo, acciones, col) in enumerate(timeline):
    tx = Inches(0.35) + i*(tw2 + Inches(0.25))
    ty = Inches(1.42)
    # Header
    add_rect(s, tx, ty, tw2, Inches(0.72), col)
    add_text_box(s, periodo, tx+Inches(0.1), ty+Inches(0.06), tw2-Inches(0.2), Inches(0.62),
                 font_size=13, bold=True, color=BLANCO, align=PP_ALIGN.CENTER)
    # Flecha conectora (excepto último)
    if i < 3:
        add_rect(s, tx+tw2, ty+Inches(0.26), Inches(0.25), Inches(0.2), col)

    # Acciones
    for j, acc in enumerate(acciones):
        ay = ty + Inches(0.82) + j*Inches(0.72)
        add_rect(s, tx, ay, tw2, Inches(0.64), BLANCO if j%2==0 else RGBColor(0xF7,0xF7,0xF7),
                 line_rgb=RGBColor(0xBF,0xBF,0xBF), line_width=Pt(0.5))
        add_rect(s, tx, ay, Inches(0.14), Inches(0.64), col)
        add_text_box(s, acc, tx+Inches(0.22), ay+Inches(0.08), tw2-Inches(0.3), Inches(0.5),
                     font_size=11, color=NEGRO)

# Bloque de condiciones de éxito
add_rect(s, Inches(0.35), Inches(5.92), Inches(12.63), Inches(1.0), AZUL_PALE,
         line_rgb=AZUL_MED, line_width=Pt(1.5))
add_text_box(s, "CONDICIÓN DE ÉXITO:",
             Inches(0.5), Inches(5.98), Inches(2.4), Inches(0.4),
             font_size=12, bold=True, color=AZUL_OSC)
add_text_box(s,
    "Gerencia General debe garantizar: (1) Autoridad real del Jefe de Mantenimiento para rechazar trabajos no priorizados  "
    "(2) Reporte mensual de KPIs en reunión de resultados  (3) No comprimir el presupuesto de mantenimiento como variable de ajuste",
    Inches(2.95), Inches(5.98), Inches(9.85), Inches(0.78),
    font_size=11, color=AZUL_OSC)

# ── GUARDAR ───────────────────────────────────────────────────────────────────
output = "/home/user/Todos-financieros/Presentacion_Mantenimiento_2026.pptx"
prs.save(output)
print(f"PowerPoint guardado: {output}")
print(f"Slides: {len(prs.slides)}")
