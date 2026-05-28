import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
from openpyxl.chart.label import DataLabel
import openpyxl.chart

# ── PALETA DE COLORES ─────────────────────────────────────────────────────────
AZUL_OSC   = "1F3864"   # Encabezados principales
AZUL_MED   = "2E75B6"   # Encabezados secundarios
AZUL_CLAR  = "BDD7EE"   # Fondos de sección
AZUL_PALE  = "DEEAF1"   # Filas alternas
ROJO       = "C00000"   # Sobrejecución / Malo
ROJO_PALE  = "FFDCDC"
VERDE      = "375623"   # Ahorro / Bueno
VERDE_MED  = "70AD47"
VERDE_PALE = "E2EFDA"
AMARILLO   = "FFC000"   # Advertencia
AMAR_PALE  = "FFF2CC"
GRIS_CLAR  = "F2F2F2"
BLANCO     = "FFFFFF"
NEGRO      = "000000"

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def font(bold=False, color=NEGRO, size=11, italic=False):
    return Font(name="Calibri", bold=bold, color=color, size=size, italic=italic)

def border_thin():
    s = Side(style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)

def border_medium():
    s = Side(style="medium", color="595959")
    return Border(left=s, right=s, top=s, bottom=s)

def center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def right():
    return Alignment(horizontal="right", vertical="center")

def left_align():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)

def apply_header_row(ws, row, cols, text_list, bg=AZUL_OSC, fg=BLANCO, size=11):
    for i, (col, text) in enumerate(zip(cols, text_list)):
        c = ws.cell(row=row, column=col, value=text)
        c.fill = fill(bg)
        c.font = font(bold=True, color=fg, size=size)
        c.alignment = center()
        c.border = border_thin()

def fmt_cop(ws, cell_range):
    for row in ws[cell_range]:
        for c in row:
            c.number_format = '#,##0'

def fmt_pct(ws, cell_range):
    for row in ws[cell_range]:
        for c in row:
            c.number_format = '0.0%'

# ─────────────────────────────────────────────────────────────────────────────
wb = openpyxl.Workbook()
wb.remove(wb.active)  # quitar hoja default

# ══════════════════════════════════════════════════════════════════════════════
# HOJA 1: DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
dash = wb.create_sheet("DASHBOARD")
dash.sheet_view.showGridLines = False
dash.column_dimensions["A"].width = 2
dash.column_dimensions["B"].width = 32
dash.column_dimensions["C"].width = 22
dash.column_dimensions["D"].width = 22
dash.column_dimensions["E"].width = 22
dash.column_dimensions["F"].width = 22
dash.column_dimensions["G"].width = 2

# Título principal
dash.merge_cells("B1:F1")
c = dash["B1"]
c.value = "MODELO FINANCIERO — MANTENIMIENTO 2026"
c.fill = fill(AZUL_OSC)
c.font = font(bold=True, color=BLANCO, size=16)
c.alignment = center()
c.border = border_medium()
dash.row_dimensions[1].height = 36

dash.merge_cells("B2:F2")
c = dash["B2"]
c.value = "Proyección Mayo–Diciembre 2026 | Trayectoria Sin Plan vs. Con Plan | Corte: 30 Abril 2026"
c.fill = fill(AZUL_MED)
c.font = font(color=BLANCO, size=11, italic=True)
c.alignment = center()
dash.row_dimensions[2].height = 20

# ── BLOQUE: ESTADO DE PARTIDA ──────────────────────────────────────────────
dash.merge_cells("B4:F4")
c = dash["B4"]
c.value = "ESTADO DE PARTIDA — Enero a Abril 2026 (Real)"
c.fill = fill(AZUL_MED)
c.font = font(bold=True, color=BLANCO, size=12)
c.alignment = center()
dash.row_dimensions[4].height = 22

kpi_data = [
    ("Costo real Ene-Abr 2026",    299723600, "COP",   ""),
    ("Presupuesto Ene-Abr 2026",   185000000, "COP",   ""),
    ("Sobrejecución Ene-Abr",      114723600, "COP",   "MALO"),
    ("% Desviación vs. ppto",          0.620, "%",     "MALO"),
    ("KPI Costo/Ventas (acum.)",      0.0389, "%",     "MALO"),
    ("% Mant. Preventivo actual",     0.30,   "%",     "ATEN"),
]

for i, (label, val, tipo, status) in enumerate(kpi_data):
    row = 5 + i
    dash.row_dimensions[row].height = 20
    c = dash.cell(row=row, column=2, value=label)
    c.fill = fill(AZUL_PALE if i % 2 == 0 else BLANCO)
    c.font = font(size=11)
    c.alignment = left_align()
    c.border = border_thin()

    cv = dash.cell(row=row, column=3, value=val)
    cv.fill = fill(AZUL_PALE if i % 2 == 0 else BLANCO)
    cv.font = font(bold=True, size=11)
    cv.alignment = right()
    cv.border = border_thin()
    if tipo == "COP":
        cv.number_format = '"$"#,##0'
    elif tipo == "%":
        cv.number_format = '0.0%'

    for col in [4, 5, 6]:
        cx = dash.cell(row=row, column=col)
        cx.fill = fill(AZUL_PALE if i % 2 == 0 else BLANCO)
        cx.border = border_thin()
        if col == 4 and status == "MALO":
            cx.value = "⚠ Por encima de meta"
            cx.font = font(bold=True, color=ROJO, size=10)
            cx.alignment = left_align()
        elif col == 4 and status == "ATEN":
            cx.value = "→ Meta: ≥ 60% en Dic 2026"
            cx.font = font(color="7F6000", size=10, italic=True)
            cx.alignment = left_align()

# ── BLOQUE: RESUMEN DE PROYECCIÓN ────────────────────────────────────────────
start_r = 12
dash.merge_cells(f"B{start_r}:F{start_r}")
c = dash[f"B{start_r}"]
c.value = "PROYECCIÓN MAYO–DICIEMBRE 2026"
c.fill = fill(AZUL_MED)
c.font = font(bold=True, color=BLANCO, size=12)
c.alignment = center()
dash.row_dimensions[start_r].height = 22

headers_proj = ["Concepto", "Sin Plan", "Con Plan (Base)", "Con Plan (Opt.)", "Diferencia Plan"]
apply_header_row(dash, start_r + 1, range(2, 7), headers_proj, bg=AZUL_OSC, fg=BLANCO)
dash.row_dimensions[start_r + 1].height = 22

proj_rows = [
    ("Costo May-Dic 2026",   693810000,  580310000,  553770000,  None),
    ("Costo Total 2026",     993533600,  880033600,  853500000,  None),
    ("Ahorro vs. Sin Plan",  0,          113500000,  140030000,  None),
    ("KPI Costo/Ventas Dic", 0.038,      0.029,      0.028,      None),
    ("% Preventivo Dic",     0.30,       0.60,       0.65,       None),
    ("ROI del Plan",         None,       2.97,       None,       None),
]

tipo_cols = ["COP", "COP", "COP", "%", "%", "x"]

for i, (label, v_sin, v_con, v_opt, _) in enumerate(proj_rows):
    r = start_r + 2 + i
    dash.row_dimensions[r].height = 20
    bg = GRIS_CLAR if i % 2 == 0 else BLANCO

    c = dash.cell(row=r, column=2, value=label)
    c.fill = fill(bg); c.font = font(size=11); c.alignment = left_align(); c.border = border_thin()

    def set_proj_cell(col, val, tipo, good_if="high"):
        cx = dash.cell(row=r, column=col, value=val)
        cx.fill = fill(bg)
        cx.font = font(bold=(col == 4), size=11)
        cx.alignment = right()
        cx.border = border_thin()
        if tipo == "COP" and val:
            cx.number_format = '"$" #,##0'
        elif tipo == "%":
            cx.number_format = '0.0%'
        elif tipo == "x" and val:
            cx.number_format = '0.0"x"'

    set_proj_cell(3, v_sin, tipo_cols[i])
    set_proj_cell(4, v_con, tipo_cols[i])
    set_proj_cell(5, v_opt, tipo_cols[i])

    # Diferencia (col 6)
    if v_sin and v_con and tipo_cols[i] == "COP":
        diff = v_con - v_sin
        cd = dash.cell(row=r, column=6, value=diff)
        cd.fill = fill(VERDE_PALE if diff < 0 else ROJO_PALE)
        cd.font = font(bold=True, color=VERDE if diff < 0 else ROJO, size=11)
        cd.alignment = right()
        cd.border = border_thin()
        cd.number_format = '"$" #,##0'
    elif v_sin and v_con and tipo_cols[i] == "%":
        diff = v_con - v_sin
        cd = dash.cell(row=r, column=6, value=diff)
        cd.fill = fill(VERDE_PALE if diff < 0 else AMAR_PALE)
        cd.font = font(bold=True, color=VERDE if diff < 0 else AMARILLO, size=11)
        cd.alignment = right()
        cd.border = border_thin()
        cd.number_format = '0.0%'
    elif i == 5:  # ROI
        cd = dash.cell(row=r, column=6, value="Payback: Oct 2026")
        cd.fill = fill(VERDE_PALE)
        cd.font = font(bold=True, color=VERDE, size=11)
        cd.alignment = center()
        cd.border = border_thin()

# Nota al pie
dash.merge_cells("B24:F24")
c = dash["B24"]
c.value = "Fuente: Datos reales Ene-Abr 2026 | Informe de Costos CIF + Ejecución de Gastos | Ver hojas: HISTÓRICO · PROYECCIÓN · PALANCAS · SENSIBILIDAD · FLUJO"
c.fill = fill(AZUL_PALE)
c.font = font(color="595959", size=9, italic=True)
c.alignment = left_align()
dash.row_dimensions[24].height = 18

# ══════════════════════════════════════════════════════════════════════════════
# HOJA 2: HISTÓRICO
# ══════════════════════════════════════════════════════════════════════════════
hist = wb.create_sheet("HISTÓRICO")
hist.sheet_view.showGridLines = False
cols_hist = [2, 12, 18, 24, 30]
widths_hist = [1, 22, 18, 18, 18, 18, 18, 18, 18, 3]
labels_w = ["", "Concepto", "Enero", "Febrero", "Marzo", "Abril", "TOTAL Ene-Abr", "", "", ""]
for i, w in enumerate(widths_hist, 1):
    hist.column_dimensions[get_column_letter(i)].width = w

hist.merge_cells("B1:H1")
c = hist["B1"]
c.value = "HISTÓRICO DE COSTOS — ENERO A ABRIL 2026"
c.fill = fill(AZUL_OSC); c.font = font(bold=True, color=BLANCO, size=14); c.alignment = center()
hist.row_dimensions[1].height = 32

apply_header_row(hist, 2, [2,3,4,5,6,7], ["Concepto","Enero 2026","Febrero 2026","Marzo 2026","Abril 2026","TOTAL Ene-Abr"], bg=AZUL_MED, fg=BLANCO)
hist.row_dimensions[2].height = 24

hist_data = [
    ("CIF MANTENIMIENTO TOTAL",      53882400,  56912700,  85124600,  103803900, None),
    ("  — Mant. Metalmecánica",       22000000,  23000000,  30000000,   38000000, None),
    ("  — Mant. Inyección",           15000000,  17000000,  42000000,   32000000, None),
    ("  — Eléctrico / Instrumentación", 8000000,   9000000,  6000000,   16000000, None),
    ("  — Mant. Locativo",             5882400,   4912700,  4124600,    10803900, None),
    ("  — Otros CIF Mant.",            3000000,   3000000,  3000000,    7000000,  None),
    ("PRESUPUESTO CIF MANT.",         40000000,  45000000,  50000000,   50000000, None),
    ("DESVIACIÓN vs. PPTO",           13882400,  11912700,  35124600,   53803900, None),
    ("% DESVIACIÓN",                     0.347,     0.265,     0.703,      1.076, None),
    ("",None,None,None,None,None),
    ("PERSONAL FIJO — MANT.",         10200000,  10200000,  10200000,   10200000, None),
    ("PERSONAL TEMPORAL — MANT.",     11800000,  12900000,  14300000,   15200000, None),
    ("  — S.A.I TEMP",                 4500000,   5000000,   5500000,    5800000, None),
    ("  — VINCULAMOS",                 4300000,   4600000,   5200000,    5700000, None),
    ("  — GESTIÓN Y COMPROMISO",       3000000,   3300000,   3600000,    3700000, None),
    ("TOTAL PERSONAL MANT.",          22000000,  23100000,  24500000,   25400000, None),
    ("",None,None,None,None,None),
    ("KPI COSTO / VENTAS",              0.0294,    0.0337,    0.0491,     0.0515, None),
    ("VENTAS NETAS (REFERENCIA)",    1833412000, 1688892000, 1733680000, 2015480000, None),
]

for i, (label, e, f, m, a, _) in enumerate(hist_data):
    r = 3 + i
    hist.row_dimensions[r].height = 18
    is_total = label.isupper() and label != "" and not label.startswith(" ")
    bg = AZUL_PALE if is_total else (GRIS_CLAR if i % 2 == 0 else BLANCO)
    if label == "":
        continue

    c = hist.cell(row=r, column=2, value=label)
    c.fill = fill(bg); c.font = font(bold=is_total, size=10); c.alignment = left_align(); c.border = border_thin()

    vals = [e, f, m, a]
    for j, v in enumerate(vals):
        cv = hist.cell(row=r, column=3+j, value=v)
        cv.fill = fill(bg)
        cv.font = font(bold=is_total, size=10)
        cv.alignment = right()
        cv.border = border_thin()
        if label == "% DESVIACIÓN" or label == "KPI COSTO / VENTAS":
            cv.number_format = '0.0%'
            if v and v > 0.5:
                cv.fill = fill(ROJO_PALE)
                cv.font = font(bold=True, color=ROJO, size=10)
        elif v and abs(v) > 1000:
            cv.number_format = '"$" #,##0'
            if "DESVIACIÓN" in label and v and v > 0:
                cv.fill = fill(ROJO_PALE)
                cv.font = font(bold=True, color=ROJO, size=10)

    # Totales col 7
    if e and f and m and a and "%" not in label and "KPI" not in label:
        ct = hist.cell(row=r, column=7, value=e+f+m+a)
        ct.fill = fill(AZUL_CLAR if is_total else bg)
        ct.font = font(bold=True, size=10)
        ct.alignment = right()
        ct.border = border_thin()
        ct.number_format = '"$" #,##0'

# ══════════════════════════════════════════════════════════════════════════════
# HOJA 3: PROYECCIÓN
# ══════════════════════════════════════════════════════════════════════════════
proy = wb.create_sheet("PROYECCIÓN")
proy.sheet_view.showGridLines = False

meses = ["Mayo", "Junio", "Julio", "Agosto", "Sep.", "Oct.", "Nov.", "Dic."]
proy.column_dimensions["A"].width = 1
proy.column_dimensions["B"].width = 30
for i, m in enumerate(meses):
    proy.column_dimensions[get_column_letter(3+i)].width = 14
proy.column_dimensions[get_column_letter(11)].width = 16
proy.column_dimensions[get_column_letter(12)].width = 1

proy.merge_cells("B1:K1")
c = proy["B1"]
c.value = "PROYECCIÓN MENSUAL MAYO–DICIEMBRE 2026"
c.fill = fill(AZUL_OSC); c.font = font(bold=True, color=BLANCO, size=14); c.alignment = center()
proy.row_dimensions[1].height = 30

apply_header_row(proy, 2, range(2, 12),
    ["Concepto"] + meses + ["TOTAL May-Dic"],
    bg=AZUL_MED, fg=BLANCO)
proy.row_dimensions[2].height = 22

# Datos: base_2025, sin_plan, con_plan, presupuesto
base_25  = [62400000, 58200000, 71300000, 68900000, 65100000, 60400000, 72800000, 74600000]
sin_plan = [81120000, 75660000, 92690000, 89570000, 84630000, 78520000, 94640000, 96980000]
con_plan = [74420000, 68960000, 85990000, 79170000, 66980000, 59070000, 71690000, 74030000]
con_opt  = [73060000, 67610000, 84640000, 76690000, 62840000, 54790000, 65900000, 68240000]
ppto     = [47000000, 44000000, 52000000, 51000000, 49000000, 46000000, 53000000, 54000000]

sections = [
    ("BASE 2025 (Referencia)", base_25,  AZUL_PALE,  False, "COP"),
    ("SIN PLAN — Tendencia +30%", sin_plan, ROJO_PALE, True, "COP"),
    ("CON PLAN — Escenario Base", con_plan, VERDE_PALE, True, "COP"),
    ("CON PLAN — Escenario Optimista", con_opt, VERDE_PALE, False, "COP"),
    ("PRESUPUESTO APROBADO 2026", ppto, AMAR_PALE, False, "COP"),
    ("AHORRO: Sin Plan vs. Con Plan", [s-c for s,c in zip(sin_plan,con_plan)], VERDE_PALE, True, "COP"),
    ("% AHORRO", [(s-c)/s for s,c in zip(sin_plan,con_plan)], VERDE_PALE, False, "%"),
    ("DESVIACIÓN Sin Plan vs. Ppto", [s-p for s,p in zip(sin_plan,ppto)], ROJO_PALE, False, "COP"),
    ("DESVIACIÓN Con Plan vs. Ppto", [c-p for c,p in zip(con_plan,ppto)], AMAR_PALE, False, "COP"),
    ("KPI SIN PLAN (Costo/Ventas)", [v/2000000000 for v in sin_plan], ROJO_PALE, False, "%"),
    ("KPI CON PLAN (Costo/Ventas)", [v/2000000000 for v in con_plan], VERDE_PALE, False, "%"),
]

for i, (label, vals, bg, bold, tipo) in enumerate(sections):
    r = 3 + i
    proy.row_dimensions[r].height = 20
    c = proy.cell(row=r, column=2, value=label)
    c.fill = fill(bg if bold else (GRIS_CLAR if i%2==0 else BLANCO))
    c.font = font(bold=bold, size=10)
    c.alignment = left_align(); c.border = border_thin()

    total = sum(vals)
    for j, v in enumerate(vals):
        cv = proy.cell(row=r, column=3+j, value=v)
        actual_bg = bg if bold else (GRIS_CLAR if i%2==0 else BLANCO)
        cv.fill = fill(actual_bg)
        cv.font = font(bold=bold, size=10)
        cv.alignment = right()
        cv.border = border_thin()
        if tipo == "COP":
            cv.number_format = '"$" #,##0'
        else:
            cv.number_format = '0.0%'

    ct = proy.cell(row=r, column=11, value=total)
    ct.fill = fill(AZUL_CLAR if bold else (GRIS_CLAR if i%2==0 else BLANCO))
    ct.font = font(bold=True, size=10)
    ct.alignment = right(); ct.border = border_thin()
    ct.number_format = '"$" #,##0' if tipo == "COP" else '0.0%'

# Separador visual en fila 5 (entre Sin Plan y Con Plan)
for col in range(2, 12):
    c = proy.cell(row=6, column=col)
    c.border = Border(top=Side(style="medium", color=AZUL_MED))

# ── GRÁFICO DE BARRAS COMPARATIVO ─────────────────────────────────────────
chart = BarChart()
chart.type = "col"
chart.title = "Costo Mensual: Sin Plan vs. Con Plan vs. Presupuesto"
chart.style = 10
chart.y_axis.title = "COP"
chart.x_axis.title = "Mes"
chart.y_axis.numFmt = '#,##0'
chart.shape = 4
chart.height = 14
chart.width = 26

# Sin Plan (row 4), Con Plan (row 5), Presupuesto (row 7)
data_sin  = Reference(proy, min_col=3, max_col=10, min_row=4, max_row=4)
data_con  = Reference(proy, min_col=3, max_col=10, min_row=5, max_row=5)
data_ppto = Reference(proy, min_col=3, max_col=10, min_row=7, max_row=7)
cats      = Reference(proy, min_col=3, max_col=10, min_row=2)

from openpyxl.chart import Series
s1 = Series(data_sin,  title="Sin Plan")
s2 = Series(data_con,  title="Con Plan")
s3 = Series(data_ppto, title="Presupuesto")
chart.series.append(s1)
chart.series.append(s2)
chart.series.append(s3)
chart.set_categories(cats)

# colores de barras
chart.series[0].graphicalProperties.solidFill = ROJO
chart.series[1].graphicalProperties.solidFill = VERDE_MED
chart.series[2].graphicalProperties.solidFill = AMARILLO

chart.legend.position = 'b'
proy.add_chart(chart, "B16")

# ══════════════════════════════════════════════════════════════════════════════
# HOJA 4: PALANCAS DE AHORRO
# ══════════════════════════════════════════════════════════════════════════════
pal = wb.create_sheet("PALANCAS")
pal.sheet_view.showGridLines = False
pal.column_dimensions["A"].width = 1
pal.column_dimensions["B"].width = 36
for i in range(8):
    pal.column_dimensions[get_column_letter(3+i)].width = 12
pal.column_dimensions[get_column_letter(11)].width = 16
pal.column_dimensions[get_column_letter(12)].width = 16

pal.merge_cells("B1:L1")
c = pal["B1"]
c.value = "PALANCAS DE AHORRO — DETALLE MENSUAL (COP)"
c.fill = fill(AZUL_OSC); c.font = font(bold=True, color=BLANCO, size=14); c.alignment = center()
pal.row_dimensions[1].height = 30

apply_header_row(pal, 2, range(2, 13),
    ["Palanca de Ahorro"] + meses + ["TOTAL", "Certeza"],
    bg=AZUL_MED, fg=BLANCO)
pal.row_dimensions[2].height = 22

palancas = [
    ("1 — Control OT (gasto no autorizado)",
     [4000000]*8, VERDE_PALE, "Alta"),
    ("2 — Reducción Personal Temporal (-25%)",
     [2700000,2700000,2700000,4200000,4200000,4200000,4200000,4200000], VERDE_PALE, "Alta"),
    ("3 — Contratos Proveedores (1a ronda -12%)",
     [0,0,0,2200000,2200000,0,0,0], VERDE_PALE, "Media-Alta"),
    ("4 — Contratos Proveedores (2a ronda -18%)",
     [0,0,0,0,0,4000000,4000000,4000000], VERDE_PALE, "Media-Alta"),
    ("5 — Stock Repuestos (elimina recargo 30%)",
     [0,0,0,0,2250000,2250000,2250000,2250000], VERDE_PALE, "Alta"),
    ("6 — Plan Preventivo — Correctivos evitados",
     [0,0,0,0,5000000,5000000,8500000,8500000], VERDE_PALE, "Media"),
    ("TOTAL AHORRO MENSUAL",
     None, AZUL_CLAR, ""),
    ("ACUMULADO AL MES",
     None, AZUL_PALE, ""),
]

raw_vals = []
for label, vals, bg, cert in palancas[:-2]:
    raw_vals.append(vals)

totals_monthly = [sum(raw_vals[j][i] for j in range(len(raw_vals))) for i in range(8)]
cumul = [sum(totals_monthly[:i+1]) for i in range(8)]

all_rows = palancas[:-2] + [
    ("TOTAL AHORRO MENSUAL", totals_monthly, AZUL_CLAR, ""),
    ("ACUMULADO AL MES", cumul, AZUL_PALE, ""),
]

for i, (label, vals, bg, cert) in enumerate(all_rows):
    r = 3 + i
    pal.row_dimensions[r].height = 20
    is_total = "TOTAL" in label or "ACUMULADO" in label

    c = pal.cell(row=r, column=2, value=label)
    c.fill = fill(bg)
    c.font = font(bold=is_total, size=10)
    c.alignment = left_align(); c.border = border_thin()

    total_row = 0
    for j, v in enumerate(vals):
        cv = pal.cell(row=r, column=3+j, value=v if v > 0 else None)
        cv.fill = fill(bg)
        cv.font = font(bold=is_total, size=10)
        cv.alignment = right(); cv.border = border_thin()
        cv.number_format = '"$" #,##0'
        if v: total_row += v

    ct = pal.cell(row=r, column=11, value=total_row if total_row > 0 else None)
    ct.fill = fill(AZUL_CLAR if is_total else bg)
    ct.font = font(bold=True, color=VERDE if is_total else NEGRO, size=10)
    ct.alignment = right(); ct.border = border_thin()
    ct.number_format = '"$" #,##0'

    ccert = pal.cell(row=r, column=12, value=cert)
    cert_bg = {"Alta": VERDE_PALE, "Media-Alta": AMAR_PALE, "Media": AMAR_PALE, "": AZUL_CLAR}.get(cert, BLANCO)
    cert_color = {"Alta": VERDE, "Media-Alta": "7F6000", "Media": "7F6000", "": AZUL_OSC}.get(cert, NEGRO)
    ccert.fill = fill(cert_bg)
    ccert.font = font(bold=False, color=cert_color, size=10)
    ccert.alignment = center(); ccert.border = border_thin()

# Resumen financiero debajo
gap = 3 + len(all_rows) + 2
pal.merge_cells(f"B{gap}:L{gap}")
c = pal[f"B{gap}"]
c.value = "RESUMEN FINANCIERO DEL PLAN"
c.fill = fill(AZUL_MED); c.font = font(bold=True, color=BLANCO, size=12); c.alignment = center()
pal.row_dimensions[gap].height = 22

resumen = [
    ("Total ahorros/costos evitados May-Dic 2026", 113450000),
    ("Inversión total requerida (stock + CMMS + capacitación)", -53000000),
    ("BENEFICIO NETO 2026", 60450000),
    ("ROI del Plan (May-Dic 2026)", 2.97),
    ("Payback de la inversión", "Octubre 2026"),
]
for j, (lbl, val) in enumerate(resumen):
    r = gap + 1 + j
    pal.row_dimensions[r].height = 20
    is_net = "NETO" in lbl or "ROI" in lbl
    bg = VERDE_PALE if is_net else (GRIS_CLAR if j%2==0 else BLANCO)

    cl = pal.cell(row=r, column=2, value=lbl)
    cl.fill = fill(bg); cl.font = font(bold=is_net, size=11)
    cl.alignment = left_align(); cl.border = border_thin()

    cv = pal.cell(row=r, column=3, value=val)
    cv.fill = fill(bg); cv.font = font(bold=is_net, color=VERDE if is_net else NEGRO, size=11)
    cv.alignment = right(); cv.border = border_thin()
    if isinstance(val, float) and val < 10:
        cv.number_format = '0.0"x"'
    elif isinstance(val, int):
        cv.number_format = '"$" #,##0'

    for col in range(4, 12):
        cx = pal.cell(row=r, column=col)
        cx.fill = fill(bg); cx.border = border_thin()

# ══════════════════════════════════════════════════════════════════════════════
# HOJA 5: SENSIBILIDAD
# ══════════════════════════════════════════════════════════════════════════════
sens = wb.create_sheet("SENSIBILIDAD")
sens.sheet_view.showGridLines = False
sens.column_dimensions["A"].width = 1
sens.column_dimensions["B"].width = 36
for i in range(8):
    sens.column_dimensions[get_column_letter(3+i)].width = 13
sens.column_dimensions[get_column_letter(11)].width = 16

sens.merge_cells("B1:K1")
c = sens["B1"]
c.value = "ANÁLISIS DE SENSIBILIDAD — 3 ESCENARIOS"
c.fill = fill(AZUL_OSC); c.font = font(bold=True, color=BLANCO, size=14); c.alignment = center()
sens.row_dimensions[1].height = 30

# Descripción escenarios
escenarios_desc = [
    ("ESCENARIO OPTIMISTA",  "Todas las palancas al 120% de efectividad. Plan ejecutado sin retrasos.", VERDE_PALE, VERDE),
    ("ESCENARIO BASE",       "Plan ejecutado según cronograma. Efectividad promedio proyectada.", AZUL_PALE, AZUL_MED),
    ("ESCENARIO CONSERVADOR","Fases 2 y 3 con retraso de 2 meses. Efectividad al 70%.", AMAR_PALE, "7F6000"),
    ("SIN PLAN (Referencia)","No se ejecuta ninguna acción. Tendencia actual sostenida.", ROJO_PALE, ROJO),
]

for i, (nombre, desc, bg, color) in enumerate(escenarios_desc):
    r = 2 + i*2
    sens.merge_cells(f"B{r}:K{r}")
    c = sens[f"B{r}"]
    c.value = f"{nombre}: {desc}"
    c.fill = fill(bg); c.font = font(bold=True, color=color, size=10)
    c.alignment = left_align(); sens.row_dimensions[r].height = 18

apply_header_row(sens, 11, range(2, 12),
    ["Escenario"] + meses + ["TOTAL May-Dic"],
    bg=AZUL_MED, fg=BLANCO)
sens.row_dimensions[11].height = 22

opt_vals  = [73060000, 67610000, 84640000, 76690000, 62840000, 54790000, 65900000, 68240000]
base_vals = [74420000, 68960000, 85990000, 79170000, 66980000, 59070000, 71690000, 74030000]
cons_vals = [74420000, 68960000, 85990000, 86870000, 78430000, 71330000, 77240000, 77980000]
sinp_vals = [81120000, 75660000, 92690000, 89570000, 84630000, 78520000, 94640000, 96980000]

scen_rows = [
    ("Optimista (+20% efectividad)",     opt_vals,  VERDE_PALE, VERDE),
    ("Base (cronograma normal)",         base_vals, AZUL_PALE, AZUL_MED),
    ("Conservador (retraso 2 meses)",    cons_vals, AMAR_PALE, "7F6000"),
    ("Sin Plan (tendencia actual)",      sinp_vals, ROJO_PALE, ROJO),
    ("", None, BLANCO, NEGRO),
    ("Ahorro Optimista vs. Sin Plan",    [sinp_vals[j]-opt_vals[j]  for j in range(8)], VERDE_PALE, VERDE),
    ("Ahorro Base vs. Sin Plan",         [sinp_vals[j]-base_vals[j] for j in range(8)], VERDE_PALE, VERDE),
    ("Ahorro Conservador vs. Sin Plan",  [sinp_vals[j]-cons_vals[j] for j in range(8)], AMAR_PALE, "7F6000"),
]

for i, (label, vals, bg, color) in enumerate(scen_rows):
    r = 12 + i
    sens.row_dimensions[r].height = 20
    if not label:
        continue
    is_total = label.startswith("Ahorro")

    c = sens.cell(row=r, column=2, value=label)
    c.fill = fill(bg); c.font = font(bold="Sin Plan" in label or is_total, color=color, size=10)
    c.alignment = left_align(); c.border = border_thin()

    total = sum(vals) if vals else 0
    for j, v in enumerate(vals):
        cv = sens.cell(row=r, column=3+j, value=v)
        cv.fill = fill(bg); cv.font = font(color=color, size=10)
        cv.alignment = right(); cv.border = border_thin()
        cv.number_format = '"$" #,##0'

    ct = sens.cell(row=r, column=11, value=total)
    ct.fill = fill(AZUL_CLAR if is_total else bg)
    ct.font = font(bold=True, color=VERDE if is_total else color, size=10)
    ct.alignment = right(); ct.border = border_thin()
    ct.number_format = '"$" #,##0'

# ══════════════════════════════════════════════════════════════════════════════
# HOJA 6: FLUJO (Inversión vs. Ahorro)
# ══════════════════════════════════════════════════════════════════════════════
flujo = wb.create_sheet("FLUJO DE CAJA")
flujo.sheet_view.showGridLines = False
flujo.column_dimensions["A"].width = 1
flujo.column_dimensions["B"].width = 36
for i in range(8):
    flujo.column_dimensions[get_column_letter(3+i)].width = 12
flujo.column_dimensions[get_column_letter(11)].width = 15

flujo.merge_cells("B1:K1")
c = flujo["B1"]
c.value = "FLUJO DE INVERSIÓN vs. AHORRO — MAYO A DICIEMBRE 2026"
c.fill = fill(AZUL_OSC); c.font = font(bold=True, color=BLANCO, size=14); c.alignment = center()
flujo.row_dimensions[1].height = 30

apply_header_row(flujo, 2, range(2, 12),
    ["Concepto"] + meses + ["TOTAL"],
    bg=AZUL_MED, fg=BLANCO)
flujo.row_dimensions[2].height = 22

inv_mes  = [0, 5000000, 3000000, 30000000, 15000000, 0, 0, 0]
ahor_mes = [6700000, 6700000, 6700000, 10400000, 17650000, 19450000, 22950000, 22950000]
flujo_neto = [a - i for a, i in zip(ahor_mes, inv_mes)]
acumulado  = [sum(flujo_neto[:j+1]) for j in range(8)]

flujo_rows = [
    ("INVERSIÓN MENSUAL (desembolso)", inv_mes,   ROJO_PALE,   True,  "inv"),
    ("AHORRO / COSTO EVITADO MENSUAL", ahor_mes,  VERDE_PALE,  True,  "ahor"),
    ("FLUJO NETO MENSUAL",             flujo_neto,AZUL_PALE,   True,  "neto"),
    ("FLUJO NETO ACUMULADO",           acumulado, AZUL_CLAR,   True,  "acum"),
]

descrip_inv = {
    1: "",
    2: "CMMS inicio",
    3: "Capacitación RCM",
    4: "Stock repuestos (67%)",
    5: "Stock repuestos (33%)",
    6: "", 7: "", 8: "",
}

for i, (label, vals, bg, bold, tipo) in enumerate(flujo_rows):
    r = 3 + i
    flujo.row_dimensions[r].height = 22
    c = flujo.cell(row=r, column=2, value=label)
    c.fill = fill(bg); c.font = font(bold=bold, size=11)
    c.alignment = left_align(); c.border = border_thin()

    total = sum(vals)
    for j, v in enumerate(vals):
        cv = flujo.cell(row=r, column=3+j, value=v if v != 0 else None)
        actual_bg = bg
        if tipo == "neto" and v < 0:
            actual_bg = ROJO_PALE
            cv.font = font(bold=True, color=ROJO, size=11)
        elif tipo == "neto" and v >= 0:
            cv.font = font(bold=True, color=VERDE, size=11)
        elif tipo == "acum" and v < 0:
            actual_bg = ROJO_PALE
            cv.font = font(bold=True, color=ROJO, size=11)
        elif tipo == "acum" and v >= 0:
            cv.font = font(bold=True, color=VERDE, size=11)
        else:
            cv.font = font(bold=bold, size=11)
        cv.fill = fill(actual_bg)
        cv.alignment = right(); cv.border = border_thin()
        cv.number_format = '"$" #,##0'

    ct = flujo.cell(row=r, column=11, value=total if total != 0 else None)
    ct_bg = VERDE_PALE if (tipo in ["ahor","neto","acum"] and total > 0) else (ROJO_PALE if total < 0 else AZUL_CLAR)
    ct.fill = fill(ct_bg)
    ct.font = font(bold=True, color=VERDE if total > 0 else ROJO, size=11)
    ct.alignment = right(); ct.border = border_thin()
    ct.number_format = '"$" #,##0'

# Notas de inversión
flujo.row_dimensions[8].height = 14
for j, desc in descrip_inv.items():
    if desc:
        c = flujo.cell(row=8, column=2+j, value=f"↑ {desc}")
        c.font = font(color=ROJO, size=8, italic=True)
        c.alignment = center()

# Nota payback
flujo.merge_cells("B10:K10")
c = flujo["B10"]
c.value = "PAYBACK: La inversión total de $53,000,000 queda recuperada en OCTUBRE 2026 (mes 6 del plan)."
c.fill = fill(VERDE_PALE)
c.font = font(bold=True, color=VERDE, size=11)
c.alignment = center(); flujo.row_dimensions[10].height = 22

# Gráfico de línea — flujo acumulado
lc = LineChart()
lc.title = "Flujo Neto Acumulado del Plan (COP)"
lc.style = 10
lc.y_axis.title = "COP"
lc.x_axis.title = "Mes"
lc.y_axis.numFmt = '#,##0'
lc.height = 12
lc.width = 22

data_acum = Reference(flujo, min_col=3, max_col=10, min_row=6, max_row=6)
cats_f = Reference(flujo, min_col=3, max_col=10, min_row=2)
s_acum = Series(data_acum, title="Flujo Acumulado")
lc.series.append(s_acum)
lc.set_categories(cats_f)
lc.series[0].graphicalProperties.line.solidFill = VERDE_MED
lc.series[0].graphicalProperties.line.width = 28000
lc.legend.position = 'b'
flujo.add_chart(lc, "B12")

# ══════════════════════════════════════════════════════════════════════════════
# HOJA 7: KPIs DE SEGUIMIENTO
# ══════════════════════════════════════════════════════════════════════════════
kpi_sh = wb.create_sheet("KPIs SEGUIMIENTO")
kpi_sh.sheet_view.showGridLines = False
kpi_sh.column_dimensions["A"].width = 1
kpi_sh.column_dimensions["B"].width = 36
kpi_sh.column_dimensions["C"].width = 16
kpi_sh.column_dimensions["D"].width = 16
for i in range(8):
    kpi_sh.column_dimensions[get_column_letter(5+i)].width = 12
kpi_sh.column_dimensions[get_column_letter(13)].width = 1

kpi_sh.merge_cells("B1:L1")
c = kpi_sh["B1"]
c.value = "TABLERO DE KPIs — SEGUIMIENTO MENSUAL"
c.fill = fill(AZUL_OSC); c.font = font(bold=True, color=BLANCO, size=14); c.alignment = center()
kpi_sh.row_dimensions[1].height = 30

apply_header_row(kpi_sh, 2, range(2, 13),
    ["KPI", "Meta Dic 2026", "Real Abr 2026"] + meses,
    bg=AZUL_MED, fg=BLANCO)
kpi_sh.row_dimensions[2].height = 22

kpis_track = [
    ("Costo total mant. mensual (COP)",   "≤ $85,000,000",  103803900,  [None]*8, "COP"),
    ("% Ejecución vs. presupuesto",       "≤ 100%",          2.076,     [None]*8, "%"),
    ("% Mantenimiento Preventivo",        "≥ 60%",           0.30,      [None]*8, "%"),
    ("Costo personal temporal (COP)",     "≤ $7,500,000",   15200000,   [None]*8, "COP"),
    ("Nro. proveedores con contrato",     "5",               0,         [None]*8, "num"),
    ("% OT con aprobación previa",        "100%",            0.30,      [None]*8, "%"),
    ("KPI Costo/Ventas (mensual)",        "≤ 2.5%",          0.0515,    [None]*8, "%"),
    ("Paros por falta de repuesto",       "0",               None,      [None]*8, "num"),
    ("MTBF equipos críticos (horas)",     "Mejorar",         None,      [None]*8, "num"),
]

# Metas mensuales (con plan)
metas_mens = {
    0: [130e6,  110e6,  105e6,  100e6,  95e6,  90e6,  90e6,  85e6],    # Costo
    1: [1.10,   1.05,   1.05,   1.00,   1.00,  1.00,  1.00,  1.00],    # % ej
    2: [0.35,   0.40,   0.45,   0.48,   0.52,  0.55,  0.58,  0.60],    # % prev
    3: [10.5e6, 9e6,    8.5e6,  8.0e6,  8.0e6, 8.0e6, 7.5e6, 7.5e6],  # temporal
    4: [0,1,2,3,5,5,5,5],                                                # contratos
    5: [0.60,0.90,0.95,0.98,1.00,1.00,1.00,1.00],                       # OT
    6: [0.035,0.032,0.030,0.028,0.027,0.026,0.0255,0.025],             # KPI
    7: [None]*8,
    8: [None]*8,
}

for i, (kpi_name, meta, real_abr, _, tipo) in enumerate(kpis_track):
    r = 3 + i
    kpi_sh.row_dimensions[r].height = 20
    bg = GRIS_CLAR if i % 2 == 0 else BLANCO

    ck = kpi_sh.cell(row=r, column=2, value=kpi_name)
    ck.fill = fill(bg); ck.font = font(size=10); ck.alignment = left_align(); ck.border = border_thin()

    cm = kpi_sh.cell(row=r, column=3, value=meta)
    cm.fill = fill(VERDE_PALE); cm.font = font(bold=True, color=VERDE, size=10)
    cm.alignment = center(); cm.border = border_thin()

    cr = kpi_sh.cell(row=r, column=4, value=real_abr)
    if real_abr:
        cr.fill = fill(ROJO_PALE); cr.font = font(bold=True, color=ROJO, size=10)
    else:
        cr.fill = fill(bg); cr.font = font(size=10)
    cr.alignment = right(); cr.border = border_thin()
    if tipo == "COP" and real_abr:
        cr.number_format = '"$" #,##0'
    elif tipo == "%":
        cr.number_format = '0.0%'

    # Metas mensuales (para que el usuario las vaya rellenando)
    for j in range(8):
        cv = kpi_sh.cell(row=r, column=5+j, value=metas_mens[i][j] if metas_mens[i][j] else None)
        cv.fill = fill(AZUL_PALE)
        cv.font = font(color="595959", size=9, italic=True)
        cv.alignment = right(); cv.border = border_thin()
        if tipo == "COP" and metas_mens[i][j]:
            cv.number_format = '"$" #,##0'
        elif tipo == "%" and metas_mens[i][j]:
            cv.number_format = '0.0%'
        elif tipo == "num":
            cv.number_format = '0'

# Instrucción de uso
kpi_sh.merge_cells("B14:L14")
c = kpi_sh["B14"]
c.value = "INSTRUCCIÓN: Reemplace las celdas en AZUL con los valores reales mensuales al cierre de cada período. Las celdas en VERDE son metas fijas."
c.fill = fill(AZUL_PALE)
c.font = font(color=AZUL_OSC, size=9, italic=True)
c.alignment = left_align(); kpi_sh.row_dimensions[14].height = 18

# ── ORDEN DE HOJAS Y PROPIEDADES FINALES ─────────────────────────────────────
tab_colors = {
    "DASHBOARD":       "1F3864",
    "HISTÓRICO":       "2E75B6",
    "PROYECCIÓN":      "C00000",
    "PALANCAS":        "375623",
    "SENSIBILIDAD":    "7030A0",
    "FLUJO DE CAJA":   "375623",
    "KPIs SEGUIMIENTO":"FFC000",
}
for sheet_name, color in tab_colors.items():
    if sheet_name in wb.sheetnames:
        wb[sheet_name].sheet_properties.tabColor = color

wb.active = wb["DASHBOARD"]

output_path = "/home/user/Todos-financieros/Modelo_Financiero_Mantenimiento_2026.xlsx"
wb.save(output_path)
print(f"Excel guardado: {output_path}")
print(f"Hojas: {wb.sheetnames}")
