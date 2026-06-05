#!/usr/bin/env python3
"""
Rutina SEO Semanal — Galería Duque Arango
Genera Inteligencia_SEO_2026-06-05.xlsx con 5 pestañas en español.
"""

import xlsxwriter
import os

DATE = "2026-06-05"
FILENAME = f"Inteligencia_SEO_{DATE}.xlsx"
FILEPATH = os.path.join("/home/user/da-routines", FILENAME)

wb = xlsxwriter.Workbook(FILEPATH)

# ── Formatos globales ──────────────────────────────────────────────────────────
fmt_title = wb.add_format({
    "bold": True, "font_size": 14, "font_color": "#FFFFFF",
    "bg_color": "#1A1A2E", "align": "center", "valign": "vcenter",
    "border": 1
})
fmt_header = wb.add_format({
    "bold": True, "font_size": 10, "font_color": "#FFFFFF",
    "bg_color": "#16213E", "align": "center", "valign": "vcenter",
    "text_wrap": True, "border": 1
})
fmt_subheader = wb.add_format({
    "bold": True, "font_size": 10, "font_color": "#FFFFFF",
    "bg_color": "#0F3460", "align": "left", "valign": "vcenter",
    "border": 1
})
fmt_data = wb.add_format({
    "font_size": 9, "align": "left", "valign": "vcenter",
    "border": 1, "text_wrap": True
})
fmt_data_center = wb.add_format({
    "font_size": 9, "align": "center", "valign": "vcenter", "border": 1
})
fmt_number = wb.add_format({
    "font_size": 9, "align": "right", "valign": "vcenter",
    "border": 1, "num_format": "#,##0"
})
fmt_highlight = wb.add_format({
    "bold": True, "font_size": 9, "font_color": "#FFFFFF",
    "bg_color": "#E94560", "align": "center", "valign": "vcenter", "border": 1
})
fmt_highlight_gold = wb.add_format({
    "bold": True, "font_size": 9, "font_color": "#1A1A2E",
    "bg_color": "#F5C518", "align": "center", "valign": "vcenter", "border": 1
})
fmt_alta = wb.add_format({
    "bold": True, "font_size": 9, "font_color": "#FFFFFF",
    "bg_color": "#C0392B", "align": "center", "valign": "vcenter", "border": 1
})
fmt_media = wb.add_format({
    "bold": True, "font_size": 9, "font_color": "#FFFFFF",
    "bg_color": "#E67E22", "align": "center", "valign": "vcenter", "border": 1
})
fmt_baja = wb.add_format({
    "font_size": 9, "font_color": "#2C3E50",
    "bg_color": "#D5DBDB", "align": "center", "valign": "vcenter", "border": 1
})
fmt_note = wb.add_format({
    "font_size": 8, "italic": True, "font_color": "#7F8C8D",
    "align": "left", "valign": "vcenter", "border": 0
})
fmt_duque = wb.add_format({
    "bold": True, "font_size": 9, "font_color": "#FFFFFF",
    "bg_color": "#1A6B3C", "align": "left", "valign": "vcenter", "border": 1
})
fmt_duque_num = wb.add_format({
    "bold": True, "font_size": 9, "font_color": "#FFFFFF",
    "bg_color": "#1A6B3C", "align": "right", "valign": "vcenter",
    "border": 1, "num_format": "#,##0"
})
fmt_section = wb.add_format({
    "bold": True, "font_size": 10, "font_color": "#FFFFFF",
    "bg_color": "#0F3460", "align": "left", "valign": "vcenter", "border": 1
})
fmt_url = wb.add_format({
    "font_size": 8, "font_color": "#2980B9", "align": "left",
    "valign": "vcenter", "border": 1, "underline": True
})

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1: DESEMPEÑO PROPIO
# ══════════════════════════════════════════════════════════════════════════════
ws1 = wb.add_worksheet("Desempeño Propio")
ws1.set_tab_color("#1A6B3C")
ws1.set_zoom(85)

col_widths_t1 = [20, 22, 18, 45]
for i, w in enumerate(col_widths_t1):
    ws1.set_column(i, i, w)
ws1.set_row(0, 28)
ws1.set_row(1, 14)

ws1.merge_range("A1:D1",
    "DESEMPEÑO PROPIO — galeriaduquearango.com | Datos al junio 2026",
    fmt_title)
ws1.write("A2", "Fuente: Semrush (base de datos CO y US) | Rutina SEO Semanal", fmt_note)

# ── Métricas generales ─────────────────────────────────────────────────────
ws1.merge_range("A4:D4", "MÉTRICAS GENERALES DEL DOMINIO", fmt_section)
ws1.write_row("A5", ["Base de Datos", "Métrica", "Valor", "Notas"], fmt_header)

metricas_generales = [
    ["Colombia (CO)", "Keywords orgánicas activas", "2.255",
     "Semrush ranking nacional: #4.908"],
    ["Colombia (CO)", "Tráfico orgánico estimado/mes", "21.691",
     "Líder absoluto entre las 17 galerías analizadas"],
    ["Colombia (CO)", "Costo orgánico equivalente (USD)", "$886",
     "Valor estimado del tráfico si fuera pagado"],
    ["Colombia (CO)", "Keywords en top 10", "~12",
     "Basado en posiciones reportadas en domain_organic"],
    ["Estados Unidos (US)", "Keywords orgánicas activas", "1.895",
     "Semrush ranking global: #605.349"],
    ["Estados Unidos (US)", "Tráfico orgánico estimado/mes", "2.228",
     "#3 entre las 17 galerías en el mercado US"],
    ["Estados Unidos (US)", "Costo orgánico equivalente (USD)", "$465",
     "Oportunidad de crecimiento significativa vs operagallery.com"],
    ["Estados Unidos (US)", "Keywords en top 10", "~5",
     "Principalmente Botero, Guayasamín, artistas modernos"],
]
for r, row in enumerate(metricas_generales):
    ws1.write(5 + r, 0, row[0], fmt_data_center)
    ws1.write(5 + r, 1, row[1], fmt_data)
    ws1.write(5 + r, 2, row[2], fmt_highlight_gold if r in [1, 4] else fmt_data_center)
    ws1.write(5 + r, 3, row[3], fmt_data)

# ── Top 20 Keywords CO ─────────────────────────────────────────────────────
ws1.merge_range("A15:D15", "TOP 20 KEYWORDS ORGÁNICAS — COLOMBIA (CO)", fmt_section)
ws1.write_row("A16", ["Keyword", "Posición", "Volumen de Búsqueda", "URL que Rankea"], fmt_header)

top20_co = [
    ["edgar negret", 2, "90.500", "galeriaduquearango.com/artista/edgar-negret/"],
    ["plaza botero medellín antioquia", 5, "22.200", "galeriaduquearango.com/blog/plaza-botero-..."],
    ["plaza botero", 5, "49.500", "galeriaduquearango.com/blog/plaza-botero-..."],
    ["galeria duque arango", 1, "720", "galeriaduquearango.com/"],
    ["omar rayo", 3, "5.400", "galeriaduquearango.com/blog/omar-rayo-..."],
    ["obras de alejandro obregón", 1, "390", "galeriaduquearango.com/artista/alejandro-obregon/"],
    ["david manzur", 2, "4.400", "galeriaduquearango.com/artista/david-manzur/"],
    ["edgar negret (blog)", 13, "90.500", "galeriaduquearango.com/blog/edgar-negret-..."],
    ["enrique grau obras", 1, "260", "galeriaduquearango.com/artista/enrique-grau/"],
    ["enrique grau", 2, "1.600", "galeriaduquearango.com/artista/enrique-grau/"],
    ["ana mercedes hoyos obras", 1, "210", "galeriaduquearango.com/artista/ana-mercedes-hoyos/"],
    ["pinturas de alejandro obregón", 1, "210", "galeriaduquearango.com/artista/alejandro-obregon/"],
    ["alejandro obregon obras", 1, "210", "galeriaduquearango.com/artista/alejandro-obregon/"],
    ["obras de arte de david manzur", 1, "170", "galeriaduquearango.com/artista/david-manzur/"],
    ["botero cuadros", 1, "170", "galeriaduquearango.com/blog/las-obras-mejor-vendidas-de-..."],
    ["edgar negret obras", 1, "140", "galeriaduquearango.com/artista/edgar-negret/"],
    ["grau", 2, "1.300", "galeriaduquearango.com/artista/enrique-grau/"],
    ["muralismo", 3, "1.600", "galeriaduquearango.com/blog/muralismo-latinoamericano-..."],
    ["obregón pinturas", 2, "480", "galeriaduquearango.com/artista/alejandro-obregon/"],
    ["cuadros de botero", 1, "480", "galeriaduquearango.com/blog/las-obras-mejor-vendidas-de-..."],
]
for r, row in enumerate(top20_co):
    ws1.write(16 + r, 0, row[0], fmt_data)
    ws1.write(16 + r, 1, row[1], fmt_data_center)
    ws1.write(16 + r, 2, row[2], fmt_number)
    ws1.write(16 + r, 3, row[3], fmt_data)

# ── Top 20 Keywords US ─────────────────────────────────────────────────────
ws1.merge_range("A38:D38", "TOP 20 KEYWORDS ORGÁNICAS — ESTADOS UNIDOS (US)", fmt_section)
ws1.write_row("A39", ["Keyword", "Posición", "Volumen de Búsqueda", "URL que Rankea"], fmt_header)

top20_us = [
    ["botero's", 6, "3.600", "galeriaduquearango.com/en/blog/all-you-need-to-know-..."],
    ["ana mercedes hoyos", 6, "3.600", "galeriaduquearango.com/artista/ana-mercedes-hoyos/"],
    ["fernando botero", 8, "18.100", "galeriaduquearango.com/en/blog/unmistakable-brand-boterism/"],
    ["colombian painters", 3, "720", "galeriaduquearango.com/en/blog/discovering-colombian-art-..."],
    ["plaza botero medellín colombia", 6, "1.900", "galeriaduquearango.com/en/blog/plaza-botero-..."],
    ["oswaldo guayasamin", 7, "1.600", "galeriaduquearango.com/en/blog/oswaldo-guayasamins-legacy-..."],
    ["guayasamin", 5, "1.300", "galeriaduquearango.com/en/blog/the-works-of-oswaldo-guayasamin/"],
    ["fernando botero artworks", 9, "1.300", "galeriaduquearango.com/en/blog/all-you-need-to-know-..."],
    ["fernando botero style", 1, "170", "galeriaduquearango.com/en/blog/unmistakable-brand-boterism/"],
    ["boterismo", 3, "480", "galeriaduquearango.com/en/blog/unmistakable-brand-boterism/"],
    ["picasso realistic art", 4, "480", "galeriaduquearango.com/en/blog/artistic-periods-of-pablo-picasso/"],
    ["fernando botero art", 8, "3.600", "galeriaduquearango.com/en/blog/unmistakable-brand-boterism/"],
    ["nombre 30 pinturas fernando botero", 2, "210", "galeriaduquearango.com/blog/las-obras-mejor-vendidas-..."],
    ["botero", 12, "9.900", "galeriaduquearango.com/en/blog/all-you-need-to-know-..."],
    ["fernando botero paintings", 13, "4.400", "galeriaduquearango.com/en/blog/the-works-of-fernando-botero-..."],
    ["mexican art", 24, "8.100", "galeriaduquearango.com/en/blog/mexican-art-history-..."],
    ["fernando botero still life", 1, "50", "galeriaduquearango.com/en/blog/the-rich-volumes-of-..."],
    ["fernando botero (alt)", 20, "18.100", "galeriaduquearango.com/en/blog/all-you-need-to-know-..."],
    ["mona lisa age twelve", 7, "480", "galeriaduquearango.com/en/blog/fernando-botero-monalisa-..."],
    ["fernando de szyszlo", 5, "480", "galeriaduquearango.com/en/artist/fernando-de-szyszlo/"],
]
for r, row in enumerate(top20_us):
    ws1.write(39 + r, 0, row[0], fmt_data)
    ws1.write(39 + r, 1, row[1], fmt_data_center)
    ws1.write(39 + r, 2, row[2], fmt_number)
    ws1.write(39 + r, 3, row[3], fmt_data)

# ── Top 10 Páginas CO ──────────────────────────────────────────────────────
ws1.merge_range("A61:D61", "TOP 10 PÁGINAS POR TRÁFICO ORGÁNICO — COLOMBIA (CO)", fmt_section)
ws1.write_row("A62", ["URL", "Tráfico Estimado", "# Keywords", "% del Tráfico Total"], fmt_header)

top10_pages_co = [
    ["/artista/edgar-negret/", "8.782", "27", "40,5 %"],
    ["/artista/david-manzur/", "2.697", "50", "12,4 %"],
    ["/blog/plaza-botero-resignificar-el-espacio-publico-...", "1.899", "33", "8,8 %"],
    ["/blog/las-obras-de-fernando-botero-y-su-significado/", "1.533", "141", "7,1 %"],
    ["/artista/alejandro-obregon/", "870", "69", "4,0 %"],
    ["/", "808", "109", "3,7 %"],
    ["/blog/artistas-colombianos-que-debes-conocer/", "514", "183", "2,4 %"],
    ["/blog/las-obras-mejor-vendidas-de-fernando-botero/", "463", "87", "2,1 %"],
    ["/artista/ana-mercedes-hoyos/", "452", "19", "2,1 %"],
    ["/blog/omar-rayo-una-historia-narrada-en-geometria/", "435", "31", "2,0 %"],
]
for r, row in enumerate(top10_pages_co):
    ws1.write(62 + r, 0, row[0], fmt_data)
    ws1.write(62 + r, 1, row[1], fmt_number)
    ws1.write(62 + r, 2, row[2], fmt_data_center)
    ws1.write(62 + r, 3, row[3], fmt_data_center)

# ── Top 10 Páginas US ──────────────────────────────────────────────────────
ws1.merge_range("A74:D74", "TOP 10 PÁGINAS POR TRÁFICO ORGÁNICO — ESTADOS UNIDOS (US)", fmt_section)
ws1.write_row("A75", ["URL", "Tráfico Estimado", "# Keywords", "% del Tráfico Total"], fmt_header)

top10_pages_us = [
    ["/artista/ana-mercedes-hoyos/", "561", "2", "25,2 %"],
    ["/en/blog/all-you-need-to-know-to-understand-fernando-boteros-artworks/", "234", "99", "10,5 %"],
    ["/en/blog/unmistakable-brand-boterism/", "218", "74", "9,8 %"],
    ["/en/blog/discovering-colombian-art-from-fernando-botero-to-contemporary-masters/", "169", "77", "7,6 %"],
    ["/en/blog/artistic-periods-of-pablo-picasso/", "100", "151", "4,5 %"],
    ["/en/blog/the-works-of-oswaldo-guayasamin/", "99", "39", "4,4 %"],
    ["/en/artist/luis-caballero/", "88", "13", "3,9 %"],
    ["/en/blog/the-works-of-fernando-botero-and-their-significance/", "79", "68", "3,5 %"],
    ["/blog/las-obras-de-fernando-botero-y-su-significado/", "61", "33", "2,7 %"],
    ["/blog/las-obras-mejor-vendidas-de-fernando-botero/", "58", "29", "2,6 %"],
]
for r, row in enumerate(top10_pages_us):
    ws1.write(75 + r, 0, row[0], fmt_data)
    ws1.write(75 + r, 1, row[1], fmt_number)
    ws1.write(75 + r, 2, row[2], fmt_data_center)
    ws1.write(75 + r, 3, row[3], fmt_data_center)

ws1.write(86, 0, "Nota: Datos Semrush actualizados a junio 2026. Tráfico es estimación mensual.", fmt_note)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2: BENCHMARK COMPETITIVO
# ══════════════════════════════════════════════════════════════════════════════
ws2 = wb.add_worksheet("Benchmark Competitivo")
ws2.set_tab_color("#0F3460")
ws2.set_zoom(85)

col_widths_t2 = [32, 8, 16, 14, 14, 16, 14, 14]
for i, w in enumerate(col_widths_t2):
    ws2.set_column(i, i, w)
ws2.set_row(0, 28)

ws2.merge_range("A1:H1",
    "BENCHMARK COMPETITIVO — 17 Galerías | Colombia (CO) y Estados Unidos (US) | junio 2026",
    fmt_title)

headers2 = [
    "Dominio", "Tier",
    "Tráfico CO", "Keywords CO",
    "Tráfico US", "Keywords US",
    "Rango Semrush CO", "Rango Semrush US"
]
ws2.write_row("A2", headers2, fmt_header)

# data sorted by Tráfico CO descending; Duque Arango highlighted
benchmark_data = [
    # dominio, tier, trafico_co, kw_co, trafico_us, kw_us, rank_co, rank_us, is_duque
    ("galeriaduquearango.com", "—",    21691, 2255, 2228, 1895, "4.908",    "605.349",   True),
    ("galeriaelmuseo.com",     "T1",   4017,  531,  20,   42,   "17.529",   "6.685.731", False),
    ("galerialacometa.com",    "T1",   3253,  360,  521,  139,  "20.478",   "1.637.554", False),
    ("casasriegner.com",       "T1",   1742,  232,  28,   55,   "32.237",   "6.035.803", False),
    ("sgr-art.com",            "T1",   674,   64,   0,    3,    "60.926",   "24.394.206",False),
    ("galeriacasacuadrada.com","T1",   512,   77,   0,    2,    "72.157",   "27.346.515",False),
    ("beatrizesguerra-art.com","T1",   169,   39,   95,   91,   "136.871",  "3.849.981", False),
    ("galeriafreites.com",     "T2",   164,   106,  13,   19,   "139.269",  "7.549.190", False),
    ("otros360grados.com",     "T1",   123,   55,   0,    9,    "161.834",  "17.396.295",False),
    ("artoftheworldgallery.com","T2",  13,    23,   1824, 754,  "461.218",  "705.085",   False),
    ("galeriaelsapineres.art", "T1",   15,    23,   0,    0,    "426.950",  "Sin datos", False),
    ("operagallery.com",       "T2",   56,    46,   8271, 3399, "237.458",  "203.500",   False),
    ("latinartcore.com",       "T2",   0,     0,    146,  116,  "Sin datos","3.189.621", False),
    ("ascasogallery.com",      "T2",   0,     12,   319,  157,  "1.308.428","2.162.533", False),
    ("cernudaarte.com",        "T2",   0,     6,    575,  408,  "1.514.672","1.544.097", False),
    ("miguelabreugallery.com", "T2",   0,     2,    3324, 439,  "2.788.593","440.776",   False),
    ("forumgallery.com",       "T2",   0,     5,    1938, 1545, "1.628.482","673.292",   False),
]

for r, row in enumerate(benchmark_data):
    dom, tier, tco, kco, tus, kus, rco, rus, is_duque = row
    if is_duque:
        ws2.write(2 + r, 0, dom, fmt_duque)
        ws2.write(2 + r, 1, tier, fmt_duque)
        ws2.write(2 + r, 2, tco, fmt_duque_num)
        ws2.write(2 + r, 3, kco, fmt_duque_num)
        ws2.write(2 + r, 4, tus, fmt_duque_num)
        ws2.write(2 + r, 5, kus, fmt_duque_num)
        ws2.write(2 + r, 6, rco, fmt_duque)
        ws2.write(2 + r, 7, rus, fmt_duque)
    else:
        ws2.write(2 + r, 0, dom, fmt_data)
        ws2.write(2 + r, 1, tier, fmt_data_center)
        ws2.write(2 + r, 2, tco, fmt_number)
        ws2.write(2 + r, 3, kco, fmt_number)
        ws2.write(2 + r, 4, tus, fmt_number)
        ws2.write(2 + r, 5, kus, fmt_number)
        ws2.write(2 + r, 6, rco, fmt_data_center)
        ws2.write(2 + r, 7, rus, fmt_data_center)

ws2.write(20, 0, "Tier T1 = Competidores directos Colombia | T2 = Mercado latinoamericano (Miami/NY)", fmt_note)
ws2.write(21, 0, "Orden: Tráfico CO descendente. galeriaduquearango.com resaltado en verde.", fmt_note)
ws2.write(22, 0, "Nota: latinartcore.com sin datos en Semrush CO. galeriaelsapineres.art sin datos en Semrush US.", fmt_note)
ws2.write(23, 0, "Authority Score: no disponible en reporte domain_rank; requiere pull separado de backlinks.", fmt_note)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3: BRECHAS DE KEYWORDS
# ══════════════════════════════════════════════════════════════════════════════
ws3 = wb.add_worksheet("Brechas de Keywords")
ws3.set_tab_color("#E94560")
ws3.set_zoom(85)

col_widths_t3 = [28, 14, 18, 26, 18, 14, 38]
for i, w in enumerate(col_widths_t3):
    ws3.set_column(i, i, w)
ws3.set_row(0, 28)

ws3.merge_range("A1:G1",
    "BRECHAS DE KEYWORDS — Competidores que rankean donde Duque Arango no | junio 2026",
    fmt_title)

headers3 = [
    "Keyword", "Volumen de Búsqueda",
    "Categoría SEO", "Competidor que Ranquea",
    "Posición Competidor", "Nivel de Oportunidad",
    "Acción de Contenido Sugerida"
]
ws3.write_row("A2", headers3, fmt_header)

# Categorías: Nombre de Artista / Marca / Editorial / Comercial
# Nivel: Alta (>5.000) / Media (1.000-5.000) / Baja (<1.000)
brechas = [
    # keyword, vol, cat, competidor, pos_comp, nivel, accion
    ("wifredo lam",
     "6.600", "Nombre de Artista",
     "cernudaarte.com (US)", "54",
     "Alta",
     "⚠️ ARTISTA EN ROSTER: Crear /en/artist/wifredo-lam/ con E-E-A-T, FAQ schema y obras principales"),
    ("museo de arte",
     "8.100", "Editorial",
     "galeriaelmuseo.com (CO)", "15",
     "Alta",
     "Crear guía editorial 'Museos y galerías de arte en Colombia 2026' para capturar tráfico informacional"),
    ("beatriz gonzalez",
     "3.600", "Nombre de Artista",
     "casasriegner.com (CO)", "4",
     "Media",
     "Crear artículo editorial sobre Beatriz González como referente del arte conceptual colombiano"),
    ("galerias bogotá",
     "2.900", "Comercial",
     "galeriaelmuseo.com (CO)", "11",
     "Media",
     "Crear landing page o guía 'Galerías de arte en Bogotá' para capturar búsquedas de intención comercial"),
    ("description of art gallery",
     "2.900", "Editorial",
     "beatrizesguerra-art.com (US)", "34",
     "Media",
     "Crear página /en/what-is-an-art-gallery/ optimizada para búsquedas de definición en inglés"),
    ("autorretrato",
     "2.400", "Editorial",
     "galeriaelmuseo.com (CO)", "5",
     "Media",
     "Crear artículo 'El autorretrato en el arte latinoamericano' aprovechando obras del roster"),
    ("jesus abad colorado",
     "2.400", "Nombre de Artista",
     "galeriaelmuseo.com (CO)", "4",
     "Media",
     "Artículo editorial sobre fotografía colombiana citando artistas del entorno (no represantado)"),
    ("emma reyes",
     "1.900", "Nombre de Artista",
     "galerialacometa.com (CO)", "4",
     "Media",
     "Artículo editorial 'Las grandes olvidadas del arte colombiano' o serie de contenido cultural"),
    ("galerias bogota",
     "1.900", "Comercial",
     "galeriaelmuseo.com (CO)", "5",
     "Media",
     "Variante sin tilde: incluir en la misma landing page de Galerías Bogotá (mismo contenido)"),
    ("define gallery",
     "1.600", "Editorial",
     "beatrizesguerra-art.com (US)", "13",
     "Media",
     "Incluir bloque 'What is an art gallery' extractable en la página About/en en inglés"),
    ("amelia pelaez",
     "1.000", "Nombre de Artista",
     "cernudaarte.com (US)", "6",
     "Media",
     "Artículo editorial sobre artistas cubanas del siglo XX con mencion de Wifredo Lam (roster)"),
    ("carlos rojas",
     "1.000", "Nombre de Artista",
     "casasriegner.com (CO)", "2",
     "Media",
     "Artículo editorial sobre abstracción geométrica colombiana (relacionar con Omar Rayo, roster)"),
    ("bienal de arte medellin",
     "880", "Editorial",
     "casasriegner.com (CO)", "16",
     "Baja",
     "Artículo 'Galería Duque Arango en la Bienal de Medellín' — vincula galería a evento relevante"),
    ("jacanamijoy",
     "880", "Nombre de Artista",
     "galeriaelmuseo.com (CO)", "7",
     "Baja",
     "Mencionar en artículo sobre arte colombiano contemporáneo con contexto cultural"),
    ("galerías",
     "880", "Comercial",
     "galeriaelmuseo.com (CO)", "8",
     "Baja",
     "Optimizar la homepage para incluir la palabra 'galerías' de forma natural en el copy"),
    ("wifredo lam paintings",
     "590", "Nombre de Artista",
     "cernudaarte.com (US)", "37",
     "Baja",
     "⚠️ ARTISTA EN ROSTER: Incluir lista de obras en /en/artist/wifredo-lam/ con nombres en inglés"),
    ("tomas sanchez artist",
     "480", "Nombre de Artista",
     "cernudaarte.com (US)", "10",
     "Baja",
     "Artículo editorial sobre realismo mágico en el arte latinoamericano con artistas relacionados"),
    ("antonio caro",
     "720", "Nombre de Artista",
     "casasriegner.com (CO)", "3",
     "Baja",
     "Artículo sobre arte conceptual y político en Colombia (contexto cultural, no representación)"),
    ("casa galeria",
     "720", "Comercial",
     "casasriegner.com (CO)", "23",
     "Baja",
     "Incluir variante 'casa galería' en páginas de sedes o en página de contacto"),
    ("pedro ruiz",
     "880", "Nombre de Artista",
     "beatrizesguerra-art.com (CO)", "12",
     "Baja",
     "Mencionar en contexto editorial de retratos colombianos contemporáneos"),
]

for r, row in enumerate(brechas):
    kw, vol, cat, comp, pos, nivel, accion = row
    ws3.write(2 + r, 0, kw, fmt_data)
    ws3.write(2 + r, 1, vol, fmt_number)
    ws3.write(2 + r, 2, cat, fmt_data_center)
    ws3.write(2 + r, 3, comp, fmt_data)
    ws3.write(2 + r, 4, pos, fmt_data_center)
    if nivel == "Alta":
        ws3.write(2 + r, 5, nivel, fmt_alta)
    elif nivel == "Media":
        ws3.write(2 + r, 5, nivel, fmt_media)
    else:
        ws3.write(2 + r, 5, nivel, fmt_baja)
    ws3.write(2 + r, 6, accion, fmt_data)

ws3.write(23, 0, "⚠️ Filas marcadas con ⚠️ corresponden a artistas del roster de Galería Duque Arango — AMENAZA DE ALTA PRIORIDAD.", fmt_note)
ws3.write(24, 0, "Fuente: domain_domains Semrush (CO y US). Volúmenes según base Semrush junio 2026.", fmt_note)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4: PÁGINAS TOP DE COMPETENCIA
# ══════════════════════════════════════════════════════════════════════════════
ws4 = wb.add_worksheet("Páginas Top Competencia")
ws4.set_tab_color("#533483")
ws4.set_zoom(85)

col_widths_t4 = [26, 10, 48, 14, 28, 20, 40]
for i, w in enumerate(col_widths_t4):
    ws4.set_column(i, i, w)
ws4.set_row(0, 28)

ws4.merge_range("A1:G1",
    "PÁGINAS TOP DE COMPETENCIA — Top 3 CO + Top 3 US | junio 2026",
    fmt_title)

headers4 = [
    "Competidor", "Mercado",
    "URL de Página", "Tráfico Estimado",
    "Keyword Principal (inferida)", "Tipo de Contenido",
    "Brecha para Duque Arango"
]
ws4.write_row("A2", headers4, fmt_header)

top_pages = [
    # CO Competitors
    ("galeriaelmuseo.com", "CO",
     "galeriaelmuseo.com/", "1.476",
     "galeria el museo", "Homepage / Marca",
     "Duque Arango supera en tráfico CO; mantener diferenciación editorial"),
    ("galeriaelmuseo.com", "CO",
     "galeriaelmuseo.com/archives/2084/", "503",
     "Exposición artista (slug sin descriptor)", "Página de exposición",
     "Crear páginas de exposición optimizadas con keywords de artistas del roster"),
    ("galeriaelmuseo.com", "CO",
     "galeriaelmuseo.com/archives/43710/", "455",
     "Artículo editorial (slug sin descriptor)", "Blog / Editorial",
     "Analizar el slug para identificar el tema y crear contenido competitivo"),
    ("galeriaelmuseo.com", "CO",
     "galeriaelmuseo.com/archives/450/", "303",
     "Contenido de artista (slug sin descriptor)", "Página artista",
     "Expandir páginas de artistas individuales con más keywords de cola larga"),
    ("galeriaelmuseo.com", "CO",
     "galeriaelmuseo.com/archives/26977/", "196",
     "Contenido editorial (slug sin descriptor)", "Blog",
     "Identificar temas sin cubrir en el blog de Duque Arango"),
    ("galerialacometa.com", "CO",
     "galerialacometa.com/", "1.860",
     "galeria la cometa", "Homepage / Marca",
     "Monitorear: La Cometa es el segundo competidor CO más grande"),
    ("galerialacometa.com", "CO",
     "galerialacometa.com/artistas/miguel-angel-rojas-es", "86",
     "miguel angel rojas", "Página de artista",
     "Artistas que galerialacometa.com representa podrían informar estrategia editorial de Duque"),
    ("galerialacometa.com", "CO",
     "galerialacometa.com/exhibiciones/bogota/emma-reyes-las-caras-de-emma-reyes-es", "83",
     "emma reyes", "Página de exposición",
     "Emma Reyes genera 1.900 búsquedas/mes CO — brecha editorial de contenido para Duque Arango"),
    ("casasriegner.com", "CO",
     "casasriegner.com/", "994",
     "casas riegner", "Homepage / Marca",
     "Casas & Riegner es competidor fuerte en Bogotá; monitorear posiciones compartidas"),
    ("casasriegner.com", "CO",
     "casasriegner.com/artistas/beatriz-gonzalez", "179",
     "beatriz gonzalez", "Página de artista",
     "beatriz gonzalez tiene 3.600 búsquedas/mes CO — Duque Arango ausente en este keyword"),
    ("casasriegner.com", "CO",
     "casasriegner.com/artistas/carlos-rojas", "116",
     "carlos rojas", "Página de artista",
     "Crear contenido editorial sobre abstraccón geométrica (Carlos Rojas ↔ Omar Rayo en roster)"),
    ("casasriegner.com", "CO",
     "casasriegner.com/artistas/antonio-caro", "66",
     "antonio caro", "Página de artista",
     "Arte conceptual colombiano — oportunidad editorial sin necesitar representar al artista"),
    # US Competitors
    ("operagallery.com", "US",
     "operagallery.com/", "1.384",
     "opera gallery", "Homepage / Marca",
     "Opera Gallery domina US con 8.271 visitas. Su fortaleza editorial en inglés es la referencia"),
    ("operagallery.com", "US",
     "operagallery.com/artist/keith-haring", "848",
     "keith haring", "Página de artista",
     "Opera usa páginas de artista profundas en inglés — modelo a replicar para artistas del roster"),
    ("operagallery.com", "US",
     "operagallery.com/artist/bernard-buffet", "611",
     "bernard buffet", "Página de artista",
     "Páginas de artistas modernos internacionales generan alto tráfico US"),
    ("operagallery.com", "US",
     "operagallery.com/viewing-rooms/botero-2023", "476",
     "botero 2023 opera gallery", "Viewing Room / Exposición virtual",
     "⚠️ Opera Gallery captura tráfico de Fernando Botero en US. Crear /en/artist/fernando-botero/"),
    ("operagallery.com", "US",
     "operagallery.com/viewing-rooms/fernando-botero", "452",
     "fernando botero gallery", "Viewing Room / Exposición virtual",
     "⚠️ Segunda página de Botero en Opera. Duque Arango debe crear página de artista dedicada en inglés"),
    ("miguelabreugallery.com", "US",
     "miguelabreugallery.com/", "2.402",
     "miguel abreu gallery", "Homepage / Marca",
     "72% del tráfico US en la homepage — Duque Arango tiene tráfico más distribuido (mejor)"),
    ("miguelabreugallery.com", "US",
     "miguelabreugallery.com/artists/quaytman/", "158",
     "r.h. quaytman artist", "Página de artista",
     "Páginas de artistas en inglés con URLs descriptivas generan tráfico consistente"),
    ("miguelabreugallery.com", "US",
     "miguelabreugallery.com/artists/francois-marie-banier/", "142",
     "francois marie banier", "Página de artista",
     "Artistas europeos representados generan demanda en US — modelo aplicable al roster de Duque"),
    ("forumgallery.com", "US",
     "forumgallery.com/", "865",
     "forum gallery nyc", "Homepage / Marca",
     "Forum Gallery tiene 1.545 keywords US vs 1.895 de Duque Arango — gap técnico moderado"),
    ("forumgallery.com", "US",
     "forumgallery.com/artists/claudio-bravo/videos", "114",
     "claudio bravo artist", "Página de artista con videos",
     "Contenido multimedia (videos) en páginas de artistas impulsa tráfico — oportunidad para Duque"),
    ("forumgallery.com", "US",
     "forumgallery.com/artists/norman-rockwell/biography", "100",
     "norman rockwell biography", "Biografía de artista",
     "Páginas de biografía específicas generan tráfico; aplicar a artistas modernos en inglés"),
    ("forumgallery.com", "US",
     "forumgallery.com/artists/gregory-gillespie", "92",
     "gregory gillespie artist", "Página de artista",
     "Artistas menos conocidos pueden generar tráfico de cola larga con páginas bien optimizadas"),
]

for r, row in enumerate(top_pages):
    comp, mkt, url, traf, kw, tipo, brecha = row
    ws4.write(2 + r, 0, comp, fmt_data)
    ws4.write(2 + r, 1, mkt, fmt_data_center)
    ws4.write(2 + r, 2, url, fmt_data)
    ws4.write(2 + r, 3, traf, fmt_number)
    ws4.write(2 + r, 4, kw, fmt_data)
    ws4.write(2 + r, 5, tipo, fmt_data_center)
    ws4.write(2 + r, 6, brecha, fmt_data)

ws4.write(27, 0, "⚠️ = amenaza directa sobre artista del roster de Galería Duque Arango", fmt_note)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 5: INSIGHTS Y RECOMENDACIONES
# ══════════════════════════════════════════════════════════════════════════════
ws5 = wb.add_worksheet("Insights y Recomendaciones")
ws5.set_tab_color("#F5C518")
ws5.set_zoom(85)

col_widths_t5 = [32, 38, 28, 42, 10]
for i, w in enumerate(col_widths_t5):
    ws5.set_column(i, i, w)
ws5.set_row(0, 28)

ws5.merge_range("A1:E1",
    "INSIGHTS Y RECOMENDACIONES ESTRATÉGICAS — Semana del 2026-06-05",
    fmt_title)

headers5 = [
    "Insight", "Evidencia",
    "Framework Aplicado",
    "Acción Recomendada",
    "Prioridad"
]
ws5.write_row("A2", headers5, fmt_header)

insights = [
    (
        "1. Liderazgo dominante en Colombia: Duque Arango es 5.4x más grande que el siguiente competidor",
        "21.691 visitas orgánicas CO vs galeriaelmuseo.com en 2° lugar con 4.017. "
        "Cobertura de 2.255 keywords orgánicas, ranking nacional #4.908.",
        "seo-audit: Content Quality — E-E-A-T, Authoritativeness.\n"
        "Posición de autoridad temática consolidada en arte colombiano.",
        "Mantener cadencia editorial mínima 2 posts/semana sobre artistas del roster. "
        "Identificar keywords de posición 3-10 para optimización (quick wins).",
        "Alta"
    ),
    (
        "2. ⚠️ Wifredo Lam (EN ROSTER) sin posicionamiento en US — cernudaarte.com captura el tráfico",
        "cernudaarte.com ranquea para 'wifredo lam' (6.600 búsquedas/mes US). "
        "Duque Arango representa al artista pero no aparece en Semrush US para este término. "
        "También: 'wifredo lam paintings' (590 búsquedas/mes US) sin presencia.",
        "competitor-alternatives: Research Process — artista del roster que competidor captura.\n"
        "seo-audit: On-Page SEO — Keyword targeting, título y H1 alineados.",
        "Crear y optimizar /en/artist/wifredo-lam/ con: texto biográfico E-E-A-T en inglés, "
        "lista de obras con nombres en inglés, FAQ schema, estadísticas de mercado y "
        "sección 'Last updated'. Meta title: 'Wifredo Lam | Works & Biography | Galería Duque Arango'.",
        "Alta"
    ),
    (
        "3. Gap comercial crítico: 'galerías bogotá' (2.900/mes CO) — Duque Arango ausente",
        "3 competidores rankean: galeriaelmuseo pos 11, casasriegner pos 19, galerialacometa pos 30. "
        "Duque Arango tiene sede en Bogotá pero no aparece. También: 'galerias bogota' sin tilde (1.900/mes).",
        "content-strategy: Keyword Research by Buyer Stage — Commercial Intent (Decision Stage).\n"
        "Captura de búsquedas de intención comercial directa.",
        "Crear landing page '/galeria-arte-bogota/' con: ubicación, artistas representados en Bogotá, "
        "exposiciones activas, y optimizada para 'galerías bogotá' + 'galerias bogota'. "
        "Enlazar desde el menú principal y desde páginas de artistas.",
        "Alta"
    ),
    (
        "4. Concentración de riesgo: Edgar Negret genera el 40.5% del tráfico orgánico CO",
        "/artista/edgar-negret/ genera 8.782 de 21.691 visitas CO (40.5%). "
        "Una sola actualización de Google podría reducir drásticamente el tráfico total.",
        "seo-audit: Content Quality — Thin Content risk, topical diversification.\n"
        "Riesgo de sobreexposición a un solo activo de contenido.",
        "Desarrollar páginas de artista con el mismo nivel de profundidad para: "
        "David Manzur, Alejandro Obregón, Enrique Grau, Ana Mercedes Hoyos y Omar Rayo. "
        "Cada página debe superar las 1.500 palabras con FAQ, timeline y lista de obras.",
        "Alta"
    ),
    (
        "5. ⚠️ operagallery.com captura tráfico de Fernando Botero en US (928 visitas/mes desde 2 páginas)",
        "Opera Gallery tiene /viewing-rooms/botero-2023 (476 visitas US) y "
        "/viewing-rooms/fernando-botero (452 visitas US). "
        "Duque Arango tiene contenido editorial de Botero pero en posiciones 8-20 y sin página de artista dedicada en inglés.",
        "ai-seo: Pillar 1 — Content Extractability. Pillar 2 — Authority (citar fuentes, estadísticas).\n"
        "competitor-alternatives: Competitor vs Competitor page format.",
        "Crear /en/artist/fernando-botero/ con: biografía E-E-A-T, lista de obras más vendidas, "
        "precios de mercado secundario, FAQ schema ('What is Fernando Botero known for?', "
        "'What is Botero's style called?'), y 'Last updated: junio 2026'. "
        "Enlazar desde todos los posts editoriales de Botero.",
        "Alta"
    ),
    (
        "6. Contenido en inglés subóptimo vs competidores US: solo 2.228 visitas vs 8.271 de operagallery",
        "Top páginas US son blogs editoriales en posiciones 6-24. "
        "Ana Mercedes Hoyos /artista/ (español) genera más tráfico US que páginas en inglés. "
        "Páginas en inglés carecen de 'Last updated', author bio visible y FAQ schema.",
        "ai-seo: Content Extractability Check — freshness, structured data, expert attribution.\n"
        "Prioridad: secciones extractables para AI Overviews en Google US.",
        "Auditar las 10 páginas en inglés con más tráfico US y aplicar: "
        "'Last updated [fecha]' visible, author bio con credenciales, "
        "sección FAQ con 3-5 preguntas frecuentes, y estadísticas citadas con fuentes. "
        "Objetivo: aumentar visibilidad en Google AI Overviews para búsquedas de arte latinoamericano en inglés.",
        "Media"
    ),
    (
        "7. galeriaelmuseo.com captura 'museo de arte' (8.100 búsquedas/mes CO) — nadie domina este keyword",
        "galeriaelmuseo.com pos 15 para 'museo de arte'. "
        "Keyword con alto volumen y baja competencia entre galerías. "
        "Ningún competidor domina con posición top-5.",
        "content-strategy: Content Pillars — Arte colombiano como pilar editorial.\n"
        "Hub and Spoke: 'Arte en Colombia' como hub con artículos de museos, galerías y artistas.",
        "Crear artículo definitivo 'Museos y galerías de arte en Colombia: guía completa 2026' "
        "(mínimo 2.000 palabras). Incluir lista de museos, galerías representativas, artistas de referencia "
        "y enlazar a páginas de artistas del roster. Actualizar anualmente.",
        "Media"
    ),
    (
        "8. Casas & Riegner captura 3 keywords de artistas históricos colombianos relevantes (no en roster)",
        "casasriegner.com pos 3-4 para: 'beatriz gonzalez' (3.600 CO), 'carlos rojas' (1.000 CO), "
        "'antonio caro' (720 CO). Estos artistas son referencias culturales clave del mercado target.",
        "competitor-alternatives: Research Process — identificar demanda orgánica de artistas clave.\n"
        "content-strategy: Thought Leadership, Data-Driven Content.",
        "Crear serie editorial 'Arte colombiano: los grandes maestros' con artículos sobre artistas "
        "colombianos clave (Beatriz González, Carlos Rojas) como autoridad cultural editorial, "
        "sin competir con su representación. Posiciona a Duque Arango como referencia educativa en arte CO.",
        "Media"
    ),
]

for r, row in enumerate(insights):
    insight, evidencia, framework, accion, prioridad = row
    ws5.write(2 + r, 0, insight, fmt_data)
    ws5.write(2 + r, 1, evidencia, fmt_data)
    ws5.write(2 + r, 2, framework, fmt_data)
    ws5.write(2 + r, 3, accion, fmt_data)
    if prioridad == "Alta":
        ws5.write(2 + r, 4, prioridad, fmt_alta)
    elif prioridad == "Media":
        ws5.write(2 + r, 4, prioridad, fmt_media)
    else:
        ws5.write(2 + r, 4, prioridad, fmt_baja)

ws5.set_row(2, 120)
ws5.set_row(3, 120)
ws5.set_row(4, 120)
ws5.set_row(5, 120)
ws5.set_row(6, 120)
ws5.set_row(7, 80)
ws5.set_row(8, 80)
ws5.set_row(9, 80)

ws5.write(11, 0,
    "Frameworks aplicados: seo-audit v1.2, ai-seo v1.2, content-strategy v1.1, competitor-alternatives v1.1 "
    "| Rutina SEO Semanal — Galería Duque Arango | Jueves 5 junio 2026",
    fmt_note)

wb.close()

# Verificar tamaño
size = os.path.getsize(FILEPATH)
print(f"✅ Archivo generado: {FILENAME}")
print(f"   Tamaño: {size:,} bytes ({size/1024:.1f} KB)")
print(f"   Ruta: {FILEPATH}")
