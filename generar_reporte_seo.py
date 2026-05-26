#!/usr/bin/env python3
"""Genera el reporte semanal de SEO para Galería Duque Arango."""

import xlsxwriter
import os

FECHA = "2026-05-26"
FILENAME = f"Inteligencia_SEO_{FECHA}.xlsx"
FILEPATH = os.path.join("/home/user/da-routines", FILENAME)

wb = xlsxwriter.Workbook(FILEPATH)

# ── Formatos ──────────────────────────────────────────────────────────────────
fmt_title = wb.add_format({
    "bold": True, "font_size": 14, "font_color": "#1F2D3D",
    "bg_color": "#F0F0F0", "border": 1,
})
fmt_header = wb.add_format({
    "bold": True, "font_color": "#FFFFFF", "bg_color": "#2C3E50",
    "border": 1, "text_wrap": True, "align": "center", "valign": "vcenter",
})
fmt_header_alt = wb.add_format({
    "bold": True, "font_color": "#FFFFFF", "bg_color": "#1A6B4A",
    "border": 1, "text_wrap": True, "align": "center", "valign": "vcenter",
})
fmt_cell = wb.add_format({"border": 1, "text_wrap": True, "valign": "top"})
fmt_cell_c = wb.add_format({"border": 1, "align": "center", "valign": "top"})
fmt_num = wb.add_format({"border": 1, "num_format": "#,##0", "align": "center"})
fmt_highlight = wb.add_format({
    "bold": True, "bg_color": "#FFF2CC", "border": 1,
    "num_format": "#,##0", "align": "center",
})
fmt_highlight_text = wb.add_format({
    "bold": True, "bg_color": "#FFF2CC", "border": 1, "text_wrap": True,
})
fmt_alta = wb.add_format({
    "bold": True, "font_color": "#FFFFFF", "bg_color": "#C0392B",
    "border": 1, "align": "center",
})
fmt_media = wb.add_format({
    "bold": True, "font_color": "#FFFFFF", "bg_color": "#E67E22",
    "border": 1, "align": "center",
})
fmt_baja = wb.add_format({
    "bold": True, "font_color": "#000000", "bg_color": "#F9E79F",
    "border": 1, "align": "center",
})
fmt_subhead = wb.add_format({
    "bold": True, "font_color": "#FFFFFF", "bg_color": "#5D6D7E",
    "border": 1, "text_wrap": True, "align": "center",
})
fmt_meta = wb.add_format({
    "italic": True, "font_color": "#7F8C8D", "border": 0,
})


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1: Desempeño Propio
# ══════════════════════════════════════════════════════════════════════════════
ws1 = wb.add_worksheet("Desempeño Propio")
ws1.set_column("A:A", 20)
ws1.set_column("B:B", 28)
ws1.set_column("C:C", 18)
ws1.set_column("D:D", 45)

ws1.merge_range("A1:D1",
    "DESEMPEÑO PROPIO — galeriaduquearango.com | Datos al mayo 2026",
    fmt_title)
ws1.write("A2", "Generado:", fmt_meta)
ws1.write("B2", f"{FECHA} — Rutina de Inteligencia SEO", fmt_meta)

# Sub-tabla 1: Resumen de dominio
ws1.merge_range("A4:D4", "RESUMEN DEL DOMINIO", fmt_subhead)
headers_dom = ["Base de Datos", "Métrica", "Valor", "Notas"]
for c, h in enumerate(headers_dom):
    ws1.write(4, c, h, fmt_header)

domain_data = [
    ("CO (Colombia)", "Keywords orgánicos",       "2.095",    "Posición Semrush: 4.648"),
    ("CO (Colombia)", "Tráfico orgánico estimado","23.273",   "Visitas/mes — liderazgo claro en el mercado CO"),
    ("CO (Colombia)", "Costo orgánico estimado",  "USD 2.201","Valor del tráfico si fuera pago"),
    ("CO (Colombia)", "Keywords top 10",           "13",       "Basado en top 20 keywords exportados"),
    ("US (Estados Unidos)", "Keywords orgánicos",  "1.908",   "Posición Semrush: 596.659"),
    ("US (Estados Unidos)", "Tráfico orgánico est.","2.296",  "3er lugar entre galerías competidoras en US"),
    ("US (Estados Unidos)", "Costo orgánico est.", "USD 383", "Potencial de mejora significativo"),
    ("US (Estados Unidos)", "Keywords top 10",      "6",       "Basado en top 20 keywords exportados"),
]
for r, row in enumerate(domain_data):
    for c, val in enumerate(row):
        fmt = fmt_highlight_text if row[0].startswith("CO") and c == 2 else fmt_cell
        ws1.write(5 + r, c, val, fmt)

# Sub-tabla 2: Top 20 Keywords CO
ws1.merge_range("A15:D15", "TOP 20 KEYWORDS — Base CO (Colombia)", fmt_subhead)
ws1.set_column("A:D", None)
ws1.set_column("A:A", 38)
ws1.set_column("B:B", 12)
ws1.set_column("C:C", 14)
ws1.set_column("D:D", 65)

kw_headers = ["Keyword", "Posición", "Volumen Búsqueda", "URL que Rankea"]
for c, h in enumerate(kw_headers):
    ws1.write(15, c, h, fmt_header)

top_kw_co = [
    ("edgar negret",                        2,  90500, "galeriaduquearango.com/artista/edgar-negret/"),
    ("plaza botero medellín antioquia",     7,  22200, "galeriaduquearango.com/blog/plaza-botero-resignificar-el-espacio..."),
    ("galeria duque arango",                1,    720, "galeriaduquearango.com/"),
    ("obras de alejandro obregón",          1,    390, "galeriaduquearango.com/artista/alejandro-obregon/"),
    ("david manzur",                        2,   4400, "galeriaduquearango.com/artista/david-manzur/"),
    ("plaza botero",                        8,  49500, "galeriaduquearango.com/blog/plaza-botero-resignificar-el-espacio..."),
    ("enrique grau",                        2,   1600, "galeriaduquearango.com/artista/enrique-grau/"),
    ("ana mercedes hoyos obras",            1,    210, "galeriaduquearango.com/artista/ana-mercedes-hoyos/"),
    ("alejandro obregon obras",             1,    210, "galeriaduquearango.com/artista/alejandro-obregon/"),
    ("guayasamin",                          3,   2400, "galeriaduquearango.com/blog/el-legado-al-arte-de-oswaldo-guayasamin/"),
    ("duque arango galeria",                1,    110, "galeriaduquearango.com/"),
    ("obras de arte de david manzur",       1,    170, "galeriaduquearango.com/artista/david-manzur/"),
    ("botero cuadros",                      1,    170, "galeriaduquearango.com/blog/las-obras-mejor-vendidas-de-fernando-botero/"),
    ("omar rayo",                           6,   5400, "galeriaduquearango.com/blog/omar-rayo-una-historia-narrada-en-geometria/"),
    ("edgar negret obras",                  1,    140, "galeriaduquearango.com/artista/edgar-negret/"),
    ("javier caraballo",                    1,    140, "galeriaduquearango.com/artista/javier-caraballo/"),
    ("grau",                                2,   1300, "galeriaduquearango.com/artista/enrique-grau/"),
    ("muralismo",                           3,   1600, "galeriaduquearango.com/blog/muralismo-latinoamericano-el-arte-como-instrumento..."),
    ("edgar negret (blog)",                19,  90500, "galeriaduquearango.com/blog/edgar-negret-el-escultor-colombiano..."),
    ("obregón pinturas",                    2,    480, "galeriaduquearango.com/artista/alejandro-obregon/"),
]
for r, (kw, pos, vol, url) in enumerate(top_kw_co):
    ws1.write(16 + r, 0, kw, fmt_cell)
    ws1.write(16 + r, 1, pos, fmt_cell_c)
    ws1.write(16 + r, 2, vol, fmt_num)
    ws1.write(16 + r, 3, url, fmt_cell)

# Sub-tabla 3: Top 20 Keywords US
ws1.merge_range("A38:D38", "TOP 20 KEYWORDS — Base US (Estados Unidos)", fmt_subhead)
for c, h in enumerate(kw_headers):
    ws1.write(38, c, h, fmt_header_alt)

top_kw_us = [
    ("fernando botero",                     7,  14800, "galeriaduquearango.com/en/blog/unmistakable-brand-boterism/"),
    ("colombian painters",                  2,    720, "galeriaduquearango.com/en/blog/discovering-colombian-art..."),
    ("fernando botero art",                 5,   2400, "galeriaduquearango.com/en/blog/unmistakable-brand-boterism/"),
    ("botero's",                            6,   1600, "galeriaduquearango.com/en/blog/all-you-need-to-know-to-understand-fernando-boteros..."),
    ("ana mercedes hoyos",                  7,   3600, "galeriaduquearango.com/artista/ana-mercedes-hoyos/"),
    ("guayasamin",                          4,   1300, "galeriaduquearango.com/en/blog/the-works-of-oswaldo-guayasamin/"),
    ("plaza botero medellín colombia",      6,   1900, "galeriaduquearango.com/en/blog/plaza-plaza-botero-resignifying..."),
    ("oswaldo guayasamin",                  7,   1600, "galeriaduquearango.com/en/blog/oswaldo-guayasamins-legacy-to-art/"),
    ("botero",                              9,   9900, "galeriaduquearango.com/en/blog/all-you-need-to-know-to-understand-fernando-boteros..."),
    ("fernando botero artworks",            9,   1300, "galeriaduquearango.com/en/blog/all-you-need-to-know-to-understand-fernando-boteros..."),
    ("fernando botero style",               1,    170, "galeriaduquearango.com/en/blog/unmistakable-brand-boterism/"),
    ("plaza botero",                        8,   1000, "galeriaduquearango.com/en/blog/plaza-plaza-botero-resignifying..."),
    ("boterismo",                           3,    480, "galeriaduquearango.com/en/blog/unmistakable-brand-boterism/"),
    ("nombre de las 30 pinturas más vistas del fernando botero", 2, 210, "galeriaduquearango.com/blog/las-obras-mejor-vendidas-de-fernando-botero/"),
    ("picasso realistic art",               6,    480, "galeriaduquearango.com/en/blog/artistic-periods-of-pablo-picasso/"),
    ("picasso realism",                     7,    590, "galeriaduquearango.com/en/blog/artistic-periods-of-pablo-picasso/"),
    ("colombian painters famous",           2,    170, "galeriaduquearango.com/en/blog/discovering-colombian-art..."),
    ("fernando botero paintings",          11,   4400, "galeriaduquearango.com/en/blog/all-you-need-to-know-to-understand-fernando-boteros..."),
    ("guayasamin artist",                   4,    390, "galeriaduquearango.com/en/blog/the-works-of-oswaldo-guayasamin/"),
    ("mexican art",                        15,   8100, "galeriaduquearango.com/en/blog/mexican-art-history-evolution-and-two-key-figures..."),
]
for r, (kw, pos, vol, url) in enumerate(top_kw_us):
    ws1.write(39 + r, 0, kw, fmt_cell)
    ws1.write(39 + r, 1, pos, fmt_cell_c)
    ws1.write(39 + r, 2, vol, fmt_num)
    ws1.write(39 + r, 3, url, fmt_cell)

# Sub-tabla 4: Top 10 Páginas CO
ws1.merge_range("A61:D61", "TOP 10 PÁGINAS — Base CO (Colombia)", fmt_subhead)
page_headers = ["URL de la Página", "Keywords", "Tráfico Estimado", "Keyword Principal"]
for c, h in enumerate(page_headers):
    ws1.write(61, c, h, fmt_header)

top_pages_co = [
    ("galeriaduquearango.com/artista/edgar-negret/",              24,  7841, "edgar negret"),
    ("galeriaduquearango.com/blog/plaza-botero-resignificar...",  35,  4388, "plaza botero medellín antioquia"),
    ("galeriaduquearango.com/artista/david-manzur/",              51,  2700, "david manzur"),
    ("galeriaduquearango.com/blog/las-obras-de-fernando-botero-y-su-significado/", 145, 1466, "obras de botero"),
    ("galeriaduquearango.com/ (homepage)",                        96,   841, "galeria duque arango"),
    ("galeriaduquearango.com/artista/alejandro-obregon/",         67,   784, "obras de alejandro obregón"),
    ("galeriaduquearango.com/blog/las-obras-mejor-vendidas-de-fernando-botero/", 91, 783, "botero cuadros"),
    ("galeriaduquearango.com/blog/por-que-fernando-botero-es-tan-importante/", 44, 497, "por qué botero es importante"),
    ("galeriaduquearango.com/artista/ana-mercedes-hoyos/",        19,   452, "ana mercedes hoyos obras"),
    ("galeriaduquearango.com/blog/artistas-colombianos-que-debes-conocer/", 168, 345, "artistas colombianos"),
]
for r, (url, kws, tr, kw) in enumerate(top_pages_co):
    ws1.write(62 + r, 0, url, fmt_cell)
    ws1.write(62 + r, 1, kws, fmt_num)
    ws1.write(62 + r, 2, tr, fmt_num)
    ws1.write(62 + r, 3, kw, fmt_cell)

# Sub-tabla 5: Top 10 Páginas US
ws1.merge_range("A74:D74", "TOP 10 PÁGINAS — Base US (Estados Unidos)", fmt_subhead)
for c, h in enumerate(page_headers):
    ws1.write(74, c, h, fmt_header_alt)

top_pages_us = [
    ("galeriaduquearango.com/artista/ana-mercedes-hoyos/",           2,  521, "ana mercedes hoyos"),
    ("galeriaduquearango.com/en/blog/unmistakable-brand-boterism/", 77,  270, "fernando botero / boterismo"),
    ("galeriaduquearango.com/en/blog/all-you-need-to-know-to-understand-fernando-boteros-artworks/", 94, 204, "botero artworks"),
    ("galeriaduquearango.com/en/blog/the-works-of-oswaldo-guayasamin/", 40, 203, "guayasamin"),
    ("galeriaduquearango.com/en/blog/discovering-colombian-art-from-fernando-botero-to-contemporary-masters/", 72, 194, "colombian painters"),
    ("galeriaduquearango.com/en/artist/luis-caballero/",             14,   88, "luis caballero"),
    ("galeriaduquearango.com/en/blog/artistic-periods-of-pablo-picasso/", 146, 80, "picasso realism"),
    ("galeriaduquearango.com/blog/las-obras-de-fernando-botero-y-su-significado/", 33, 69, "obras botero significado"),
    ("galeriaduquearango.com/en/blog/plaza-plaza-botero-resignifying-public-space-through-art-in-medellin/", 7, 66, "plaza botero medellín colombia"),
    ("galeriaduquearango.com/blog/las-obras-mejor-vendidas-de-fernando-botero/", 29, 58, "nombre de las 30 pinturas más vistas del fernando botero"),
]
for r, (url, kws, tr, kw) in enumerate(top_pages_us):
    ws1.write(75 + r, 0, url, fmt_cell)
    ws1.write(75 + r, 1, kws, fmt_num)
    ws1.write(75 + r, 2, tr, fmt_num)
    ws1.write(75 + r, 3, kw, fmt_cell)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2: Benchmark Competitivo
# ══════════════════════════════════════════════════════════════════════════════
ws2 = wb.add_worksheet("Benchmark Competitivo")
ws2.set_column("A:A", 32)
ws2.set_column("B:B", 10)
ws2.set_column("C:H", 16)

ws2.merge_range("A1:H1",
    "BENCHMARK COMPETITIVO — 16 galerías vs. Duque Arango | Datos al mayo 2026",
    fmt_title)

bench_headers = [
    "Dominio", "Tier",
    "Tráfico CO", "Keywords CO", "Authority CO",
    "Tráfico US", "Keywords US", "Authority US",
]
for c, h in enumerate(bench_headers):
    ws2.write(2, c, h, fmt_header)

# Datos ordenados por Tráfico CO descendente
# galeriaduquearango.com primero (resaltado)
bench_data = [
    # dominio, tier, tr_co, kw_co, as_co, tr_us, kw_us, as_us, highlight
    ("galeriaduquearango.com",   "Referencia", 23273, 2095, "N/D",  2296, 1908, "N/D",  True),
    ("galeriaelmuseo.com",       "Tier 1",      4252,  509, "N/D",    20,   43, "N/D",  False),
    ("galerialacometa.com",      "Tier 1",      3169,  349, "N/D",   511,  146, "N/D",  False),
    ("casasriegner.com",         "Tier 1",      1656,  239, "N/D",    27,   58, "N/D",  False),
    ("sgr-art.com",              "Tier 1",       675,   63, "N/D",     0,    3, "N/D",  False),
    ("galeriacasacuadrada.com",  "Tier 1",       516,   64, "N/D",     0,    1, "N/D",  False),
    ("beatrizesguerra-art.com",  "Tier 1",       169,   41, "N/D",   134,   88, "N/D",  False),
    ("galeriafreites.com",       "Tier 2",       162,  105, "N/D",    13,   18, "N/D",  False),
    ("otros360grados.com",       "Tier 1",       105,   54, "N/D",     0,    6, "N/D",  False),
    ("operagallery.com",         "Tier 2",        54,   49, "N/D",  7263, 3482, "N/D",  False),
    ("artoftheworldgallery.com", "Tier 2",        12,   21, "N/D",  1445,  717, "N/D",  False),
    ("galeriaelsapineres.art",   "Tier 1",        15,   20, "N/D",     0,    0, "Sin datos CO/US", False),
    ("galeriaelsapineres.art — nota", "Tier 1",   0,    0, "N/D",     0,    0, "Sin datos en Semrush", False),
    ("ascasogallery.com",        "Tier 2",         0,   11, "N/D",   576,  158, "N/D",  False),
    ("cernudaarte.com",          "Tier 2",         0,    5, "N/D",   546,  416, "N/D",  False),
    ("forumgallery.com",         "Tier 2",         0,    5, "N/D",  1930, 1548, "N/D",  False),
    ("miguelabreugallery.com",   "Tier 2",         0,    3, "N/D",  3176,  461, "N/D",  False),
    ("latinartcore.com",         "Tier 2",         0,    0, "Sin datos", 145, 120, "N/D", False),
]

# Remove the duplicate galeriaelsapineres line
bench_clean = [
    ("galeriaduquearango.com",   "Referencia", 23273, 2095, "N/D",  2296, 1908, "N/D",  True),
    ("galeriaelmuseo.com",       "Tier 1",      4252,  509, "N/D",    20,   43, "N/D",  False),
    ("galerialacometa.com",      "Tier 1",      3169,  349, "N/D",   511,  146, "N/D",  False),
    ("casasriegner.com",         "Tier 1",      1656,  239, "N/D",    27,   58, "N/D",  False),
    ("sgr-art.com",              "Tier 1",       675,   63, "N/D",     0,    3, "N/D",  False),
    ("galeriacasacuadrada.com",  "Tier 1",       516,   64, "N/D",     0,    1, "N/D",  False),
    ("beatrizesguerra-art.com",  "Tier 1",       169,   41, "N/D",   134,   88, "N/D",  False),
    ("galeriafreites.com",       "Tier 2",       162,  105, "N/D",    13,   18, "N/D",  False),
    ("otros360grados.com",       "Tier 1",       105,   54, "N/D",     0,    6, "N/D",  False),
    ("operagallery.com",         "Tier 2",        54,   49, "N/D",  7263, 3482, "N/D",  False),
    ("artoftheworldgallery.com", "Tier 2",        12,   21, "N/D",  1445,  717, "N/D",  False),
    ("galeriaelsapineres.art",   "Tier 1",        15,   20, "N/D", 0, 0, "Sin datos US", False),
    ("ascasogallery.com",        "Tier 2",         0,   11, "N/D",   576,  158, "N/D",  False),
    ("cernudaarte.com",          "Tier 2",         0,    5, "N/D",   546,  416, "N/D",  False),
    ("forumgallery.com",         "Tier 2",         0,    5, "N/D",  1930, 1548, "N/D",  False),
    ("miguelabreugallery.com",   "Tier 2",         0,    3, "N/D",  3176,  461, "N/D",  False),
    ("latinartcore.com",         "Tier 2", "Sin datos", "Sin datos", "N/D", 145, 120, "N/D", False),
]

for r, row in enumerate(bench_clean):
    dom, tier, tr_co, kw_co, as_co, tr_us, kw_us, as_us, hl = row
    if hl:
        f_text = fmt_highlight_text
        f_num  = fmt_highlight
    else:
        f_text = fmt_cell
        f_num  = fmt_num
    ws2.write(3 + r, 0, dom,   f_text)
    ws2.write(3 + r, 1, tier,  fmt_cell_c)
    ws2.write(3 + r, 2, tr_co if isinstance(tr_co, int) else tr_co, f_num if isinstance(tr_co, int) else fmt_cell_c)
    ws2.write(3 + r, 3, kw_co if isinstance(kw_co, int) else kw_co, f_num if isinstance(kw_co, int) else fmt_cell_c)
    ws2.write(3 + r, 4, as_co, fmt_cell_c)
    ws2.write(3 + r, 5, tr_us if isinstance(tr_us, int) else tr_us, f_num if isinstance(tr_us, int) else fmt_cell_c)
    ws2.write(3 + r, 6, kw_us if isinstance(kw_us, int) else kw_us, f_num if isinstance(kw_us, int) else fmt_cell_c)
    ws2.write(3 + r, 7, as_us, fmt_cell_c)

ws2.write(21, 0, "Nota: Authority Score no disponible en columnas estándar de domain_rank. N/D = No Disponible. 'Sin datos' = dominio no encontrado en base de datos Semrush.", fmt_meta)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3: Brechas de Keywords
# ══════════════════════════════════════════════════════════════════════════════
ws3 = wb.add_worksheet("Brechas de Keywords")
ws3.set_column("A:A", 35)
ws3.set_column("B:B", 18)
ws3.set_column("C:C", 22)
ws3.set_column("D:D", 30)
ws3.set_column("E:E", 22)
ws3.set_column("F:F", 18)
ws3.set_column("G:G", 50)

ws3.merge_range("A1:G1",
    "BRECHAS DE KEYWORDS — Keywords que rankean competidores y Duque Arango no | Datos al mayo 2026",
    fmt_title)

gap_headers = [
    "Keyword",
    "Volumen de Búsqueda",
    "Categoría SEO",
    "Competidor que Ranquea",
    "Posición Competidor",
    "Nivel de Oportunidad",
    "Acción de Contenido Sugerida",
]
for c, h in enumerate(gap_headers):
    ws3.write(2, c, h, fmt_header)

# Datos de brechas — top 15 por volumen, filtradas y relevantes
gap_data = [
    # keyword, vol, categoria, competidor, pos, nivel, accion
    ("museo de arte",             8100,  "Editorial",        "galeriaelmuseo.com (CO)",     34, "Alta",
     "Artículo editorial: 'Arte en Colombia — museos y galerías que definen la escena' con enlaces a artistas del roster."),
    ("wifredo lam",               5400,  "Nombre de Artista","cernudaarte.com (US)",        61, "Alta",
     "Crear/fortalecer página de artista Wifredo Lam en inglés: obras, estilo, importancia histórica, schema Artist."),
    ("omar rayo",                 5400,  "Nombre de Artista","galeriaduquearango.com ya rankea pos 6 — mejorar blog", 6, "Alta",
     "Redirigir tráfico de blog Omar Rayo a página de artista; añadir contenido de obras disponibles."),
    ("beatriz gonzalez",          3600,  "Nombre de Artista","casasriegner.com (CO)",        3, "Alta",
     "Crear artículo editorial sobre Beatriz González y contexto generacional con artistas del roster (Negret, Grau)."),
    ("galerias bogotá",           2900,  "Comercial",        "galeriaelmuseo.com / galerialacometa.com / casasriegner.com (CO)", 17, "Alta",
     "Optimizar página de sede Bogotá con copy que incluya 'galerías Bogotá'; añadir landing de colección Bogotá."),
    ("wilfredo lam",              2400,  "Nombre de Artista","cernudaarte.com (US)",        51, "Alta",
     "Añadir variante ortográfica 'wilfredo lam' en contenido y metadata de la página de Wifredo Lam."),
    ("jesus abad colorado",       2400,  "Nombre de Artista","galeriaelmuseo.com (CO)",      3, "Media",
     "Artículo editorial mencionando a Jesús Abad Colorado en contexto de fotografía y arte colombiano."),
    ("autorretrato",              2400,  "Editorial",        "galeriaelmuseo.com (CO)",      3, "Media",
     "Post editorial: 'El autorretrato en el arte latinoamericano' con ejemplos de Grau, Caballero, Morales."),
    ("emma reyes",                1900,  "Nombre de Artista","galerialacometa.com (CO)",     4, "Media",
     "Artículo sobre Emma Reyes y artistas colombianas del siglo XX: conexión con Ana Mercedes Hoyos y Alejandra Aristizábal."),
    ("feliza bursztyn",           1900,  "Nombre de Artista","galerialacometa.com (CO)",    32, "Media",
     "Artículo editorial sobre escultura abstracta colombiana: Feliza Bursztyn, Edgar Negret, Bernardo Salcedo."),
    ("pedro ruiz",                1600,  "Nombre de Artista","beatrizesguerra-art.com (US)", 8, "Media",
     "Artículo en inglés sobre figuración colombiana: Pedro Ruiz, Darío Morales, Luis Caballero."),
    ("amelia pelaez",             1000,  "Nombre de Artista","cernudaarte.com (US)",        49, "Media",
     "Post editorial: arte caribeño latinoamericano — puente entre Wifredo Lam, Amelia Peláez y la colección de la galería."),
    ("carlos rojas",              1000,  "Nombre de Artista","casasriegner.com (CO)",        2, "Media",
     "Artículo sobre abstracción geométrica latinoamericana: Carlos Rojas, Omar Rayo, Edgar Negret, Cruz-Diez."),
    ("bienal de arte medellin",    880,  "Editorial",        "casasriegner.com (CO)",       16, "Media",
     "Artículo sobre historia y relevancia de la Bienal de Medellín, posicionando la galería como parte del ecosistema."),
    ("wifredo lam paintings",      590,  "Nombre de Artista","cernudaarte.com (US)",        37, "Baja",
     "Añadir sección 'obras destacadas' en la página de Wifredo Lam con imágenes y fichas técnicas."),
]

for r, (kw, vol, cat, comp, pos, nivel, accion) in enumerate(gap_data):
    ws3.write(3 + r, 0, kw, fmt_cell)
    ws3.write(3 + r, 1, vol, fmt_num)
    ws3.write(3 + r, 2, cat, fmt_cell_c)
    ws3.write(3 + r, 3, comp, fmt_cell)
    ws3.write(3 + r, 4, pos, fmt_cell_c)
    if nivel == "Alta":
        ws3.write(3 + r, 5, nivel, fmt_alta)
    elif nivel == "Media":
        ws3.write(3 + r, 5, nivel, fmt_media)
    else:
        ws3.write(3 + r, 5, nivel, fmt_baja)
    ws3.write(3 + r, 6, accion, fmt_cell)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4: Páginas Top de Competencia
# ══════════════════════════════════════════════════════════════════════════════
ws4 = wb.add_worksheet("Páginas Top Competencia")
ws4.set_column("A:A", 28)
ws4.set_column("B:B", 12)
ws4.set_column("C:C", 60)
ws4.set_column("D:D", 16)
ws4.set_column("E:E", 30)
ws4.set_column("F:F", 25)
ws4.set_column("G:G", 50)

ws4.merge_range("A1:G1",
    "PÁGINAS TOP DE COMPETENCIA — Top 5 páginas por tráfico orgánico | Datos al mayo 2026",
    fmt_title)

comp_headers = [
    "Competidor", "Mercado", "URL de Página",
    "Tráfico Estimado", "Keyword Principal",
    "Tipo de Contenido", "Brecha para Duque Arango",
]
for c, h in enumerate(comp_headers):
    ws4.write(2, c, h, fmt_header)

comp_pages = [
    # CO top 3
    ("galeriaelmuseo.com", "CO", "www.galeriaelmuseo.com/", 1455, "galería el museo",
     "Homepage — Brand", "Nuestra homepage CO genera 841 visitas vs. 1.455 de El Museo. Optimizar SEO de homepage."),
    ("galeriaelmuseo.com", "CO", "www.galeriaelmuseo.com/archives/43710/", 669, "No identificado (artículo de blog)",
     "Artículo de blog", "El Museo usa arquitectura de blog para capturar tráfico editorial. Oportunidad: crear más artículos temáticos."),
    ("galeriaelmuseo.com", "CO", "www.galeriaelmuseo.com/archives/2084/", 503, "No identificado (artículo de blog)",
     "Artículo de blog", "Segundo artículo de alto tráfico. Duque Arango tiene equivalentes pero debe producir más volumen."),
    ("galeriaelmuseo.com", "CO", "www.galeriaelmuseo.com/archives/450/", 303, "No identificado (artículo de blog)",
     "Artículo de blog", "Cobertura amplia de artistas colombianos no incluidos en roster propio."),
    ("galeriaelmuseo.com", "CO", "www.galeriaelmuseo.com/archives/26977/", 204, "No identificado (artículo de blog)",
     "Artículo de blog", "Patrón claro: El Museo monetiza tráfico editorial. Estrategia replicable."),
    ("galerialacometa.com", "CO", "galerialacometa.com/", 1862, "galeria la cometa",
     "Homepage — Brand", "Tráfico de brand concentrado. Duque Arango supera ampliamente en keywords no-brand."),
    ("galerialacometa.com", "CO", "galerialacometa.com/exhibiciones/bogota/gabriela-pinilla-clandestina-es", 136, "Gabriela Pinilla",
     "Página de exposición", "Páginas de exposición individuales pueden capturar tráfico de artistas en exhibición."),
    ("galerialacometa.com", "CO", "galerialacometa.com/exhibiciones/bogota/emma-reyes-las-caras-de-emma-reyes-es", 82, "emma reyes",
     "Página de exposición / artista", "Página de Emma Reyes genera 82 visitas. Oportunidad: cubrir artistas contemporáneas colombianas."),
    ("galerialacometa.com", "CO", "galerialacometa.com/artistas/adrian-gaitan-es", 50, "Adrián Gaitán",
     "Página de artista", "Páginas de artistas individuales indexan bien. Fortalecer páginas de artistas contemporáneos propios."),
    ("casasriegner.com", "CO", "www.casasriegner.com/", 983, "casas riegner galería",
     "Homepage — Brand", "Tráfico de brand. Superados ampliamente por Duque Arango."),
    ("casasriegner.com", "CO", "www.casasriegner.com/artistas/beatriz-gonzalez", 180, "beatriz gonzalez artista",
     "Página de artista", "Beatriz González genera 180 visitas para Casas Riegner. Artista sin página en Duque Arango — brecha de contenido."),
    ("casasriegner.com", "CO", "www.casasriegner.com/artistas/carlos-rojas", 118, "carlos rojas",
     "Página de artista", "Carlos Rojas: artista de abstracción geométrica colombiana. Puede cubrirse editorialmente junto al roster."),
    ("casasriegner.com", "CO", "www.casasriegner.com/artistas/antonio-caro", 64, "antonio caro",
     "Página de artista", "Antonio Caro: artista conceptual colombiano. Oportunidad de artículo editorial contextual."),
    ("casasriegner.com", "CO", "www.casasriegner.com/artistas/rosemberg-sandoval", 61, "rosemberg sandoval",
     "Página de artista", "Artistas colombianos de renombre que el mercado busca activamente."),
    # US top 3
    ("operagallery.com", "US", "www.operagallery.com/", 1528, "opera gallery",
     "Homepage — Brand", "Opera Gallery lidera US con 7.263 visitas/mes. Su homepage genera 1.528. Enfoque multi-artista global."),
    ("operagallery.com", "US", "www.operagallery.com/artist/simon-hantai", 665, "simon hantai",
     "Página de artista", "Artista abstracto francés. Opera Gallery posiciona artistas internacionales de alto valor comercial."),
    ("operagallery.com", "US", "www.operagallery.com/artist/keith-haring", 596, "keith haring",
     "Página de artista", "Keith Haring: 596 visitas. Opera Gallery trabaja artistas de mercado secundario global."),
    ("operagallery.com", "US", "www.operagallery.com/artist/marc-chagall", 581, "marc chagall",
     "Página de artista", "Chagall, Haring, Botero en la misma galería. Estrategia de anclar artistas trophy assets."),
    ("operagallery.com", "US", "www.operagallery.com/viewing-rooms/fernando-botero", 557, "fernando botero viewing room",
     "Viewing Room — Artista roster", "⚠️ AMENAZA DIRECTA: Opera Gallery tiene viewing room de Botero (557 visitas/mes US). Artista de nuestro roster. Crear equivalente."),
    ("miguelabreugallery.com", "US", "miguelabreugallery.com/", 2402, "miguel abreu gallery",
     "Homepage — Brand", "75.6% del tráfico en homepage. Modelo brand-centric. Su audiencia es muy específica y especializada."),
    ("miguelabreugallery.com", "US", "miguelabreugallery.com/artists/quaytman/", 158, "quaytman",
     "Página de artista", "Artista conceptual de alto nivel. Audiencia institucional/coleccionistas sofisticados."),
    ("miguelabreugallery.com", "US", "miguelabreugallery.com/artists/francois-marie-banier/", 138, "francois marie banier",
     "Página de artista", "François-Marie Banier: artista multidisciplinario. Ejemplo de páginas de artista con buen SEO."),
    ("forumgallery.com", "US", "www.forumgallery.com/", 891, "forum gallery",
     "Homepage — Brand", "Forum Gallery fuerte en figuración realista. Genera 1.930 visitas/mes en US."),
    ("forumgallery.com", "US", "forumgallery.com/artists/holly-lane/biography", 237, "holly lane",
     "Biografía de artista", "Páginas de BIOGRAFÍA generan tráfico significativo. Duque Arango puede añadir secciones biográficas ricas."),
    ("forumgallery.com", "US", "www.forumgallery.com/artists/claudio-bravo/videos", 114, "claudio bravo",
     "Página de artista con video", "Claudio Bravo (figuración sudamericana) genera tráfico en US. Conexión directa con artistas del roster de Duque Arango."),
]

for r, row in enumerate(comp_pages):
    comp, merc, url, tr, kw, tipo, brecha = row
    ws4.write(3 + r, 0, comp, fmt_cell)
    ws4.write(3 + r, 1, merc, fmt_cell_c)
    ws4.write(3 + r, 2, url, fmt_cell)
    ws4.write(3 + r, 3, tr, fmt_num)
    ws4.write(3 + r, 4, kw, fmt_cell)
    ws4.write(3 + r, 5, tipo, fmt_cell)
    ws4.write(3 + r, 6, brecha, fmt_cell)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 5: Insights y Recomendaciones
# ══════════════════════════════════════════════════════════════════════════════
ws5 = wb.add_worksheet("Insights y Recomendaciones")
ws5.set_column("A:A", 40)
ws5.set_column("B:B", 45)
ws5.set_column("C:C", 35)
ws5.set_column("D:D", 55)
ws5.set_column("E:E", 12)

ws5.merge_range("A1:E1",
    "INSIGHTS Y RECOMENDACIONES ESTRATÉGICAS — Semana del 2026-05-26",
    fmt_title)

ins_headers = [
    "Insight", "Evidencia", "Framework Aplicado",
    "Acción Recomendada", "Prioridad",
]
for c, h in enumerate(ins_headers):
    ws5.write(2, c, h, fmt_header)

insights = [
    (
        "⚠️ Opera Gallery tiene viewing room de Fernando Botero (artista del roster) generando 557 visitas/mes en US",
        "operagallery.com/viewing-rooms/fernando-botero es la 5ª página más visitada de ese dominio en US con 557 visitas/mes. Fernando Botero está en el roster de Duque Arango.",
        "competitor-alternatives: Research Process — competitive page analysis. content-strategy: Consideration Stage (coleccionistas evaluando compra).",
        "Crear página de colección / viewing room dedicada a Fernando Botero en inglés, con schema de ArtGallery, imágenes de obras disponibles, y texto editorial optimizado para 'fernando botero for sale', 'buy botero art'.",
        "Alta",
    ),
    (
        "Wifredo Lam (artista del roster) tiene 5.400 búsquedas/mes en US y la galería no rankea para ese keyword",
        "cernudaarte.com aparece en pos 61 para 'wifredo lam' (5.400 vol/mes US) y 'wilfredo lam' (2.400 vol/mes US). Duque Arango no aparece pese a tener a Lam en su roster.",
        "seo-audit: On-Page SEO — Keyword Targeting. ai-seo: Pillar 2 Authority — E-E-A-T signals para artista de renombre internacional.",
        "Crear o fortalecer página de artista /en/artist/wifredo-lam/ en inglés con: definición de su estilo (surrealismo caribeño), obras representativas con fichas técnicas, FAQPage schema, y sección 'obras disponibles en galería'. Añadir variante 'wilfredo lam'.",
        "Alta",
    ),
    (
        "Duque Arango domina Colombia 5.5x sobre el competidor más cercano pero es 3er lugar en US",
        "CO: Duque Arango 23.273 visitas vs. galeriaelmuseo.com 4.252 (segundo lugar). US: operagallery.com 7.263 > miguelabreugallery.com 3.176 > Duque Arango 2.296.",
        "competitor-alternatives: Competitive Landscape Analysis. content-strategy: Búsqueda por etapa — Decision Stage en mercado US.",
        "Priorizar producción de contenido en inglés orientado a coleccionistas US: páginas de artistas en inglés, artículos de mercado (precios, proveniencia), y landing pages comerciales para los artistas con mayor demanda (Botero, Lam, Ana Mercedes Hoyos).",
        "Alta",
    ),
    (
        "Edgar Negret concentra el 33.69% del tráfico orgánico CO en una sola página — riesgo de dependencia",
        "Página /artista/edgar-negret/ genera 7.841 visitas/mes CO (33.69% del total). El keyword 'edgar negret' tiene 90.500 búsquedas/mes. Rankeamos pos 2, muy cerca del top 1.",
        "seo-audit: Content Quality — E-E-A-T signals para defender posición. ai-seo: Content Extractability Check — estructura para citación en AI Overviews.",
        "Fortalecer la página de Edgar Negret: añadir datos de proveniencia de obras, precios de subasta históricos citando fuentes, FAQ sobre su obra (optimizado para AI Overviews), schema de Artist y ArtWork. Crear artículo complementario en inglés para capturar tráfico US.",
        "Alta",
    ),
    (
        "'Galerías Bogotá' — término comercial con 2.900 búsquedas/mes CO donde 3 competidores rankean y Duque Arango no aparece",
        "galeriaelmuseo.com pos 17, galerialacometa.com pos 23, casasriegner.com pos 22 para 'galerias bogotá' (2.900 vol CO). Duque Arango tiene sede en Bogotá pero no rankea.",
        "seo-audit: On-Page SEO — Keyword Targeting y Site Architecture. content-strategy: Decision Stage — intención comercial de descubrimiento de galerías.",
        "Optimizar la página de sede Bogotá con copy que incluya naturalmente 'galerías Bogotá', añadir datos de dirección estructurados con schema LocalBusiness, y crear landing page de colección Bogotá diferenciada de Medellín.",
        "Alta",
    ),
    (
        "Artistas contemporáneos exclusivos del roster (Caraballo, Aristizábal, Vélez, Echeverri) sin presencia en top 10 páginas CO ni US",
        "Top 10 páginas CO y US dominadas por Botero, Edgar Negret, David Manzur, Guayasamín. Artistas contemporáneos exclusivos como Javier Caraballo (vol 140 CO, pos 1) tienen presencia pero no escalan.",
        "content-strategy: Content Pillars — artistas contemporáneos como pilar diferenciador. seo-audit: Content Gap — páginas de artistas sin suficiente profundidad editorial.",
        "Crear artículos editoriales de los artistas contemporáneos exclusivos con intención informacional: 'Quién es Javier Caraballo y por qué importa', 'La obra de Gustavo Vélez', 'Alejandra Aristizábal: entre abstracción y territorio'. Añadir FAQs y schema de artista.",
        "Media",
    ),
    (
        "Blog en español captura tráfico masivo CO pero el blog en inglés tiene bajo rendimiento por artículo (~200 visitas promedio)",
        "Blog ES: top 3 páginas CO promedian 3.600 visitas/artículo. Blog EN: top 5 páginas US promedian 194 visitas/artículo. La página de artista Ana Mercedes Hoyos EN genera 521 visitas con solo 2 keywords — potencial no explotado.",
        "ai-seo: Pillar 1 Structure — Content Extractability para citación por AI. seo-audit: Content Optimization — depth y long-tail keyword coverage.",
        "Auditar las 5 páginas de artistas EN con mayor tráfico para añadir: FAQ schema, bloques de definición (qué hace único a este artista), estadísticas de mercado, y 'last updated' date. La página de Ana Mercedes Hoyos podría triplicar tráfico con optimización básica.",
        "Media",
    ),
    (
        "Páginas de artistas de Casas Riegner (Beatriz González 180 visitas, Carlos Rojas 118) evidencian demanda no cubierta por el roster",
        "casasriegner.com/artistas/beatriz-gonzalez genera 180 visitas/mes CO; /artistas/carlos-rojas genera 118. Ambos artistas tienen alta búsqueda y son pilares del arte colombiano moderno — contexto natural para el roster de Duque Arango.",
        "competitor-alternatives: Research Process — identificar qué pages convierten para competidores. content-strategy: Awareness Stage — artistas que el coleccionista ya conoce y busca.",
        "Crear artículos editoriales contextuales sobre estos artistas (no páginas de artista propias ya que no los representan), posicionando a Duque Arango como autoridad editorial del arte colombiano moderno. Enlazar desde dichos artículos hacia artistas del roster de época similar.",
        "Media",
    ),
]

for r, (ins, ev, fw, acc, pri) in enumerate(insights):
    ws5.write(3 + r, 0, ins, fmt_cell)
    ws5.write(3 + r, 1, ev, fmt_cell)
    ws5.write(3 + r, 2, fw, fmt_cell)
    ws5.write(3 + r, 3, acc, fmt_cell)
    if pri == "Alta":
        ws5.write(3 + r, 4, pri, fmt_alta)
    elif pri == "Media":
        ws5.write(3 + r, 4, pri, fmt_media)
    else:
        ws5.write(3 + r, 4, pri, fmt_baja)

# Row heights for readability
for ws in [ws1, ws2, ws3, ws4, ws5]:
    ws.set_default_row(18)

wb.close()

# Verify
size = os.path.getsize(FILEPATH)
print(f"Archivo generado: {FILEPATH}")
print(f"Tamaño: {size:,} bytes")
print("OK" if size > 0 else "ERROR: archivo vacío")
