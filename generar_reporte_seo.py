#!/usr/bin/env python3
"""Genera el reporte semanal de SEO para Galería Duque Arango."""

import xlsxwriter
import os

DATE = "2026-05-28"
FILENAME = f"Inteligencia_SEO_{DATE}.xlsx"
OUTPUT_PATH = os.path.join("/home/user/da-routines", FILENAME)

wb = xlsxwriter.Workbook(OUTPUT_PATH)

# ── Formatos ──────────────────────────────────────────────────────────────────
fmt_title = wb.add_format({
    "bold": True, "font_size": 14, "font_color": "#FFFFFF",
    "bg_color": "#1A1A2E", "align": "center", "valign": "vcenter",
    "border": 1
})
fmt_header = wb.add_format({
    "bold": True, "font_color": "#FFFFFF", "bg_color": "#16213E",
    "align": "center", "valign": "vcenter", "border": 1, "text_wrap": True
})
fmt_subheader = wb.add_format({
    "bold": True, "font_color": "#FFFFFF", "bg_color": "#0F3460",
    "align": "center", "valign": "vcenter", "border": 1
})
fmt_highlight = wb.add_format({
    "bold": True, "font_color": "#FFFFFF", "bg_color": "#E94560",
    "align": "center", "valign": "vcenter", "border": 1
})
fmt_cell = wb.add_format({
    "align": "left", "valign": "vcenter", "border": 1,
    "text_wrap": True, "font_size": 10
})
fmt_num = wb.add_format({
    "num_format": "#,##0", "align": "right", "valign": "vcenter",
    "border": 1, "font_size": 10
})
fmt_num_bold = wb.add_format({
    "bold": True, "num_format": "#,##0", "align": "right",
    "valign": "vcenter", "border": 1, "font_size": 10,
    "bg_color": "#E94560", "font_color": "#FFFFFF"
})
fmt_pct = wb.add_format({
    "num_format": "0.00%", "align": "right", "valign": "vcenter",
    "border": 1, "font_size": 10
})
fmt_url = wb.add_format({
    "font_color": "#0563C1", "underline": True, "align": "left",
    "valign": "vcenter", "border": 1, "font_size": 9, "text_wrap": True
})
fmt_alta = wb.add_format({
    "bold": True, "font_color": "#FFFFFF", "bg_color": "#C0392B",
    "align": "center", "valign": "vcenter", "border": 1
})
fmt_media = wb.add_format({
    "bold": True, "font_color": "#FFFFFF", "bg_color": "#E67E22",
    "align": "center", "valign": "vcenter", "border": 1
})
fmt_baja = wb.add_format({
    "bold": True, "font_color": "#FFFFFF", "bg_color": "#27AE60",
    "align": "center", "valign": "vcenter", "border": 1
})
fmt_duque = wb.add_format({
    "bold": True, "bg_color": "#E94560", "font_color": "#FFFFFF",
    "align": "left", "valign": "vcenter", "border": 1, "font_size": 10
})
fmt_duque_num = wb.add_format({
    "bold": True, "bg_color": "#E94560", "font_color": "#FFFFFF",
    "num_format": "#,##0", "align": "right", "valign": "vcenter",
    "border": 1, "font_size": 10
})
fmt_note = wb.add_format({
    "italic": True, "font_color": "#666666", "font_size": 9,
    "align": "left", "valign": "vcenter"
})

# ── TAB 1: Desempeño Propio ────────────────────────────────────────────────
ws1 = wb.add_worksheet("Desempeño Propio")
ws1.set_zoom(90)
ws1.set_row(0, 30)

ws1.merge_range("A1:D1",
    f"GALERÍA DUQUE ARANGO — DESEMPEÑO SEO PROPIO | Datos al {DATE}",
    fmt_title)
ws1.set_column("A:A", 18)
ws1.set_column("B:B", 32)
ws1.set_column("C:C", 22)
ws1.set_column("D:D", 48)

# Métricas generales
ws1.merge_range("A3:D3", "MÉTRICAS GENERALES DEL DOMINIO", fmt_subheader)
headers_metricas = ["Base de Datos", "Métrica", "Valor", "Notas"]
for c, h in enumerate(headers_metricas):
    ws1.write(3, c, h, fmt_header)

metricas = [
    ("Colombia (CO)", "Keywords orgánicas", "2.109",
     "Semrush rank #4.681 — posición sólida en Colombia"),
    ("Colombia (CO)", "Tráfico orgánico estimado", "23.040 visitas/mes",
     "Líder indiscutible del sector en CO"),
    ("Colombia (CO)", "Valor orgánico estimado", "$2.032 USD/mes",
     "Costo equivalente si se pagara por los mismos clicks en Ads"),
    ("Colombia (CO)", "Inversión en Ads", "Sin inversión",
     "El tráfico es 100% orgánico"),
    ("Estados Unidos (US)", "Keywords orgánicas", "1.914",
     "Semrush rank #585.634 — margen de crecimiento en US"),
    ("Estados Unidos (US)", "Tráfico orgánico estimado", "2.357 visitas/mes",
     "#3 de 17 galerías en el ranking US"),
    ("Estados Unidos (US)", "Valor orgánico estimado", "$393 USD/mes",
     "Oportunidad de monetización vía contenido en inglés"),
    ("Estados Unidos (US)", "Inversión en Ads", "Sin inversión",
     "El tráfico es 100% orgánico"),
]
for r, (db, met, val, nota) in enumerate(metricas, start=4):
    ws1.write(r, 0, db, fmt_cell)
    ws1.write(r, 1, met, fmt_cell)
    ws1.write(r, 2, val, fmt_cell)
    ws1.write(r, 3, nota, fmt_cell)

# Top 20 Keywords CO
ws1.merge_range("A13:D13", "TOP 20 KEYWORDS — COLOMBIA (CO)", fmt_subheader)
headers_kw = ["Keyword", "Posición", "Volumen de Búsqueda", "URL que Rankea"]
for c, h in enumerate(headers_kw):
    ws1.write(13, c, h, fmt_header)

kw_co = [
    ("edgar negret", 2, 90500, "galeriaduquearango.com/artista/edgar-negret/"),
    ("plaza botero medellín antioquia", 7, 22200, "galeriaduquearango.com/blog/plaza-botero-resignificar-el-espacio-publico..."),
    ("plaza botero", 8, 49500, "galeriaduquearango.com/blog/plaza-botero-resignificar-el-espacio-publico..."),
    ("expresionismo", 10, 5400, "galeriaduquearango.com/blog/el-expresionismo-en-latinoamerica-y-5-exponentes/"),
    ("david manzur", 2, 4400, "galeriaduquearango.com/artista/david-manzur/"),
    ("guayasamin", 3, 2400, "galeriaduquearango.com/blog/el-legado-al-arte-de-oswaldo-guayasamin/"),
    ("enrique grau", 2, 1600, "galeriaduquearango.com/artista/enrique-grau/"),
    ("muralismo", 3, 1600, "galeriaduquearango.com/blog/muralismo-latinoamericano..."),
    ("galeria duque arango", 1, 720, "galeriaduquearango.com/"),
    ("obras de alejandro obregón", 1, 390, "galeriaduquearango.com/artista/alejandro-obregon/"),
    ("pinturas de alejandro obregón", 1, 210, "galeriaduquearango.com/artista/alejandro-obregon/"),
    ("alejandro obregon obras", 1, 210, "galeriaduquearango.com/artista/alejandro-obregon/"),
    ("ana mercedes hoyos obras", 1, 210, "galeriaduquearango.com/artista/ana-mercedes-hoyos/"),
    ("obras de arte de david manzur", 1, 170, "galeriaduquearango.com/artista/david-manzur/"),
    ("botero cuadros", 1, 170, "galeriaduquearango.com/blog/las-obras-mejor-vendidas-de-fernando-botero/"),
    ("edgar negret obras", 1, 140, "galeriaduquearango.com/artista/edgar-negret/"),
    ("javier caraballo", 1, 140, "galeriaduquearango.com/artista/javier-caraballo/"),
    ("duque arango galeria", 1, 110, "galeriaduquearango.com/"),
    ("grau", 2, 1300, "galeriaduquearango.com/artista/enrique-grau/"),
    ("edgar negret (blog)", 19, 90500, "galeriaduquearango.com/blog/edgar-negret-el-escultor-colombiano..."),
]
for r, (kw, po, vol, url) in enumerate(kw_co, start=14):
    ws1.write(r, 0, kw, fmt_cell)
    ws1.write(r, 1, po, fmt_num)
    ws1.write(r, 2, vol, fmt_num)
    ws1.write(r, 3, url, fmt_url)

# Top 20 Keywords US
ws1.merge_range("A35:D35", "TOP 20 KEYWORDS — ESTADOS UNIDOS (US)", fmt_subheader)
for c, h in enumerate(headers_kw):
    ws1.write(35, c, h, fmt_header)

kw_us = [
    ("fernando botero", 7, 14800, "galeriaduquearango.com/en/blog/unmistakable-brand-boterism/"),
    ("mexican art", 15, 8100, "galeriaduquearango.com/en/blog/mexican-art-history..."),
    ("botero", 9, 9900, "galeriaduquearango.com/en/blog/all-you-need-to-know..."),
    ("fernando botero paintings", 9, 4400, "galeriaduquearango.com/en/blog/all-you-need-to-know..."),
    ("ana mercedes hoyos", 7, 3600, "galeriaduquearango.com/artista/ana-mercedes-hoyos/"),
    ("fernando botero art", 5, 2400, "galeriaduquearango.com/en/blog/unmistakable-brand-boterism/"),
    ("plaza botero medellín colombia", 6, 1900, "galeriaduquearango.com/en/blog/plaza-botero-resignifying..."),
    ("botero's", 6, 1600, "galeriaduquearango.com/en/blog/all-you-need-to-know..."),
    ("oswaldo guayasamin", 7, 1600, "galeriaduquearango.com/en/blog/oswaldo-guayasamins-legacy..."),
    ("guayasamin", 4, 1300, "galeriaduquearango.com/en/blog/the-works-of-oswaldo-guayasamin/"),
    ("fernando botero artworks", 9, 1300, "galeriaduquearango.com/en/blog/all-you-need-to-know..."),
    ("colombian painters", 2, 720, "galeriaduquearango.com/en/blog/discovering-colombian-art..."),
    ("plaza botero", 8, 1000, "galeriaduquearango.com/en/blog/plaza-botero-resignifying..."),
    ("picasso realism", 7, 590, "galeriaduquearango.com/en/blog/artistic-periods-of-pablo-picasso/"),
    ("boterismo", 3, 480, "galeriaduquearango.com/en/blog/unmistakable-brand-boterism/"),
    ("picasso realistic art", 6, 480, "galeriaduquearango.com/en/blog/artistic-periods-of-pablo-picasso/"),
    ("guayasamin artist", 4, 390, "galeriaduquearango.com/en/blog/the-works-of-oswaldo-guayasamin/"),
    ("fernando botero style", 1, 170, "galeriaduquearango.com/en/blog/unmistakable-brand-boterism/"),
    ("colombian painters famous", 2, 170, "galeriaduquearango.com/en/blog/discovering-colombian-art..."),
    ("nombre de las 30 pinturas más vistas de botero", 2, 210, "galeriaduquearango.com/blog/las-obras-mejor-vendidas-de-fernando-botero/"),
]
for r, (kw, po, vol, url) in enumerate(kw_us, start=36):
    ws1.write(r, 0, kw, fmt_cell)
    ws1.write(r, 1, po, fmt_num)
    ws1.write(r, 2, vol, fmt_num)
    ws1.write(r, 3, url, fmt_url)

# Top 10 Páginas CO
ws1.merge_range("A57:D57", "TOP 10 PÁGINAS — COLOMBIA (CO)", fmt_subheader)
headers_pg = ["URL de la Página", "Tráfico Estimado", "Nº de Keywords", "Keyword Principal"]
for c, h in enumerate(headers_pg):
    ws1.write(57, c, h, fmt_header)

pages_co = [
    ("galeriaduquearango.com/artista/edgar-negret/", 7841, 24, "edgar negret"),
    ("galeriaduquearango.com/blog/plaza-botero-resignificar...", 4177, 34, "plaza botero medellín antioquia"),
    ("galeriaduquearango.com/artista/david-manzur/", 2700, 51, "david manzur"),
    ("galeriaduquearango.com/blog/las-obras-de-fernando-botero-y-su-significado/", 1664, 143, "obras de fernando botero"),
    ("galeriaduquearango.com/artista/alejandro-obregon/", 861, 67, "alejandro obregon obras"),
    ("galeriaduquearango.com/ (homepage)", 846, 99, "galeria duque arango"),
    ("galeriaduquearango.com/blog/las-obras-mejor-vendidas-de-fernando-botero/", 793, 92, "botero cuadros"),
    ("galeriaduquearango.com/artista/ana-mercedes-hoyos/", 452, 19, "ana mercedes hoyos obras"),
    ("galeriaduquearango.com/blog/artistas-colombianos-que-debes-conocer/", 353, 174, "artistas colombianos"),
    ("galeriaduquearango.com/artista/fernando-botero/", 291, 169, "fernando botero obras"),
]
for r, (url, tr, nkw, top_kw) in enumerate(pages_co, start=58):
    ws1.write(r, 0, url, fmt_url)
    ws1.write(r, 1, tr, fmt_num)
    ws1.write(r, 2, nkw, fmt_num)
    ws1.write(r, 3, top_kw, fmt_cell)

# Top 10 Páginas US
ws1.merge_range("A69:D69", "TOP 10 PÁGINAS — ESTADOS UNIDOS (US)", fmt_subheader)
for c, h in enumerate(headers_pg):
    ws1.write(69, c, h, fmt_header)

pages_us = [
    ("galeriaduquearango.com/artista/ana-mercedes-hoyos/", 521, 2, "ana mercedes hoyos"),
    ("galeriaduquearango.com/en/blog/unmistakable-brand-boterism/", 282, 76, "fernando botero style"),
    ("galeriaduquearango.com/en/blog/all-you-need-to-know-to-understand-fernando-boteros-artworks/", 251, 100, "fernando botero paintings"),
    ("galeriaduquearango.com/en/blog/the-works-of-oswaldo-guayasamin/", 217, 43, "guayasamin"),
    ("galeriaduquearango.com/en/blog/discovering-colombian-art-from-fernando-botero-to-contemporary-masters/", 193, 72, "colombian painters"),
    ("galeriaduquearango.com/en/artist/luis-caballero/", 88, 14, "luis caballero"),
    ("galeriaduquearango.com/en/blog/artistic-periods-of-pablo-picasso/", 74, 142, "picasso realism"),
    ("galeriaduquearango.com/blog/las-obras-de-fernando-botero-y-su-significado/", 69, 34, "botero obras"),
    ("galeriaduquearango.com/en/blog/plaza-plaza-botero-resignifying-public-space-through-art-in-medellin/", 66, 7, "plaza botero colombia"),
    ("galeriaduquearango.com/en/blog/the-works-of-fernando-botero-and-their-significance/", 65, 61, "fernando botero artworks"),
]
for r, (url, tr, nkw, top_kw) in enumerate(pages_us, start=70):
    ws1.write(r, 0, url, fmt_url)
    ws1.write(r, 1, tr, fmt_num)
    ws1.write(r, 2, nkw, fmt_num)
    ws1.write(r, 3, top_kw, fmt_cell)

ws1.write(81, 0,
    "Nota: Datos de Semrush — actualización mensual. Tráfico = estimado de visitas orgánicas.",
    fmt_note)

# ── TAB 2: Benchmark Competitivo ─────────────────────────────────────────────
ws2 = wb.add_worksheet("Benchmark Competitivo")
ws2.set_zoom(85)
ws2.set_row(0, 30)
ws2.set_row(1, 45)

ws2.merge_range("A1:H1",
    f"BENCHMARK COMPETITIVO — 17 GALERÍAS | Datos al {DATE}",
    fmt_title)

headers_bm = [
    "Dominio", "Tier",
    "Tráfico CO", "Keywords CO", "Rank Semrush CO",
    "Tráfico US", "Keywords US", "Rank Semrush US"
]
cols_bm = [32, 8, 14, 14, 16, 14, 14, 16]
for c, (h, w) in enumerate(zip(headers_bm, cols_bm)):
    ws2.write(1, c, h, fmt_header)
    ws2.set_column(c, c, w)

# Ordenado por Tráfico CO descendente
benchmarks = [
    ("galeriaduquearango.com", "Propia", 23040, 2109, 4681, 2357, 1914, 585634),
    ("galeriaelmuseo.com", "Tier 1", 4297, 507, 16534, 28, 43, 5976598),
    ("galerialacometa.com", "Tier 1", 3174, 350, 20640, 511, 145, 1640651),
    ("casasriegner.com", "Tier 1", 1656, 235, 32784, 27, 58, 6030082),
    ("sgr-art.com", "Tier 1", 668, 62, 60024, 0, 4, 21888652),
    ("galeriacasacuadrada.com", "Tier 1", 514, 62, 70771, 0, 1, 34515550),
    ("beatrizesguerra-art.com", "Tier 1", 169, 41, 134872, 135, 94, 3277854),
    ("galeriafreites.com", "Tier 2", 166, 107, 136340, 13, 18, 7443941),
    ("otros360grados.com", "Tier 1", 128, 58, 156524, 0, 6, 18719484),
    ("operagallery.com", "Tier 2", 54, 50, 238084, 7402, 3470, 233573),
    ("galeriaelsapineres.art", "Tier 1", 15, 20, 417543, 0, 0, 0),
    ("artoftheworldgallery.com", "Tier 2", 12, 21, 459976, 1447, 721, 836564),
    ("latinartcore.com", "Tier 2", 0, 0, 0, 145, 122, 3165891),
    ("ascasogallery.com", "Tier 2", 0, 11, 1289899, 576, 161, 1528598),
    ("cernudaarte.com", "Tier 2", 0, 5, 1546762, 553, 414, 1565424),
    ("forumgallery.com", "Tier 2", 0, 5, 1581541, 2008, 1554, 660152),
    ("miguelabreugallery.com", "Tier 2", 0, 3, 2089765, 3214, 460, 461693),
]

for r, row in enumerate(benchmarks, start=2):
    domain, tier, tr_co, kw_co_v, rank_co, tr_us, kw_us_v, rank_us = row
    if domain == "galeriaduquearango.com":
        ws2.write(r, 0, domain, fmt_duque)
        ws2.write(r, 1, tier, fmt_duque)
        ws2.write(r, 2, tr_co if tr_co else "Sin datos", fmt_duque_num if tr_co else fmt_duque)
        ws2.write(r, 3, kw_co_v if kw_co_v else "Sin datos", fmt_duque_num if kw_co_v else fmt_duque)
        ws2.write(r, 4, rank_co if rank_co else "Sin datos", fmt_duque_num if rank_co else fmt_duque)
        ws2.write(r, 5, tr_us if tr_us else "Sin datos", fmt_duque_num if tr_us else fmt_duque)
        ws2.write(r, 6, kw_us_v if kw_us_v else "Sin datos", fmt_duque_num if kw_us_v else fmt_duque)
        ws2.write(r, 7, rank_us if rank_us else "Sin datos", fmt_duque_num if rank_us else fmt_duque)
    else:
        ws2.write(r, 0, domain, fmt_cell)
        ws2.write(r, 1, tier, fmt_cell)
        ws2.write(r, 2, tr_co if tr_co else "Sin datos", fmt_num if tr_co else fmt_cell)
        ws2.write(r, 3, kw_co_v if kw_co_v else "Sin datos", fmt_num if kw_co_v else fmt_cell)
        ws2.write(r, 4, rank_co if rank_co else "Sin datos", fmt_num if rank_co else fmt_cell)
        ws2.write(r, 5, tr_us if tr_us else "Sin datos", fmt_num if tr_us else fmt_cell)
        ws2.write(r, 6, kw_us_v if kw_us_v else "Sin datos", fmt_num if kw_us_v else fmt_cell)
        ws2.write(r, 7, rank_us if rank_us else "Sin datos", fmt_num if rank_us else fmt_cell)

ws2.write(21, 0,
    "Notas: Rank Semrush = posición global del dominio (menor = mayor autoridad). "
    "latinartcore.com y galeriaelsapineres.art sin datos en CO. "
    "Datos al " + DATE + ".",
    fmt_note)
ws2.write(22, 0,
    "Posición en CO: galeriaduquearango.com es #1 de 17 galerías. "
    "Posición en US: #3 de 17 galerías (detrás de operagallery.com y miguelabreugallery.com).",
    fmt_note)

# ── TAB 3: Brechas de Keywords ─────────────────────────────────────────────
ws3 = wb.add_worksheet("Brechas de Keywords")
ws3.set_zoom(85)
ws3.set_row(0, 30)
ws3.set_row(1, 45)

ws3.merge_range("A1:G1",
    f"BRECHAS DE KEYWORDS — CO + US | Datos al {DATE}",
    fmt_title)

headers_gap = [
    "Keyword", "Volumen de Búsqueda", "Categoría SEO",
    "Competidor que Rankea", "Posición Competidor",
    "Nivel de Oportunidad", "Acción de Contenido Sugerida"
]
cols_gap = [28, 18, 18, 28, 18, 18, 52]
for c, (h, w) in enumerate(zip(headers_gap, cols_gap)):
    ws3.write(1, c, h, fmt_header)
    ws3.set_column(c, c, w)

# Mercado CO
ws3.merge_range("A3:G3", "COLOMBIA (CO) — Keywords que competidores rankean y galeriaduquearango.com no", fmt_subheader)

gaps_co = [
    ("galerias bogotá", 2900, "Comercial", "galeriaelmuseo.com / galerialacometa.com / casasriegner.com", "pos 12–21 (múltiples)", "Alta",
     "Optimizar homepage y crear página de presencia Bogotá: 'Galería de arte en Bogotá — Duque Arango'"),
    ("beatriz gonzalez", 3600, "Nombre de Artista", "casasriegner.com", "pos 3", "Alta",
     "Crear artículo editorial 'Beatriz González: la artista colombiana que reinterpretó lo cotidiano' con contexto histórico"),
    ("museo de arte", 8100, "Editorial", "galeriaelmuseo.com", "pos 45", "Alta",
     "Artículo 'Galerías vs Museos de Arte en Colombia: diferencias y complementos' con SEO interno"),
    ("autorretrato", 2400, "Editorial", "galeriaelmuseo.com", "pos 3", "Alta",
     "Artículo 'El autorretrato en el arte latinoamericano' destacando obras de artistas del roster (Botero, Obregón, Manzur)"),
    ("jesus abad colorado", 2400, "Nombre de Artista", "galeriaelmuseo.com", "pos 3", "Alta",
     "Artículo editorial 'Jesús Abad Colorado y la fotografía documental colombiana' — posicionarse como referencia editorial"),
    ("emma reyes", 1900, "Nombre de Artista", "galerialacometa.com", "pos 4", "Alta",
     "Artículo 'Emma Reyes: la artista colombiana autodidacta del siglo XX' — serie 'Maestros Colombianos'"),
    ("feliza bursztyn", 1900, "Nombre de Artista", "galerialacometa.com", "pos 32", "Alta",
     "Artículo 'Feliza Bursztyn y la escultura cinética en Colombia' — contexto para posicionar a Edgar Negret y Omar Rayo"),
    ("galeria de arte", 1300, "Comercial", "casasriegner.com / galerialacometa.com", "pos 2–6", "Media",
     "Revisar y optimizar meta title + H1 del homepage para incluir 'galería de arte Medellín Bogotá'"),
    ("bienal de arte medellin", 880, "Editorial", "casasriegner.com", "pos 16", "Media",
     "Crear nota editorial sobre artistas de la galería en la Bienal de Arte de Antioquia con fotos y contexto curatorial"),
    ("carlos jacanamijoy", 1000, "Nombre de Artista", "galeriaelmuseo.com", "pos 5", "Media",
     "Artículo 'Carlos Jacanamijoy y el arte indígena contemporáneo colombiano' — SEO editorial de largo plazo"),
    ("pedro ruiz", 880, "Nombre de Artista", "galeriaelmuseo.com / beatrizesguerra-art.com", "pos 12–30", "Media",
     "Artículo sobre Pedro Ruiz y la nueva figuración colombiana — vincular con artistas del roster de la galería"),
    ("antonio caro", 720, "Nombre de Artista", "casasriegner.com", "pos 3", "Baja",
     "Artículo educativo sobre Antonio Caro y el conceptualismo en Colombia — refuerza autoridad editorial"),
    ("art gallery", 1900, "Comercial", "galerialacometa.com", "pos 2", "Alta",
     "Optimizar versión en inglés para 'art gallery Colombia' y 'contemporary art gallery Medellín'"),
    ("galeria de arte bogota", 590, "Comercial", "casasriegner.com", "pos 1", "Baja",
     "Crear página de ubicación Bogotá optimizada: 'Galería de Arte en Bogotá — Duque Arango'"),
    ("lemuria arte", 1300, "Editorial", "casasriegner.com", "pos 3", "Media",
     "Cobertura de exposición 'Lemuria' y otras muestras conceptuales — contenido de feria y exposiciones"),
]

for r, row in enumerate(gaps_co, start=3):
    kw, vol, cat, comp, pos, nivel, accion = row
    ws3.write(r, 0, kw, fmt_cell)
    ws3.write(r, 1, vol, fmt_num)
    ws3.write(r, 2, cat, fmt_cell)
    ws3.write(r, 3, comp, fmt_cell)
    ws3.write(r, 4, pos, fmt_cell)
    nf = fmt_alta if nivel == "Alta" else (fmt_media if nivel == "Media" else fmt_baja)
    ws3.write(r, 5, nivel, nf)
    ws3.write(r, 6, accion, fmt_cell)

# Mercado US
us_row = len(gaps_co) + 4
ws3.merge_range(us_row, 0, us_row, 6,
    "ESTADOS UNIDOS (US) — Keywords que competidores rankean y galeriaduquearango.com no",
    fmt_subheader)

gaps_us = [
    ("wifredo lam", 5400, "Nombre de Artista", "cernudaarte.com", "pos 61", "Alta",
     "URGENTE: Crear página de artista completa en inglés para Wifredo Lam — artista del roster sin visibilidad US"),
    ("lam artist", 2900, "Nombre de Artista", "cernudaarte.com", "pos 15", "Alta",
     "Incluir como variante en la página de Wifredo Lam en inglés con FAQs y contexto histórico"),
    ("description of art gallery", 2900, "Editorial", "beatrizesguerra-art.com", "pos 22", "Alta",
     "Crear artículo en inglés 'What is an Art Gallery? The Role of Galleries in the Art Market' con schema FAQ"),
    ("colombian artists", 1300, "Editorial", "beatrizesguerra-art.com", "pos 23", "Media",
     "Ampliar artículo existente 'Colombian Painters' con más artistas del roster y datos citables"),
    ("pedro ruiz", 1600, "Nombre de Artista", "beatrizesguerra-art.com", "pos 8", "Media",
     "Crear perfil de artista Pedro Ruiz en inglés con obras y bibliografía"),
    ("define gallery", 1600, "Editorial", "beatrizesguerra-art.com", "pos 13", "Media",
     "Crear contenido educativo en inglés 'Defining an Art Gallery: Purpose, Role and History'"),
    ("amelia pelaez", 1000, "Nombre de Artista", "cernudaarte.com", "pos 49", "Media",
     "Artículo editorial 'Amelia Peláez and Cuban Modernism' — contexto para posicionar arte caribeño"),
    ("wifredo lam paintings", 590, "Nombre de Artista", "cernudaarte.com", "pos 37", "Baja",
     "Sección de obras en la página de Wifredo Lam con imágenes y fichas técnicas"),
    ("colombian painters", 720, "Editorial", "beatrizesguerra-art.com", "pos 8", "Baja",
     "Actualizar y ampliar el artículo 'Colombian Painters' ya existente con más artistas del roster"),
    ("artists of colombia", 590, "Editorial", "beatrizesguerra-art.com", "pos 34", "Baja",
     "Crear artículo 'Artists of Colombia: A Comprehensive Guide' con todos los artistas representados"),
    ("latin american art gallery miami", 480, "Comercial", "ascasogallery.com", "pos varios", "Media",
     "Crear landing page específica para el mercado Miami/US con enfoque en arte latinoamericano"),
    ("colombian art", 590, "Editorial", "beatrizesguerra-art.com", "pos 52", "Baja",
     "Hub de contenido 'Colombian Art' que enlace artículos existentes sobre artistas colombianos"),
    ("what is a gallery", 390, "Editorial", "beatrizesguerra-art.com", "pos 13", "Baja",
     "Artículo introductorio en inglés que sirva de entrada al blog para nuevos coleccionistas"),
    ("famous colombian artists", 320, "Editorial", "beatrizesguerra-art.com", "pos 55", "Baja",
     "Optimizar artículo existente 'Colombian Painters' añadiendo sección 'Famous Colombian Artists'"),
    ("cuban painter", 390, "Editorial", "cernudaarte.com", "pos 19", "Baja",
     "Artículo sobre pintores cubanos modernos y la relación con el arte latinoamericano representado por la galería"),
]

for r, row in enumerate(gaps_us, start=us_row + 1):
    kw, vol, cat, comp, pos, nivel, accion = row
    ws3.write(r, 0, kw, fmt_cell)
    ws3.write(r, 1, vol, fmt_num)
    ws3.write(r, 2, cat, fmt_cell)
    ws3.write(r, 3, comp, fmt_cell)
    ws3.write(r, 4, pos, fmt_cell)
    nf = fmt_alta if nivel == "Alta" else (fmt_media if nivel == "Media" else fmt_baja)
    ws3.write(r, 5, nivel, nf)
    ws3.write(r, 6, accion, fmt_cell)

final_note_row = us_row + len(gaps_us) + 2
ws3.write(final_note_row, 0,
    "Leyenda: Alta = volumen >2.000 búsquedas/mes o artista del roster sin posición. "
    "Media = 500–2.000. Baja = <500. Categorías: Nombre de Artista / Marca / Editorial / Comercial.",
    fmt_note)

# ── TAB 4: Páginas Top de Competencia ────────────────────────────────────────
ws4 = wb.add_worksheet("Páginas Top Competencia")
ws4.set_zoom(85)
ws4.set_row(0, 30)
ws4.set_row(1, 45)

ws4.merge_range("A1:G1",
    f"PÁGINAS TOP DE COMPETENCIA — CO + US | Datos al {DATE}",
    fmt_title)

headers_pg2 = [
    "Competidor", "Mercado", "URL de Página",
    "Tráfico Estimado", "Keyword Principal",
    "Tipo de Contenido", "Brecha para Duque Arango"
]
cols_pg2 = [24, 8, 48, 15, 24, 18, 46]
for c, (h, w) in enumerate(zip(headers_pg2, cols_pg2)):
    ws4.write(1, c, h, fmt_header)
    ws4.set_column(c, c, w)

top_pages_data = [
    # CO - galeriaelmuseo.com
    ("galeriaelmuseo.com", "CO",
     "www.galeriaelmuseo.com/ (homepage)", 1465,
     "galeria el museo / galerias bogota",
     "Homepage / Directorio",
     "Optimizar homepage de la galería para keywords 'galerías bogotá' y 'galería de arte Bogotá'"),
    ("galeriaelmuseo.com", "CO",
     "www.galeriaelmuseo.com/archives/43710/ (Jesús Abad Colorado)", 663,
     "jesus abad colorado",
     "Perfil de Artista / Noticia",
     "Crear artículo editorial sobre Jesús Abad Colorado — nombre de artista con 2.400 vol CO"),
    ("galeriaelmuseo.com", "CO",
     "www.galeriaelmuseo.com/archives/2084/ (Carlos Rojas)", 503,
     "carlos rojas",
     "Perfil de Artista",
     "Crear contenido editorial sobre Carlos Rojas y el arte cinético colombiano — 1.000 vol CO"),
    ("galeriaelmuseo.com", "CO",
     "www.galeriaelmuseo.com/archives/450/ (Luis Caballero)", 303,
     "luis caballero",
     "Perfil de Artista",
     "La galería representa a Luis Caballero — reforzar su página en español y en inglés"),
    ("galeriaelmuseo.com", "CO",
     "www.galeriaelmuseo.com/archives/32269/ (Autorretrato)", 204,
     "autorretrato / autoretrato que es",
     "Contenido Editorial",
     "Crear artículo 'El autorretrato en el arte colombiano' — keyword 2.400 vol CO sin posición"),
    # CO - galerialacometa.com
    ("galerialacometa.com", "CO",
     "galerialacometa.com/ (homepage)", 1868,
     "galeria la cometa / art gallery",
     "Homepage / Directorio",
     "Benchmark: galerialacometa rankea #2 para 'art gallery' en CO — optimizar versión inglesa"),
    ("galerialacometa.com", "CO",
     "galerialacometa.com/exhibiciones/bogota/gabriela-pinilla-clandestina-es", 139,
     "la chiqui del m 19",
     "Página de Exposición",
     "Crear contenido de cobertura de exposiciones del roster — aumenta keywords de long tail"),
    ("galerialacometa.com", "CO",
     "galerialacometa.com/exhibiciones/bogota/emma-reyes-las-caras-de-emma-reyes-es", 82,
     "emma reyes",
     "Página de Exposición / Artista",
     "Crear artículo editorial sobre Emma Reyes — keyword 1.900 vol CO, competidor rankea #4"),
    # CO - casasriegner.com
    ("casasriegner.com", "CO",
     "www.casasriegner.com/ (homepage)", 983,
     "galeria casa riegner / galerias bogotá",
     "Homepage",
     "Referencia: casasriegner rankea #1 para 'galería de arte bogotá' — crear landing de Bogotá"),
    ("casasriegner.com", "CO",
     "www.casasriegner.com/artistas/beatriz-gonzalez", 180,
     "beatriz gonzalez / beatriz gonzález artista",
     "Perfil de Artista",
     "Beatriz González: 3.600 búsquedas/mes — crear artículo editorial para capturar tráfico"),
    # US - operagallery.com
    ("operagallery.com", "US",
     "www.operagallery.com/ (homepage)", 1528,
     "opera gallery / international art gallery",
     "Homepage Internacional",
     "Referencia competitiva de galería internacional — benchmark de estructura y contenido en inglés"),
    ("operagallery.com", "US",
     "www.operagallery.com/artist/marc-chagall", 917,
     "marc chagall",
     "Perfil de Artista Internacional",
     "Opera rankea para artistas internacionales del roster — reforzar páginas en inglés de nuestros maestros"),
    ("operagallery.com", "US",
     "www.operagallery.com/viewing-rooms/fernando-botero", 491,
     "fernando botero / botero art",
     "Viewing Room — Artista del Roster",
     "⚠️ AMENAZA DIRECTA: Opera Gallery genera 491 visitas US con una página de Botero. "
     "Nuestra página en inglés debe superar esa profundidad de contenido."),
    # US - miguelabreugallery.com
    ("miguelabreugallery.com", "US",
     "miguelabreugallery.com/ (homepage)", 2402,
     "miguel abreu gallery / miguel gallery",
     "Homepage",
     "74% del tráfico concentrado en homepage — modelo de branding enfocado en el director"),
    ("miguelabreugallery.com", "US",
     "miguelabreugallery.com/artists/tishan-hsu/", 60,
     "tishan hsu",
     "Perfil de Artista Contemporáneo",
     "Referencia de perfil de artista contemporáneo en inglés — aplicar estructura similar al roster"),
    # US - forumgallery.com
    ("forumgallery.com", "US",
     "www.forumgallery.com/ (homepage)", 895,
     "forum gallery nyc",
     "Homepage",
     "Galería NY con 1.554 keywords US — analizar su arquitectura de contenido para artistas"),
    ("forumgallery.com", "US",
     "www.forumgallery.com/artists/ernie-barnes", 132,
     "ernie barnes",
     "Perfil de Artista",
     "Referencia de perfil de artista con tráfico sostenido — aplicar estructura a artistas contemporáneos del roster"),
    ("forumgallery.com", "US",
     "www.forumgallery.com/artists/claudio-bravo/videos", 114,
     "claudio bravo",
     "Perfil de Artista + Video",
     "Contenido multimedia de artista genera tráfico — integrar videos y multimedia en páginas de artistas"),
]

for r, row in enumerate(top_pages_data, start=2):
    comp, mkt, url, tr, kw, tipo, brecha = row
    ws4.write(r, 0, comp, fmt_cell)
    ws4.write(r, 1, mkt, fmt_cell)
    ws4.write(r, 2, url, fmt_url)
    ws4.write(r, 3, tr, fmt_num)
    ws4.write(r, 4, kw, fmt_cell)
    ws4.write(r, 5, tipo, fmt_cell)
    # Highlight Opera Gallery Botero threat
    if "AMENAZA DIRECTA" in brecha:
        fmt_brecha = fmt_alta
    else:
        fmt_brecha = fmt_cell
    ws4.write(r, 6, brecha, fmt_brecha)

# ── TAB 5: Insights y Recomendaciones ────────────────────────────────────────
ws5 = wb.add_worksheet("Insights y Recomendaciones")
ws5.set_zoom(85)
ws5.set_row(0, 30)
ws5.set_row(1, 45)

ws5.merge_range("A1:E1",
    f"INSIGHTS Y RECOMENDACIONES ESTRATÉGICAS | Datos al {DATE}",
    fmt_title)

headers_ins = [
    "Insight", "Evidencia", "Framework Aplicado",
    "Acción Recomendada", "Prioridad"
]
cols_ins = [38, 38, 24, 56, 10]
for c, (h, w) in enumerate(zip(headers_ins, cols_ins)):
    ws5.write(1, c, h, fmt_header)
    ws5.set_column(c, c, w)

insights = [
    (
        "Liderazgo absoluto en CO: galeriaduquearango.com es #1 del sector con 5× más tráfico que el segundo lugar",
        "23.040 visitas/mes CO vs galeriaelmuseo.com con 4.297 (2× ranking). "
        "La galería domina keywords de artistas clave: edgar negret (#2, 90.500 vol), david manzur (#2, 4.400 vol), alejandro obregón (#1).",
        "SEO Audit — Content Quality Assessment (E-E-A-T: Experience & Expertise)",
        "Mantener ritmo de publicación de perfiles de artistas. "
        "Priorizar artistas con volumen medio sin posición: Julio Larraz, Carlos Vega, Nadín Ospina. "
        "Cada perfil debe incluir: obra, biografía, bibliografía, FAQs y schema de artista.",
        "Alta"
    ),
    (
        "⚠️ AMENAZA CRÍTICA: Opera Gallery rankea para Fernando Botero en US con una página dedicada",
        "operagallery.com/viewing-rooms/fernando-botero genera 491 visitas US. "
        "'Fernando botero' tiene 14.800 búsquedas/mes US y 'fernando botero paintings' 4.400/mes. "
        "La galería rankea pos 7 para 'fernando botero' US pero con un artículo de blog, no una página de artista.",
        "competitor-alternatives — Research Process: Content Depth Over Surface + SEO Audit On-Page Optimization",
        "URGENTE (esta semana): Crear o potenciar la página en inglés '/en/artist/fernando-botero/' con: "
        "obras disponibles, historial de exposiciones, provenance, FAQs (schema FAQPage), estadísticas de mercado. "
        "Profundidad mínima: superar el viewing room de Opera Gallery. "
        "Target keywords: 'fernando botero paintings for sale', 'botero artworks', 'buy botero art'.",
        "Alta"
    ),
    (
        "⚠️ AMENAZA CRÍTICA: Wifredo Lam — artista del roster sin ninguna visibilidad en US",
        "'wifredo lam' tiene 5.400 búsquedas/mes en US. cernudaarte.com rankea pos 61. "
        "galeriaduquearango.com no aparece en top 100. "
        "'lam artist' (2.900 vol US) — cernudaarte rankea pos 15, galería sin posición.",
        "AI SEO — Pillar 1: Structure (make it extractable) + content-strategy Keyword Research: Awareness Stage",
        "Crear página en inglés '/en/artist/wifredo-lam/' con estructura AI-extractable: "
        "definición del artista en primer párrafo, bloque de estadísticas de mercado, tabla comparativa de períodos, "
        "FAQs con schema, fecha de actualización visible. "
        "Citar fuentes externas (Christie's, Sotheby's) para +40% de visibilidad AI (Princeton GEO study).",
        "Alta"
    ),
    (
        "Brecha comercial: keywords 'galerías bogotá' y 'galería de arte' sin posición en CO",
        "'galerias bogotá' (2.900 vol CO): galeriaelmuseo pos 12, galerialacometa pos 20, casasriegner pos 21. "
        "'galería de arte' (1.300 vol): casasriegner pos 6, galerialacometa pos 2. "
        "La galería tiene presencia en Bogotá pero no captura estas búsquedas comerciales.",
        "SEO Audit — Keyword Targeting: Commercial Intent + content-strategy Content-Market Fit",
        "Crear página de presencia Bogotá con URL '/galeriaduquearango-bogota/' o similar. "
        "Incluir: dirección Bogotá, artistas representados, horarios, enlace a exposiciones actuales. "
        "Revisar meta title y H1 del homepage para incluir 'Galería de arte en Medellín y Bogotá'.",
        "Alta"
    ),
    (
        "La galería no rankea en US para artistas clave del roster con páginas solo en español",
        "Edgar Negret: pos 2 CO para 'edgar negret' (90.500 vol) pero sin posición US. "
        "David Manzur: pos 2 CO (4.400 vol) sin página en inglés que rankee. "
        "Alejandro Obregón: pos 1 CO en múltiples variantes, sin presencia US.",
        "SEO Audit — International SEO: Hreflang + AI SEO Pillar 3: Presence",
        "Priorizar creación de páginas en inglés para los 3 artistas de mayor tráfico CO: "
        "Edgar Negret, David Manzur, Alejandro Obregón. "
        "Verificar configuración hreflang ES↔EN. "
        "Añadir 'Last updated' visible y author attribution para señales de frescura (AI visibility +25%).",
        "Alta"
    ),
    (
        "Beatriz González y Emma Reyes: artistas sin representación capturan búsquedas del sector en CO",
        "Beatriz González: 3.600 búsquedas/mes CO, casasriegner pos 3. "
        "Emma Reyes: 1.900 búsquedas/mes CO, galerialacometa pos 4. "
        "Feliza Bursztyn: 1.900 vol CO, galerialacometa pos 32. "
        "La galería no tiene contenido editorial sobre estas artistas icónicas colombianas.",
        "content-strategy — Competitor Analysis: Topics competitors rank for + AI SEO Pillar 2: Authority",
        "Crear serie editorial 'Maestros del Arte Colombiano Moderno' con artículos sobre Beatriz González, "
        "Emma Reyes, Feliza Bursztyn, Jorge Elías Triana. "
        "Estos artículos posicionan a la galería como referencia editorial aunque no los represente, "
        "y generan autoridad de dominio que transfiere a páginas de artistas del roster.",
        "Media"
    ),
    (
        "'Expresionismo' (5.400 vol CO): la galería ya está en posición 10 — quick win para subir al top 5",
        "galeriaduquearango.com/blog/el-expresionismo-en-latinoamerica-y-5-exponentes/ rankea pos 10 "
        "para 'expresionismo' (5.400 vol CO). "
        "Posición 10 = última página 1 o primera página 2 según el dispositivo. "
        "Pequeñas mejoras pueden generar +50–100% de tráfico adicional.",
        "SEO Audit — On-Page Optimization: Content Depth + AI SEO Pillar 1: Content Extractability",
        "Quick win: ampliar artículo con: tabla comparativa de movimientos expresionistas, "
        "más ejemplos de artistas del roster, estadísticas del mercado del arte expresionista, "
        "schema FAQ y schema Article con author attribution. "
        "Meta: subir de posición 10 a posiciones 3–6 para capturar ~400–800 visitas adicionales/mes.",
        "Media"
    ),
    (
        "Opera Gallery (US) lidera el mercado americano con 7.402 visitas orgánicas — modelo a estudiar",
        "operagallery.com genera 7.402 visitas/mes en US vs galeriaduquearango.com 2.357. "
        "Sus top páginas son: artista Marc Chagall (917 visits), Simon Hantaï (665), Keith Haring (598), "
        "Fernando Botero viewing room (491). "
        "Modelo: páginas profundas por artista con 100+ keywords cada una.",
        "competitor-alternatives — Research Process: Deep Competitor Research + content-strategy Hub and Spoke",
        "Benchmarking de estructura Opera Gallery: crear 'viewing rooms' o 'colecciones disponibles' "
        "en inglés por artista con: obras disponibles, precio estimado (o 'solicitar información'), "
        "historial de subastas. "
        "Prioridad: Fernando Botero, Alejandro Obregón, Ana Mercedes Hoyos, Rufino Tamayo, Wifredo Lam.",
        "Media"
    ),
]

for r, row in enumerate(insights, start=2):
    insight, evidencia, framework, accion, prioridad = row
    ws5.write(r, 0, insight, fmt_cell)
    ws5.write(r, 1, evidencia, fmt_cell)
    ws5.write(r, 2, framework, fmt_cell)
    ws5.write(r, 3, accion, fmt_cell)
    nf = fmt_alta if prioridad == "Alta" else (fmt_media if prioridad == "Media" else fmt_baja)
    ws5.write(r, 4, prioridad, nf)

ws5.write(len(insights) + 3, 0,
    f"Reporte generado por Rutina de Inteligencia SEO — Galería Duque Arango. "
    f"Fuente: Semrush (CO + US). Fecha de datos: {DATE}. "
    "Frameworks aplicados: seo-audit v1.2, ai-seo v1.2, content-strategy v1.1, competitor-alternatives v1.1.",
    fmt_note)

wb.close()

size = os.path.getsize(OUTPUT_PATH)
print(f"Archivo generado: {OUTPUT_PATH}")
print(f"Tamaño: {size:,} bytes")
