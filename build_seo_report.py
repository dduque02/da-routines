#!/usr/bin/env python3
"""Genera el reporte Excel de Inteligencia SEO Semanal — Galería Duque Arango."""

import xlsxwriter
import os

DATE = "2026-06-12"
FILENAME = f"Inteligencia_SEO_{DATE}.xlsx"
FILEPATH = os.path.join("/home/user/da-routines", FILENAME)

wb = xlsxwriter.Workbook(FILEPATH)

# ── Formatos ──────────────────────────────────────────────────────────────────
def fmt(wb, opts):
    return wb.add_format(dict({'border': 1, 'valign': 'vcenter', 'text_wrap': True}, **opts))

H  = fmt(wb, {'bold': True, 'bg_color': '#1A237E', 'font_color': '#FFFFFF', 'align': 'center', 'font_size': 10})
SH = fmt(wb, {'bold': True, 'bg_color': '#283593', 'font_color': '#FFD54F', 'align': 'center', 'font_size': 9})
N  = fmt(wb, {'font_size': 9})
NA = fmt(wb, {'bg_color': '#F5F5F5', 'font_size': 9})
NB = fmt(wb, {'bold': True, 'font_size': 9})
NUM  = fmt(wb, {'num_format': '#,##0', 'font_size': 9, 'align': 'right'})
NUMA = fmt(wb, {'num_format': '#,##0', 'bg_color': '#F5F5F5', 'font_size': 9, 'align': 'right'})
GOLD = fmt(wb, {'bold': True, 'bg_color': '#FFF9C4', 'font_size': 9})
GOLDC= fmt(wb, {'bold': True, 'bg_color': '#FFF9C4', 'num_format': '#,##0', 'font_size': 9, 'align': 'right'})
ALTA  = fmt(wb, {'bold': True, 'bg_color': '#C62828', 'font_color': '#FFFFFF', 'align': 'center', 'font_size': 9})
MEDIA = fmt(wb, {'bold': True, 'bg_color': '#F57F17', 'font_color': '#FFFFFF', 'align': 'center', 'font_size': 9})
BAJA  = fmt(wb, {'bold': True, 'bg_color': '#2E7D32', 'font_color': '#FFFFFF', 'align': 'center', 'font_size': 9})
WARN  = fmt(wb, {'bold': True, 'bg_color': '#B71C1C', 'font_color': '#FFFFFF', 'font_size': 9})
TITLE = wb.add_format({'bold': True, 'font_size': 14, 'font_color': '#1A237E'})
INFO  = wb.add_format({'italic': True, 'font_color': '#616161', 'font_size': 9})

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — DESEMPEÑO PROPIO
# ══════════════════════════════════════════════════════════════════════════════
ws1 = wb.add_worksheet("Desempeño Propio")
ws1.set_zoom(85)
ws1.set_column(0, 0, 20)
ws1.set_column(1, 1, 14)
ws1.set_column(2, 2, 16)
ws1.set_column(3, 3, 36)

ws1.write(0, 0, f"INTELIGENCIA SEO — GALERÍA DUQUE ARANGO", TITLE)
ws1.write(1, 0, f"Semana: {DATE}  |  Datos al: junio 2026  |  SEMANA BASE (sin histórico previo)", INFO)
ws1.write(2, 0, "Dominio analizado: galeriaduquearango.com", INFO)

# Métricas generales
ws1.write(4, 0, "BASE DE DATOS", H)
ws1.write(4, 1, "MÉTRICA", H)
ws1.write(4, 2, "VALOR", H)
ws1.write(4, 3, "NOTAS", H)

metricas = [
    ("Colombia (CO)", "Keywords orgánicas", "2.335", "Semrush DB: co"),
    ("Colombia (CO)", "Tráfico orgánico estimado", "19.511", "Visitas/mes estimadas"),
    ("Colombia (CO)", "Semrush Rank CO", "5.314", "Posición global en base CO"),
    ("Colombia (CO)", "Valor tráfico orgánico", "$530", "Costo equivalente en paid"),
    ("Colombia (CO)", "Keywords en top 3", "12", "Estimado de keywords posición ≤ 3"),
    ("Colombia (CO)", "Keywords en top 10", "19+", "Estimado de keywords posición ≤ 10"),
    ("Estados Unidos (US)", "Keywords orgánicas", "1.946", "Semrush DB: us"),
    ("Estados Unidos (US)", "Tráfico orgánico estimado", "2.646", "Visitas/mes estimadas"),
    ("Estados Unidos (US)", "Semrush Rank US", "526.874", "Posición global en base US"),
    ("Estados Unidos (US)", "Valor tráfico orgánico", "$758", "Costo equivalente en paid"),
    ("Estados Unidos (US)", "Keywords en top 3", "3", "Posición ≤ 3 en US"),
    ("Estados Unidos (US)", "Keywords en top 10", "14+", "Posición ≤ 10 en US"),
]
for i, (db, met, val, nota) in enumerate(metricas):
    r = i + 5
    f1, f2, f3, f4 = (NA, NA, NUMA, NA) if i % 2 else (N, N, NUM, N)
    ws1.write(r, 0, db, f1)
    ws1.write(r, 1, met, f2)
    ws1.write(r, 2, val, f3)
    ws1.write(r, 3, nota, f4)

# Top 20 keywords CO
r = 19
ws1.write(r, 0, "TOP 20 KEYWORDS ORGÁNICAS — COLOMBIA (CO)", SH)
ws1.merge_range(r, 0, r, 3, "TOP 20 KEYWORDS ORGÁNICAS — COLOMBIA (CO)", SH)
r += 1
ws1.write(r, 0, "Keyword", H)
ws1.write(r, 1, "Posición", H)
ws1.write(r, 2, "Volumen Búsqueda", H)
ws1.write(r, 3, "URL que Rankea", H)
ws1.set_column(3, 3, 55)
r += 1

kw_co = [
    ("edgar negret", 2, 90500, "galeriaduquearango.com/artista/edgar-negret/"),
    ("plaza botero medellín antioquia", 8, 22200, "galeriaduquearango.com/blog/plaza-botero-resignificar-el-espacio-publico-a-traves-del-arte-en-medellin/"),
    ("galeria duque arango", 1, 720, "galeriaduquearango.com/"),
    ("escultura", 3, 5400, "galeriaduquearango.com/blog/escultura-en-colombia-y-sus-representantes/"),
    ("plaza botero", 8, 49500, "galeriaduquearango.com/blog/plaza-botero-resignificar-el-espacio-publico-a-traves-del-arte-en-medellin/"),
    ("obregón pinturas", 1, 480, "galeriaduquearango.com/artista/alejandro-obregon/"),
    ("omar rayo", 6, 5400, "galeriaduquearango.com/blog/omar-rayo-una-historia-narrada-en-geometria/"),
    ("obras de alejandro obregón", 1, 320, "galeriaduquearango.com/artista/alejandro-obregon/"),
    ("alejandro obregón obras", 1, 320, "galeriaduquearango.com/artista/alejandro-obregon/"),
    ("david manzur", 2, 4400, "galeriaduquearango.com/artista/david-manzur/"),
    ("enrique grau obras", 1, 260, "galeriaduquearango.com/artista/enrique-grau/"),
    ("guayasamin", 3, 2400, "galeriaduquearango.com/blog/el-legado-al-arte-de-oswaldo-guayasamin/"),
    ("enrique grau", 2, 1600, "galeriaduquearango.com/artista/enrique-grau/"),
    ("pinturas de alejandro obregón", 1, 210, "galeriaduquearango.com/artista/alejandro-obregon/"),
    ("alejandro obregon obras", 1, 210, "galeriaduquearango.com/artista/alejandro-obregon/"),
    ("parque botero", 8, 3600, "galeriaduquearango.com/blog/plaza-botero-resignificar-el-espacio-publico-a-traves-del-arte-en-medellin/"),
    ("obras de enrique grau", 1, 170, "galeriaduquearango.com/artista/enrique-grau/"),
    ("obras de arte de david manzur", 1, 170, "galeriaduquearango.com/artista/david-manzur/"),
    ("botero cuadros", 1, 170, "galeriaduquearango.com/blog/las-obras-mejor-vendidas-de-fernando-botero/"),
    ("edgar negret obras", 1, 140, "galeriaduquearango.com/artista/edgar-negret/"),
]
for i, (kw, pos, vol, url) in enumerate(kw_co):
    f1, f2 = (NA, NUMA) if i % 2 else (N, NUM)
    ws1.write(r + i, 0, kw, f1)
    ws1.write(r + i, 1, pos, f2)
    ws1.write(r + i, 2, vol, f2)
    ws1.write(r + i, 3, url, f1)

# Top 20 keywords US
r += len(kw_co) + 2
ws1.merge_range(r, 0, r, 3, "TOP 20 KEYWORDS ORGÁNICAS — ESTADOS UNIDOS (US)", SH)
r += 1
ws1.write(r, 0, "Keyword", H)
ws1.write(r, 1, "Posición", H)
ws1.write(r, 2, "Volumen Búsqueda", H)
ws1.write(r, 3, "URL que Rankea", H)
r += 1

kw_us = [
    ("colombian painters", 3, 720, "galeriaduquearango.com/en/blog/discovering-colombian-art-from-fernando-botero-to-contemporary-masters/"),
    ("plaza botero medellín colombia", 6, 1900, "galeriaduquearango.com/en/blog/plaza-plaza-botero-resignifying-public-space-through-art-in-medellin/"),
    ("galeria de arte", 2, 480, "galeriaduquearango.com/"),
    ("oswaldo guayasamin", 7, 1600, "galeriaduquearango.com/en/blog/oswaldo-guayasamins-legacy-to-art/"),
    ("botero's", 6, 3600, "galeriaduquearango.com/en/blog/all-you-need-to-know-to-understand-fernando-boteros-artworks/"),
    ("guayasamin", 5, 1300, "galeriaduquearango.com/en/blog/the-works-of-oswaldo-guayasamin/"),
    ("fernando botero paintings", 8, 4400, "galeriaduquearango.com/en/blog/all-you-need-to-know-to-understand-fernando-boteros-artworks/"),
    ("botero", 9, 9900, "galeriaduquearango.com/en/blog/all-you-need-to-know-to-understand-fernando-boteros-artworks/"),
    ("botero paintings", 8, 3600, "galeriaduquearango.com/en/blog/all-you-need-to-know-to-understand-fernando-boteros-artworks/"),
    ("fernando botero artworks", 9, 1300, "galeriaduquearango.com/en/blog/all-you-need-to-know-to-understand-fernando-boteros-artworks/"),
    ("fernando botero paintings", 9, 4400, "galeriaduquearango.com/en/blog/all-you-need-to-know-to-understand-fernando-boteros-artworks/"),
    ("fernando botero style", 1, 170, "galeriaduquearango.com/en/blog/unmistakable-brand-boterism/"),
    ("boterismo", 3, 480, "galeriaduquearango.com/en/blog/unmistakable-brand-boterism/"),
    ("picasso realistic art", 4, 480, "galeriaduquearango.com/en/blog/artistic-periods-of-pablo-picasso/"),
    ("fernando botero art", 10, 3600, "galeriaduquearango.com/en/blog/all-you-need-to-know-to-understand-fernando-boteros-artworks/"),
    ("botero artist", 12, 3600, "galeriaduquearango.com/en/blog/all-you-need-to-know-to-understand-fernando-boteros-artworks/"),
    ("ana mercedes hoyos", 7, 3600, "galeriaduquearango.com/artista/ana-mercedes-hoyos/"),
    ("boterism", 1, 90, "galeriaduquearango.com/en/blog/unmistakable-brand-boterism/"),
    ("fernando botero", 22, 18100, "galeriaduquearango.com/en/blog/all-you-need-to-know-to-understand-fernando-boteros-artworks/"),
    ("nombre de las 30 pinturas más vistas del fernando botero", 2, 210, "galeriaduquearango.com/blog/las-obras-mejor-vendidas-de-fernando-botero/"),
]
for i, (kw, pos, vol, url) in enumerate(kw_us):
    f1, f2 = (NA, NUMA) if i % 2 else (N, NUM)
    ws1.write(r + i, 0, kw, f1)
    ws1.write(r + i, 1, pos, f2)
    ws1.write(r + i, 2, vol, f2)
    ws1.write(r + i, 3, url, f1)

# Top 10 páginas CO
r += len(kw_us) + 2
ws1.merge_range(r, 0, r, 3, "TOP 10 PÁGINAS POR TRÁFICO — COLOMBIA (CO)", SH)
r += 1
ws1.write(r, 0, "URL de Página", H)
ws1.write(r, 1, "Keywords", H)
ws1.write(r, 2, "Tráfico Estimado", H)
ws1.write(r, 3, "% del Tráfico Total CO", H)
r += 1

pages_co = [
    ("galeriaduquearango.com/artista/edgar-negret/", 30, 8786, "45.03%"),
    ("galeriaduquearango.com/blog/las-obras-de-fernando-botero-y-su-significado/", 143, 1493, "7.65%"),
    ("galeriaduquearango.com/artista/alejandro-obregon/", 73, 1181, "6.05%"),
    ("galeriaduquearango.com/blog/plaza-botero-resignificar-el-espacio-publico-a-traves-del-arte-en-medellin/", 34, 1137, "5.82%"),
    ("galeriaduquearango.com/artista/david-manzur/", 54, 1072, "5.49%"),
    ("galeriaduquearango.com/", 117, 819, "4.19%"),
    ("galeriaduquearango.com/blog/las-obras-mejor-vendidas-de-fernando-botero/", 91, 463, "2.37%"),
    ("galeriaduquearango.com/artista/enrique-grau/", 21, 419, "2.14%"),
    ("galeriaduquearango.com/blog/escultura-en-colombia-y-sus-representantes/", 57, 387, "1.98%"),
    ("galeriaduquearango.com/artista/ana-mercedes-hoyos/", 20, 385, "1.97%"),
]
for i, (url, kws, traf, pct) in enumerate(pages_co):
    f1, f2 = (NA, NUMA) if i % 2 else (N, NUM)
    ws1.write(r + i, 0, url, f1)
    ws1.write(r + i, 1, kws, f2)
    ws1.write(r + i, 2, traf, f2)
    ws1.write(r + i, 3, pct, f1)

# Top 10 páginas US
r += len(pages_co) + 2
ws1.merge_range(r, 0, r, 3, "TOP 10 PÁGINAS POR TRÁFICO — ESTADOS UNIDOS (US)", SH)
r += 1
ws1.write(r, 0, "URL de Página", H)
ws1.write(r, 1, "Keywords", H)
ws1.write(r, 2, "Tráfico Estimado", H)
ws1.write(r, 3, "% del Tráfico Total US", H)
r += 1

pages_us = [
    ("galeriaduquearango.com/artista/ana-mercedes-hoyos/", 2, 493, "18.63%"),
    ("galeriaduquearango.com/en/blog/the-works-of-fernando-botero-and-their-significance/", 69, 396, "14.96%"),
    ("galeriaduquearango.com/en/blog/all-you-need-to-know-to-understand-fernando-boteros-artworks/", 109, 296, "11.18%"),
    ("galeriaduquearango.com/en/blog/discovering-colombian-art-from-fernando-botero-to-contemporary-masters/", 72, 156, "5.89%"),
    ("galeriaduquearango.com/en/blog/exploring-fernando-botero-most-famous-works/", 21, 148, "5.59%"),
    ("galeriaduquearango.com/en/blog/unmistakable-brand-boterism/", 82, 145, "5.47%"),
    ("galeriaduquearango.com/en/blog/artistic-periods-of-pablo-picasso/", 150, 104, "3.93%"),
    ("galeriaduquearango.com/en/blog/the-works-of-oswaldo-guayasamin/", 40, 99, "3.74%"),
    ("galeriaduquearango.com/en/artist/luis-caballero/", 13, 88, "3.32%"),
    ("galeriaduquearango.com/blog/plaza-botero-resignificar-el-espacio-publico-a-traves-del-arte-en-medellin/", 12, 82, "3.09%"),
]
for i, (url, kws, traf, pct) in enumerate(pages_us):
    f1, f2 = (NA, NUMA) if i % 2 else (N, NUM)
    ws1.write(r + i, 0, url, f1)
    ws1.write(r + i, 1, kws, f2)
    ws1.write(r + i, 2, traf, f2)
    ws1.write(r + i, 3, pct, f1)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — BENCHMARK COMPETITIVO
# ══════════════════════════════════════════════════════════════════════════════
ws2 = wb.add_worksheet("Benchmark Competitivo")
ws2.set_zoom(85)
ws2.write(0, 0, "BENCHMARK COMPETITIVO — GALERÍA DUQUE ARANGO VS 16 COMPETIDORES", TITLE)
ws2.write(1, 0, f"Datos al: junio 2026  |  Ordenado por Tráfico CO descendente  |  DB: Semrush CO y US", INFO)

cols = ["Dominio", "Tier", "Tráfico CO", "Keywords CO", "Tráfico US", "Keywords US", "Posición CO de 17", "Posición US de 17"]
widths = [32, 12, 14, 14, 14, 14, 18, 18]
for c, (col, w) in enumerate(zip(cols, widths)):
    ws2.write(3, c, col, H)
    ws2.set_column(c, c, w)

competitors = [
    # (dominio, tier, trafico_co, kw_co, trafico_us, kw_us, pos_co, pos_us, es_duque)
    ("galeriaduquearango.com",   "—",    19511, 2335, 2646, 1946, 1, 3, True),
    ("galeriaelmuseo.com",       "Tier 1", 4038, 547,  59,   44,  2, 13, False),
    ("galerialacometa.com",      "Tier 1", 3326, 360,  526,  132, 3, 7,  False),
    ("casasriegner.com",         "Tier 1", 1881, 261,  24,   55,  4, 11, False),
    ("sgr-art.com",              "Tier 1", 628,  62,   0,    5,   5, 14, False),
    ("galeriacasacuadrada.com",  "Tier 1", 514,  88,   1,    3,   6, 13, False),
    ("beatrizesguerra-art.com",  "Tier 1", 180,  39,   93,   85,  7, 10, False),
    ("galeriafreites.com",       "Tier 2", 166,  104,  2,    16,  8, 12, False),
    ("otros360grados.com",       "Tier 1", 122,  56,   0,    9,   9, 14, False),
    ("operagallery.com",         "Tier 2", 58,   41,   7958, 2994,10, 1,  False),
    ("galeriaelsapineres.art",   "Tier 1", 15,   25,   0,    0,   11, 14, False),
    ("artoftheworldgallery.com", "Tier 2", 13,   23,   1671, 780, 12, 5,  False),
    ("galeriafreites.com",       "Tier 2", 166,  104,  2,    16,  8, 12, False),
    ("latinartcore.com",         "Tier 2", 0,    0,    145,  113, 14, 9,  False),
    ("ascasogallery.com",        "Tier 2", 0,    10,   321,  160, 14, 8,  False),
    ("cernudaarte.com",          "Tier 2", 0,    6,    554,  398, 14, 6,  False),
    ("miguelabreugallery.com",   "Tier 2", 0,    2,    2815, 411, 14, 2,  False),
    ("forumgallery.com",         "Tier 2", 0,    5,    1806, 1586,14, 4,  False),
]

# Deduplicate galeriafreites
seen = set()
comps_clean = []
for row in competitors:
    if row[0] not in seen:
        seen.add(row[0])
        comps_clean.append(row)

# Re-sort by trafico_co descending
comps_clean.sort(key=lambda x: x[2], reverse=True)
# Put galeriaduquearango at top always
duque = [c for c in comps_clean if c[0] == "galeriaduquearango.com"][0]
rest  = [c for c in comps_clean if c[0] != "galeriaduquearango.com"]
comps_clean = [duque] + rest

for i, row in enumerate(comps_clean):
    r = i + 4
    dom, tier, tco, kwco, tus, kwus, pco, pus, is_duque = row
    rf = GOLD if is_duque else (NA if i % 2 else N)
    nf = GOLDC if is_duque else (NUMA if i % 2 else NUM)
    ws2.write(r, 0, dom, rf)
    ws2.write(r, 1, tier, rf)
    ws2.write(r, 2, tco if tco else "Sin datos", nf if tco else rf)
    ws2.write(r, 3, kwco if kwco else "Sin datos", nf if kwco else rf)
    ws2.write(r, 4, tus if tus else "Sin datos", nf if tus else rf)
    ws2.write(r, 5, kwus if kwus else "Sin datos", nf if kwus else rf)
    ws2.write(r, 6, pco, nf)
    ws2.write(r, 7, pus, nf)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — BRECHAS DE KEYWORDS
# ══════════════════════════════════════════════════════════════════════════════
ws3 = wb.add_worksheet("Brechas de Keywords")
ws3.set_zoom(85)
ws3.write(0, 0, "BRECHAS DE KEYWORDS — OPORTUNIDADES PARA GALERÍA DUQUE ARANGO", TITLE)
ws3.write(1, 0, f"Datos al: junio 2026  |  Keywords que competidores rankean y galeriaduquearango.com NO rankea", INFO)

cols3 = ["Keyword", "Volumen Búsqueda", "Mercado", "Categoría SEO", "Competidor que Rankea", "Posición Competidor", "En Roster Duque Arango", "Nivel de Oportunidad", "Acción de Contenido Sugerida"]
widths3 = [28, 16, 10, 22, 26, 18, 22, 18, 40]
for c, (col, w) in enumerate(zip(cols3, widths3)):
    ws3.write(3, c, col, H)
    ws3.set_column(c, c, w)

gaps = [
    # (keyword, vol, mercado, categoria, competidor, pos_comp, en_roster, nivel, accion)
    ("museo de arte",           8100,  "CO",  "Editorial",          "galeriaelmuseo.com",   13,  "No", "Alta",   "Artículo: 'Los mejores museos de arte en Colombia y las galerías que los complementan'"),
    ("beatriz gonzalez",        8100,  "CO",  "Nombre de Artista",  "casasriegner.com",      3,  "No", "Alta",   "Artículo editorial: historia del arte colombiano y maestras del siglo XX"),
    ("wifredo lam",             6600,  "US",  "Nombre de Artista",  "cernudaarte.com",      54,  "SÍ ⚠️","Alta",  "Página de artista en inglés: Wifredo Lam — vida, obras y mercado secundario"),
    ("galerias bogotá",         2900,  "CO",  "Comercial",          "galerialacometa.com / casasriegner.com / galeriaelmuseo.com", 11, "No", "Alta", "Página: 'Galerías de arte en Bogotá — guía de coleccionistas'"),
    ("wilfredo lam",            3600,  "US",  "Nombre de Artista",  "cernudaarte.com",      54,  "SÍ ⚠️","Alta",  "Capturar con la misma página de Wifredo Lam (incluir variante ortográfica)"),
    ("autorretrato",            2400,  "CO",  "Editorial",          "galeriaelmuseo.com",    4,  "No", "Alta",   "Artículo: 'El autorretrato en el arte colombiano: de Grau a Manzur'"),
    ("jesus abad colorado",     2400,  "CO",  "Nombre de Artista",  "galeriaelmuseo.com",    4,  "No", "Media",  "Artículo editorial sobre fotografía colombiana (no representado, pero relevante)"),
    ("emma reyes",              1900,  "CO",  "Nombre de Artista",  "galerialacometa.com",   4,  "No", "Media",  "Artículo sobre maestras colombianas: Emma Reyes, Beatriz González, Ana Mercedes Hoyos"),
    ("feliza bursztyn",         1900,  "CO",  "Nombre de Artista",  "galerialacometa.com",  32,  "No", "Media",  "Incluir en artículo de escultoras colombianas del siglo XX"),
    ("beatriz gonzalez artista",1600,  "CO",  "Nombre de Artista",  "casasriegner.com",      3,  "No", "Media",  "Artículo: 'Grandes artistas colombianas y su legado'"),
    ("amelia pelaez",           1000,  "US",  "Nombre de Artista",  "cernudaarte.com",       6,  "No", "Media",  "Artículo: 'Latin American Modern Masters — from Lam to Peláez'"),
    ("wifredo lam paintings",    590,  "US",  "Nombre de Artista",  "cernudaarte.com",      37,  "SÍ ⚠️","Media", "Incluir sección de obras en la página de artista Wifredo Lam"),
    ("tomas sanchez artist",     480,  "US",  "Nombre de Artista",  "cernudaarte.com",      10,  "No", "Media",  "Artículo: 'Contemporary Cuban and Latin American artists'"),
    ("bienal de arte medellin",  880,  "CO",  "Editorial",          "casasriegner.com",     16,  "No", "Baja",   "Artículo de contexto sobre el ecosistema del arte en Medellín"),
    ("jacanamijoy",              880,  "CO",  "Nombre de Artista",  "galeriaelmuseo.com",    7,  "No", "Baja",   "Mención en artículo sobre arte contemporáneo colombiano"),
    ("antonio caro",             720,  "CO",  "Nombre de Artista",  "casasriegner.com",      3,  "No", "Baja",   "Artículo histórico sobre arte conceptual colombiano"),
    ("carlos rojas",            1000,  "CO",  "Nombre de Artista",  "casasriegner.com",      2,  "No", "Baja",   "Mención en artículo sobre arte geométrico colombiano (cf. Omar Rayo)"),
]

for i, row in enumerate(gaps):
    r = i + 4
    kw, vol, mer, cat, comp, pos, roster, nivel, accion = row
    fa = NA if i % 2 else N
    fn = NUMA if i % 2 else NUM
    nf_nivel = ALTA if nivel == "Alta" else (MEDIA if nivel == "Media" else BAJA)
    warn = WARN if "SÍ ⚠️" in roster else fa
    ws3.write(r, 0, kw, fa)
    ws3.write(r, 1, vol, fn)
    ws3.write(r, 2, mer, fa)
    ws3.write(r, 3, cat, fa)
    ws3.write(r, 4, comp, fa)
    ws3.write(r, 5, pos, fn)
    ws3.write(r, 6, roster, warn)
    ws3.write(r, 7, nivel, nf_nivel)
    ws3.write(r, 8, accion, fa)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — PÁGINAS TOP DE COMPETENCIA
# ══════════════════════════════════════════════════════════════════════════════
ws4 = wb.add_worksheet("Páginas Top de Competencia")
ws4.set_zoom(85)
ws4.write(0, 0, "PÁGINAS TOP DE COMPETIDORES — COLOMBIA Y ESTADOS UNIDOS", TITLE)
ws4.write(1, 0, "Top 3 competidores por tráfico en CO y US  |  Datos al: junio 2026", INFO)

cols4 = ["Competidor", "Mercado", "URL de Página", "Tráfico Estimado", "Keywords", "Tipo de Contenido", "Brecha para Duque Arango"]
widths4 = [28, 10, 50, 15, 12, 22, 45]
for c, (col, w) in enumerate(zip(cols4, widths4)):
    ws4.write(3, c, col, H)
    ws4.set_column(c, c, w)

top_pages = [
    # CO
    ("galeriaelmuseo.com (CO #2)", "CO", "www.galeriaelmuseo.com/", 1425, 86, "Homepage / Portafolio", "Duque Arango rankea más alto que El Museo en CO pero puede fortalecer homepage con más keywords"),
    ("galeriaelmuseo.com (CO #2)", "CO", "www.galeriaelmuseo.com/archives/43710/", 555, 17, "Artículo editorial", "Identificar qué tema cubre — posible oportunidad de contenido similar"),
    ("galeriaelmuseo.com (CO #2)", "CO", "www.galeriaelmuseo.com/archives/2084/", 503, 3, "Página de artista/exhibit", "Autorretrato — crear artículo editorial homólogo"),
    ("galerialacometa.com (CO #3)", "CO", "galerialacometa.com/", 1845, 60, "Homepage", "La Cometa usa su homepage para capturar branded searches; Duque Arango tiene 819 vs 1,845"),
    ("galerialacometa.com (CO #3)", "CO", "galerialacometa.com/exhibiciones/bogota/gabriela-pinilla-clandestina-es", 182, 19, "Página de exposición", "Crear páginas individuales de exposición optimizadas — Duque Arango carece de estas páginas"),
    ("galerialacometa.com (CO #3)", "CO", "galerialacometa.com/artistas/carlos-castro-es", 95, 12, "Página de artista", "Modelo de página de artista contemporáneo bien rankeada"),
    ("casasriegner.com (CO #4)", "CO", "www.casasriegner.com/", 736, 64, "Homepage", "Beatriz González genera el 30% del tráfico total de Casa Riegner"),
    ("casasriegner.com (CO #4)", "CO", "www.casasriegner.com/artistas/beatriz-gonzalez", 558, 26, "Página de artista (Beatriz González)", "Beatriz González (8,100 búsquedas/mes) — amenaza de artista no en roster pero domina CO"),
    ("casasriegner.com (CO #4)", "CO", "www.casasriegner.com/artistas/carlos-rojas", 117, 11, "Página de artista (Carlos Rojas)", "Carlos Rojas (1,000 búsquedas/mes CO) — art colombiano geométrico, relacionable con Omar Rayo"),
    # US
    ("operagallery.com (US #1)", "US", "www.operagallery.com/", 1468, 109, "Homepage", "Opera Gallery domina US con 7,958 vs nuestras 2,646 visitas/mes"),
    ("operagallery.com (US #1)", "US", "www.operagallery.com/viewing-rooms/fernando-botero", 436, 88, "Viewing room — Fernando Botero ⚠️", "AMENAZA: Opera rankea para 88 keywords de Botero. Crear viewing room / galería virtual de Botero en inglés"),
    ("operagallery.com (US #1)", "US", "www.operagallery.com/artist/keith-haring", 659, 160, "Página de artista internacional", "Modelo: páginas de artista con muchas keywords — Duque Arango puede replicar para Plensa, Rondinone, Le Parc"),
    ("miguelabreugallery.com (US #2)", "US", "miguelabreugallery.com/", 1923, 17, "Homepage", "68% del tráfico US en homepage; Duque Arango homepage US tiene menor tracción"),
    ("miguelabreugallery.com (US #2)", "US", "miguelabreugallery.com/artists/quaytman/", 158, 15, "Página de artista contemporáneo", "Modelo de página de artista con good ranking en US"),
    ("artoftheworldgallery.com (US #5)", "US", "www.artoftheworldgallery.com/", 1333, 325, "Homepage", "Homepage capta el 80% del tráfico US — keywords muy diversas"),
    ("artoftheworldgallery.com (US #5)", "US", "www.artoftheworldgallery.com/blog/botero-blog/", 49, 32, "Blog sobre Botero ⚠️", "Competidor de blog en US rankea también para Botero — espacio muy disputado"),
]

for i, row in enumerate(top_pages):
    r = i + 4
    comp, mer, url, traf, kws, tipo, brecha = row
    fa = NA if i % 2 else N
    fn = NUMA if i % 2 else NUM
    ws4.write(r, 0, comp, fa)
    ws4.write(r, 1, mer, fa)
    ws4.write(r, 2, url, fa)
    ws4.write(r, 3, traf, fn)
    ws4.write(r, 4, kws, fn)
    ws4.write(r, 5, tipo, WARN if "⚠️" in tipo else fa)
    ws4.write(r, 6, brecha, fa)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — INSIGHTS Y RECOMENDACIONES
# ══════════════════════════════════════════════════════════════════════════════
ws5 = wb.add_worksheet("Insights y Recomendaciones")
ws5.set_zoom(85)
ws5.write(0, 0, "INSIGHTS Y RECOMENDACIONES ESTRATÉGICAS — SEMANA BASE", TITLE)
ws5.write(1, 0, "Cada insight cita el framework del skill que lo respalda  |  Datos al: junio 2026", INFO)

cols5 = ["Insight", "Evidencia", "Framework Aplicado", "Acción Recomendada", "Prioridad"]
widths5 = [32, 42, 28, 50, 10]
for c, (col, w) in enumerate(zip(cols5, widths5)):
    ws5.write(3, c, col, H)
    ws5.set_column(c, c, w)

insights = [
    (
        "AMENAZA: cernudaarte.com rankea para Wifredo Lam (roster) en US",
        "wifredo lam = 6,600 búsquedas/mes US. cernudaarte.com pos. 54. Duque Arango NO aparece. Wifredo Lam está explícitamente en el roster como Maestro Moderno (mercado secundario).",
        "competitor-alternatives (Research Process) + CLAUDE.md: 'When a competitor ranks for any artist in the roster, flag as high-priority threat'",
        "Crear/optimizar página de artista 'Wifredo Lam' en inglés con texto completo, obras representativas, biography y schema markup. Incluir variante 'wilfredo lam' (3,600/mes). Deadline: 2 semanas.",
        "Alta"
    ),
    (
        "AMENAZA: operagallery.com con viewing room de Fernando Botero en US (88 keywords)",
        "Opera Gallery tiene página /viewing-rooms/fernando-botero con 88 keywords y 436 visitas/mes US. Fernando Botero está en el roster de Duque Arango. Nosotros rankeamos para Botero en US solo desde el blog (posición 9+).",
        "competitor-alternatives (Format 4: Competitor vs Competitor) + seo-audit (On-Page SEO — Keyword Targeting)",
        "Crear viewing room o galería virtual de Fernando Botero en inglés con URL /en/artist/fernando-botero/, schema ArtWork, y páginas de obras individuales. Target: capturar 'fernando botero paintings', 'botero art for sale'.",
        "Alta"
    ),
    (
        "Dependencia crítica: Edgar Negret concentra el 45% del tráfico CO",
        "La página /artista/edgar-negret/ genera 8,786 visitas CO (45.03% del total). Si esta página pierde posición, el tráfico cae casi a la mitad. Solo 30 keywords respaldan esta concentración.",
        "seo-audit (Content Quality — E-E-A-T + Content Depth) + content-strategy (Hub and Spoke)",
        "Diversificar: crear 3-5 páginas adicionales para artistas del roster con alto potencial (omar rayo, david manzur, enrique grau) y ampliar keywords de Negret con sub-páginas de obras específicas.",
        "Alta"
    ),
    (
        "Oportunidad comercial: 'galerias bogotá' (2,900/mes CO) sin rankear",
        "3 competidores (galerialacometa, casasriegner, galeriaelmuseo) aparecen para 'galerias bogotá' — Duque Arango, con espacio en Bogotá, NO rankea. Keyword de intención comercial clara.",
        "content-strategy (Keyword Research by Buyer Stage — Consideration Stage) + seo-audit (On-Page SEO)",
        "Crear página /galerías/bogota o artículo 'Guía de galerías de arte en Bogotá' con sección sobre Duque Arango Bogotá. Schema LocalBusiness para la sede de Bogotá. Targetear también 'galeria arte bogota'.",
        "Alta"
    ),
    (
        "Omar Rayo bajó de posición 3 a 6 en CO (-3 posiciones)",
        "Semrush muestra que 'omar rayo' pasó de posición 3 a posición 6 en CO (volumen: 5,400). Primera señal de deterioro detectada para un artista del roster en esta semana base.",
        "seo-audit (Crawlability & Indexation + Content Quality) + ai-seo (Content Extractability Check)",
        "Auditar página /artista/omar-rayo/ y el blog de Omar Rayo: verificar velocidad, internal linking, actualizar contenido con datos de subasta y estadísticas citadas (GEO framework: +37% visibility con datos numéricos).",
        "Alta"
    ),
    (
        "Galería El Museo lidera en CO en contenido editorial diversificado",
        "galeriaelmuseo.com: 4,038 visitas CO con 547 keywords — ratio keywords/tráfico más eficiente que La Cometa. Rankean para 'autorretrato' (2,400/mes), 'museo de arte' (8,100), 'jacanamijoy'.",
        "competitor-alternatives (Deep Competitor Research) + content-strategy (Competitor-led pillars)",
        "Crear cluster editorial sobre técnicas y géneros del arte colombiano: autorretrato, escultura abstracta, arte geométrico. Cada artículo enlazado a páginas de artistas del roster relevantes.",
        "Media"
    ),
    (
        "Oportunidad AI-SEO: Botero en US tiene visibilidad pero sin estructura extractable",
        "En US rankeamos para 'fernando botero' en posición 22 con 18,100 búsquedas. Los artículos del blog no tienen schema Article, FAQ markup, ni bloques de respuesta directa — limitando citabilidad por AI.",
        "ai-seo (Pillar 1: Structure — Content Extractability Check) — Princeton GEO: citar fuentes y estadísticas = +40% visibilidad en AI",
        "Añadir FAQPage schema y bloques de definición clara a los artículos de Botero en inglés. Incluir datos de subastas con fuente citada (ej. 'Christie's 2023: obra de Botero en $X'). Añadir 'Last updated' visible.",
        "Media"
    ),
    (
        "Galerías US (Opera, Miguel Abreu) dominan por páginas de artista individualmente optimizadas",
        "Opera Gallery: 7,958 US con páginas por artista (Keith Haring: 659 visitas, 160 keywords). Miguel Abreu: 2,815 US. Duque Arango: 2,646 US. La diferencia está en páginas de artistas internacionales del roster (Plensa, Rondinone, Le Parc, Lam) sin versión en inglés.",
        "seo-audit (On-Page SEO — Keyword Targeting per Page) + competitor-alternatives (Research Process)",
        "Priorizar creación de páginas de artista en inglés para: Jaume Plensa, Ugo Rondinone, Julio Le Parc, Leonora Carrington, Wifredo Lam. Estos artistas tienen demanda global y Duque Arango los representa.",
        "Media"
    ),
]

for i, row in enumerate(insights):
    r = i + 4
    insight, evidencia, framework, accion, prioridad = row
    fa = NA if i % 2 else N
    nf_prio = ALTA if prioridad == "Alta" else (MEDIA if prioridad == "Media" else BAJA)
    ws5.write(r, 0, insight, WARN if i < 2 else fa)
    ws5.write(r, 1, evidencia, fa)
    ws5.write(r, 2, framework, fa)
    ws5.write(r, 3, accion, fa)
    ws5.write(r, 4, prioridad, nf_prio)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 6 — EVOLUCIÓN Y TENDENCIAS (SEMANA BASE)
# ══════════════════════════════════════════════════════════════════════════════
ws6 = wb.add_worksheet("Evolución y Tendencias")
ws6.set_zoom(85)
ws6.set_column(0, 0, 22)
ws6.set_column(1, 1, 70)

ws6.write(0, 0, "EVOLUCIÓN Y TENDENCIAS — NOTA DE SEMANA BASE", TITLE)
ws6.write(1, 0, "Esta es la primera ejecución de la rutina. No hay histórico de 4 semanas. El siguiente reporte podrá mostrar cambios.", INFO)

ws6.write(3, 0, "CONTEXTO", H)
ws6.write(3, 1, "DESCRIPCIÓN", H)

base_notes = [
    ("SEMANA BASE", "Esta es la semana baseline. Los próximos reportes clasificarán cada hallazgo como NUEVO / CONTINÚA / CAMBIÓ / DESAPARECIÓ en relación con esta semana."),
    ("Estado actual CO", "Galería Duque Arango lidera en Colombia con 19,511 visitas/mes orgánicas — 4.8x por encima del segundo competidor (galeriaelmuseo.com, 4,038). Posición sólida pero concentrada en pocos artistas (Negret = 45% del tráfico)."),
    ("Estado actual US", "En US ocupamos el puesto #3 con 2,646 visitas, detrás de operagallery.com (7,958) y miguelabreugallery.com (2,815). Brecha reducible con páginas de artistas internacionales del roster en inglés."),
    ("Patrón artistas CO", "Los artistas del roster más fuertes en CO son: Edgar Negret (#1 en visitas), Alejandro Obregón, David Manzur, Enrique Grau. Omar Rayo muestra primera caída (-3 posiciones)."),
    ("Patrón artistas US", "En US, el tráfico está dominado por contenido sobre Fernando Botero (blog) y Ana Mercedes Hoyos (artista). Falta cobertura en inglés de Wifredo Lam, Julio Le Parc, Jaume Plensa, Leonora Carrington."),
    ("Competidores a vigilar", "casasriegner.com: Beatriz González genera el 30% de su tráfico CO (8,100 búsquedas). galerialacometa.com: Emma Reyes y Feliza Bursztyn son drivers de tráfico. operagallery.com: Fernando Botero viewing room es amenaza directa en US."),
    ("Próximas semanas — qué observar", "1) ¿Sube o baja Omar Rayo en CO? 2) ¿Gana Opera Gallery más keywords de Botero en US? 3) ¿Algún competidor empieza a rankear para artistas del roster que hoy solo tienen presencia débil? 4) ¿Evoluciona casasriegner hacia más artistas de alto volumen?"),
]

for i, (ctx, desc) in enumerate(base_notes):
    r = i + 4
    fa = NA if i % 2 else N
    ws6.write(r, 0, ctx, NB if i == 0 else fa)
    ws6.write(r, 1, desc, fa)

wb.close()
print(f"Excel creado: {FILEPATH}")
import os
size = os.path.getsize(FILEPATH)
print(f"Tamaño: {size:,} bytes")
