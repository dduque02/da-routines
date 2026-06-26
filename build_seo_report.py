#!/usr/bin/env python3
"""Genera Inteligencia_SEO_2026_06_26.xlsx para Galería Duque Arango."""

import xlsxwriter
import os

OUTPUT_FILE = "/home/user/da-routines/Inteligencia_SEO_2026_06_26.xlsx"
DATE_STR = "26 de junio de 2026"
DATA_DATE = "Datos al: junio 2026 (Semrush actualización mensual)"

wb = xlsxwriter.Workbook(OUTPUT_FILE)

# ── Formatos ──────────────────────────────────────────────────────────────────
def fmt(bold=False, bg=None, font_color="#000000", align="left", wrap=False,
        border=0, font_size=10, italic=False, num_format=None):
    d = {"font_name": "Arial", "font_size": font_size, "align": align,
         "valign": "vcenter", "font_color": font_color, "bold": bold,
         "italic": italic, "border": border, "text_wrap": wrap}
    if bg:
        d["bg_color"] = bg
    if num_format:
        d["num_format"] = num_format
    return d

HEADER_DARK  = fmt(bold=True, bg="#1A237E", font_color="#FFFFFF", align="center", border=1, font_size=10)
HEADER_MED   = fmt(bold=True, bg="#283593", font_color="#FFFFFF", align="center", border=1)
HEADER_LIGHT = fmt(bold=True, bg="#E8EAF6", font_color="#1A237E", align="center", border=1)
TITLE_FMT    = fmt(bold=True, font_size=14, bg="#1A237E", font_color="#FFFFFF")
SUBTITLE_FMT = fmt(bold=True, font_size=11, bg="#3F51B5", font_color="#FFFFFF")
SECTION_FMT  = fmt(bold=True, font_size=10, bg="#E8EAF6", font_color="#1A237E", border=1)
DATA_FMT     = fmt(border=1, wrap=True)
DATA_CENTER  = fmt(border=1, align="center")
DATA_NUM     = fmt(border=1, align="right", num_format="#,##0")
DATA_PCT     = fmt(border=1, align="right", num_format="0.00%")
BOLD_DATA    = fmt(bold=True, border=1, bg="#FFF9C4")
BOLD_CENTER  = fmt(bold=True, border=1, align="center", bg="#FFF9C4")
BOLD_NUM     = fmt(bold=True, border=1, align="right", bg="#FFF9C4", num_format="#,##0")
GREEN_FMT    = fmt(bold=True, bg="#E8F5E9", font_color="#1B5E20", border=1, align="center")
RED_FMT      = fmt(bold=True, bg="#FFEBEE", font_color="#B71C1C", border=1, align="center")
ORANGE_FMT   = fmt(bold=True, bg="#FFF3E0", font_color="#E65100", border=1, align="center")
BLUE_FMT     = fmt(bold=True, bg="#E3F2FD", font_color="#0D47A1", border=1, align="center")
GRAY_FMT     = fmt(italic=True, font_color="#666666", border=1)
NOTE_FMT     = fmt(italic=True, font_color="#555555", font_size=9)

# helper: formato para estados
def estado_fmt(estado):
    if "NUEVO" in estado or "↑" in estado:
        return wb.add_format(fmt(bold=True, bg="#E8F5E9", font_color="#1B5E20", border=1, align="center"))
    elif "CAMBIÓ ↓" in estado or "ALERTA" in estado:
        return wb.add_format(fmt(bold=True, bg="#FFEBEE", font_color="#B71C1C", border=1, align="center"))
    elif "CONTINÚA" in estado:
        return wb.add_format(fmt(bold=True, bg="#E3F2FD", font_color="#0D47A1", border=1, align="center"))
    elif "DESAPARECIÓ" in estado:
        return wb.add_format(fmt(bold=True, bg="#F3E5F5", font_color="#4A148C", border=1, align="center"))
    else:
        return wb.add_format(fmt(border=1, align="center"))

# Convertir formatos a objetos workbook
def W(d):
    if isinstance(d, dict):
        return wb.add_format(d)
    return d

# ── TAB 1: Desempeño Propio ────────────────────────────────────────────────────
ws1 = wb.add_worksheet("Desempeño Propio")
ws1.set_column("A:A", 45)
ws1.set_column("B:B", 22)
ws1.set_column("C:C", 22)
ws1.set_column("D:D", 30)
ws1.set_row(0, 28)

ws1.merge_range("A1:D1", f"DESEMPEÑO PROPIO — Galería Duque Arango | {DATE_STR}", W(TITLE_FMT))
ws1.merge_range("A2:D2", DATA_DATE, W(NOTE_FMT))

# Sub-tabla métricas
ws1.write("A4", "RESUMEN DE MÉTRICAS", W(SUBTITLE_FMT))
ws1.merge_range("B4:D4", "", W(SUBTITLE_FMT))

headers_metricas = ["Métrica", "Colombia (CO)", "Estados Unidos (US)"]
for c, h in enumerate(headers_metricas):
    ws1.write(4, c, h, W(HEADER_DARK))
ws1.write(4, 3, "Variación vs. Semana Anterior", W(HEADER_DARK))

metricas = [
    ["Tráfico orgánico estimado", "20.778 visitas/mes", "1.778 visitas/mes",
     "CO: CAMBIÓ ↑ +9,7% (era 18.939) | US: CONTINÚA -1,3% (era 1.801)"],
    ["Keywords orgánicas totales", "2.503", "1.952",
     "CO: Keywords ↓ (era 4.156 — dato Semrush recalibrado) | US: Keywords ↑ (era 892)"],
    ["Posición ranking (de 17 galerías)", "#1 de 17 🥇", "#5 de 17",
     "CO: CONTINÚA #1 | US: CONTINÚA #5 (3ª semana)"],
    ["Keyword de mayor volumen", "edgar negret — pos.2 | vol. 90.500", "fernando botero — pos.12 | vol. 18.100",
     "US ALERTA: 'fernando botero' cayó pos.9→pos.11→pos.12 (3ª semana ↓)"],
    ["2ª keyword de mayor tráfico", "plaza botero — pos.5 | vol. 60.500", "ana mercedes hoyos — pos.6 | vol. 3.600",
     "US: ana mercedes hoyos emerge como keyword importante"],
    ["Tráfico de top página CO", "/artista/edgar-negret/ — 8.787 visitas (42,3%)", "—",
     "CO: Concentración en Negret se mantiene. Riesgo de dependencia de una sola página."],
    ["Tráfico de top página US", "—", "/en/blog/...boteros-artworks/ — 279 visitas (15,7%)",
     "US: Todo el tráfico es blog-based. Sin páginas de artista en inglés."],
]

for r, row in enumerate(metricas):
    for c, val in enumerate(row):
        style = W(BOLD_DATA) if c == 0 else W(DATA_FMT)
        if c == 0:
            ws1.write(5+r, c, val, W(BOLD_DATA))
        else:
            ws1.write(5+r, c, val, W(DATA_FMT))

# Top 20 Keywords CO
ws1.write("A14", "TOP 20 KEYWORDS — Colombia (CO)", W(SUBTITLE_FMT))
ws1.merge_range("B14:D14", "", W(SUBTITLE_FMT))

kw_headers = ["Keyword", "Posición", "Volumen/mes", "URL Rankeable", "Categoría SEO"]
ws1.set_column("E:E", 18)
for c, h in enumerate(kw_headers):
    ws1.write(14, c, h, W(HEADER_MED))

kws_co = [
    ("edgar negret", 2, 90500, "/artista/edgar-negret/", "Nombre Artista ROSTER"),
    ("plaza botero", 5, 60500, "/blog/plaza-botero-.../", "Editorial"),
    ("plaza botero medellín antioquia", 9, 22200, "/blog/plaza-botero-.../", "Editorial"),
    ("fernando botero", 5, 8100, "/blog/las-obras-de-fernando-botero.../", "Nombre Artista ROSTER"),
    ("escultura", 5, 5400, "/blog/escultura-en-colombia.../", "Editorial"),
    ("david manzur", 2, 4400, "/artista/david-manzur/", "Nombre Artista ROSTER"),
    ("alejandro obregon", 3, 3600, "/artista/alejandro-obregon/", "Nombre Artista ROSTER"),
    ("galeria duque arango", 1, 720, "/", "Marca"),
    ("obregon pinturas", 1, 480, "/artista/alejandro-obregon/", "Nombre Artista ROSTER"),
    ("david manzur obras", 1, 390, "/artista/david-manzur/", "Nombre Artista ROSTER"),
    ("obras de alejandro obregon", 1, 320, "/artista/alejandro-obregon/", "Nombre Artista ROSTER"),
    ("alejandro obregon obras", 1, 320, "/artista/alejandro-obregon/", "Nombre Artista ROSTER"),
    ("enrique grau obras", 1, 260, "/artista/enrique-grau/", "Nombre Artista ROSTER"),
    ("enrique grau", 2, 1600, "/artista/enrique-grau/", "Nombre Artista ROSTER"),
    ("guayasamin", 3, 2400, "/blog/el-legado-al-arte-de-oswaldo-guayasamin/", "Nombre Artista ROSTER"),
    ("david manzur biografia", 1, 170, "/artista/david-manzur/", "Nombre Artista ROSTER"),
    ("obras de enrique grau", 1, 170, "/artista/enrique-grau/", "Nombre Artista ROSTER"),
    ("pinturas de alejandro obregon", 1, 210, "/artista/alejandro-obregon/", "Nombre Artista ROSTER"),
    ("olga de amaral", 2, 1000, "/artista/olga-de-amaral/", "Nombre Artista ROSTER"),
    ("galeria arte medellin", 4, 1900, "/", "Comercial"),
]

for r, (kw, pos, vol, url, cat) in enumerate(kws_co):
    ws1.write(15+r, 0, kw, W(DATA_FMT))
    ws1.write(15+r, 1, pos, W(DATA_CENTER))
    ws1.write(15+r, 2, vol, W(DATA_NUM))
    ws1.write(15+r, 3, url, W(DATA_FMT))
    ws1.write(15+r, 4, cat, W(DATA_FMT))

# Top 20 Keywords US
ws1.write("A37", "TOP 20 KEYWORDS — Estados Unidos (US)", W(SUBTITLE_FMT))
ws1.merge_range("B37:D37", "", W(SUBTITLE_FMT))
for c, h in enumerate(kw_headers):
    ws1.write(37, c, h, W(HEADER_MED))

kws_us = [
    ("fernando botero art", 5, 3600, "/en/blog/unmistakable-brand-boterism/", "Nombre Artista ROSTER"),
    ("ana mercedes hoyos", 6, 3600, "/artista/ana-mercedes-hoyos/", "Nombre Artista ROSTER"),
    ("botero's", 6, 3600, "/en/blog/all-you-need-to-know.../", "Nombre Artista ROSTER"),
    ("botero artist", 7, 3600, "/en/blog/all-you-need-to-know.../", "Nombre Artista ROSTER"),
    ("botero paintings", 10, 3600, "/en/blog/all-you-need-to-know.../", "Nombre Artista ROSTER"),
    ("fernando botero paintings", 9, 4400, "/en/blog/all-you-need-to-know.../", "Nombre Artista ROSTER ⚠️"),
    ("fernando botero", 12, 18100, "/en/blog/all-you-need-to-know.../", "Nombre Artista ROSTER ⚠️ CAYENDO"),
    ("plaza botero medellin colombia", 6, 1900, "/en/blog/plaza-botero.../", "Editorial"),
    ("oswaldo guayasamin", 6, 1600, "/en/blog/oswaldo-guayasamins-legacy.../", "Nombre Artista ROSTER"),
    ("boterismo in columbia", 9, 1600, "/en/blog/fernando-botero-a-master.../", "Editorial"),
    ("guayasamin", 5, 1300, "/en/blog/the-works-of-oswaldo-guayasamin/", "Nombre Artista ROSTER"),
    ("fernando botero artworks", 9, 1300, "/en/blog/all-you-need-to-know.../", "Nombre Artista ROSTER"),
    ("colombian painters", 3, 720, "/en/blog/discovering-colombian-art.../", "Editorial"),
    ("boterismo", 5, 590, "/en/blog/fernando-botero-a-master.../", "Editorial"),
    ("picasso realistic art", 4, 480, "/en/blog/artistic-periods-of-pablo-picasso/", "Editorial"),
    ("plaza botero medellin", 5, 480, "/blog/plaza-botero-.../", "Editorial"),
    ("mona lisa age twelve", 7, 480, "/en/blog/fernando-botero-monalisa-at-12.../", "Editorial"),
    ("galeria de arte", 2, 480, "/", "Comercial"),
    ("fernando botero style", 1, 170, "/en/blog/unmistakable-brand-boterism/", "Editorial"),
    ("boterism", 1, 90, "/en/blog/unmistakable-brand-boterism/", "Editorial"),
]

for r, (kw, pos, vol, url, cat) in enumerate(kws_us):
    ws1.write(38+r, 0, kw, W(DATA_FMT))
    ws1.write(38+r, 1, pos, W(DATA_CENTER))
    ws1.write(38+r, 2, vol, W(DATA_NUM))
    ws1.write(38+r, 3, url, W(DATA_FMT))
    ws1.write(38+r, 4, cat, W(DATA_FMT))

# Top 10 Páginas CO
ws1.write("A60", "TOP 10 PÁGINAS — Colombia (CO)", W(SUBTITLE_FMT))
ws1.merge_range("B60:D60", "", W(SUBTITLE_FMT))
pg_headers = ["URL de Página", "Tráfico Estimado", "% del Total CO", "Nº Keywords"]
for c, h in enumerate(pg_headers):
    ws1.write(60, c, h, W(HEADER_MED))

pages_co = [
    ("/artista/edgar-negret/", 8787, 0.4228, 29),
    ("/blog/plaza-botero-.../", 2067, 0.0994, 31),
    ("/blog/las-obras-de-fernando-botero-y-su-significado/", 1610, 0.0774, 156),
    ("/artista/david-manzur/", 1339, 0.0644, 53),
    ("/artista/alejandro-obregon/", 1186, 0.0570, 78),
    ("/ (homepage)", 780, 0.0375, 122),
    ("/blog/las-obras-mejor-vendidas-de-fernando-botero/", 691, 0.0332, 92),
    ("/artista/enrique-grau/", 414, 0.0199, 22),
    ("/artista/ana-mercedes-hoyos/", 385, 0.0185, 20),
    ("/blog/artistas-colombianos-que-debes-conocer/", 328, 0.0157, 188),
]
for r, (url, traffic, pct, kws) in enumerate(pages_co):
    ws1.write(61+r, 0, url, W(DATA_FMT))
    ws1.write(61+r, 1, traffic, W(DATA_NUM))
    ws1.write(61+r, 2, pct, W(DATA_PCT))
    ws1.write(61+r, 3, kws, W(DATA_CENTER))

# Top 10 Páginas US
ws1.write("A73", "TOP 10 PÁGINAS — Estados Unidos (US)", W(SUBTITLE_FMT))
ws1.merge_range("B73:D73", "", W(SUBTITLE_FMT))
for c, h in enumerate(pg_headers):
    ws1.write(73, c, h, W(HEADER_MED))

pages_us = [
    ("/en/blog/all-you-need-to-know-to-understand-fernando-boteros-artworks/", 279, 0.1569, 106),
    ("/en/blog/unmistakable-brand-boterism/", 222, 0.1248, 80),
    ("/en/blog/discovering-colombian-art-from-fernando-botero-to-contemporary-masters/", 143, 0.0804, 67),
    ("/en/blog/artistic-periods-of-pablo-picasso/", 106, 0.0596, 157),
    ("/en/blog/the-works-of-oswaldo-guayasamin/", 85, 0.0478, 38),
    ("/blog/plaza-botero-.../", 81, 0.0455, 10),
    ("/artista/ana-mercedes-hoyos/", 79, 0.0444, 1),
    ("/en/blog/the-works-of-fernando-botero-and-their-significance/", 76, 0.0427, 70),
    ("/en/blog/fernando-botero-a-master-of-colombian-art-and-latin-american-excellence/", 56, 0.0314, 75),
    ("/en/blog/plaza-plaza-botero-resignifying-public-space-through-art-in-medellin/", 49, 0.0275, 8),
]
for r, (url, traffic, pct, kws) in enumerate(pages_us):
    ws1.write(74+r, 0, url, W(DATA_FMT))
    ws1.write(74+r, 1, traffic, W(DATA_NUM))
    ws1.write(74+r, 2, pct, W(DATA_PCT))
    ws1.write(74+r, 3, kws, W(DATA_CENTER))

ws1.merge_range("A85:D85",
    "NOTA: Semrush actualiza datos mensualmente. Variaciones pueden reflejar recalibraciones del índice además de cambios reales.",
    W(NOTE_FMT))

# ── TAB 2: Benchmark Competitivo ──────────────────────────────────────────────
ws2 = wb.add_worksheet("Benchmark Competitivo")
ws2.set_column("A:A", 32)
ws2.set_column("B:B", 10)
ws2.set_column("C:C", 14)
ws2.set_column("D:D", 14)
ws2.set_column("E:E", 14)
ws2.set_column("F:F", 14)
ws2.set_column("G:G", 14)
ws2.set_column("H:H", 14)
ws2.set_column("I:I", 22)

ws2.merge_range("A1:I1",
    f"BENCHMARK COMPETITIVO — {DATE_STR} | {DATA_DATE}", W(TITLE_FMT))

bench_headers = ["Dominio", "Tier", "Tráfico CO", "Keywords CO", "Authority CO",
                 "Tráfico US", "Keywords US", "Authority US", "Estado semana"]
for c, h in enumerate(bench_headers):
    ws2.write(1, c, h, W(HEADER_DARK))

# Datos del benchmark (ordenados por Tráfico CO desc)
# Todos los cambios vs semana anterior (Jun 19)
bench_data = [
    # dominio, tier, tr_co, kw_co, auth_co, tr_us, kw_us, auth_us, estado
    ("galeriaduquearango.com", "—", 20778, 2503, "N/D", 1778, 1952, "N/D", "CO CAMBIÓ ↑ / US CONTINÚA #5"),
    ("galeriaelmuseo.com", "Tier 1", 3609, 569, "N/D", 35, 41, "N/D", "CO CAMBIÓ ↑ +9,1%"),
    ("galerialacometa.com", "Tier 1", 3082, 374, "N/D", 516, 125, "N/D", "CO CONTINÚA"),
    ("casasriegner.com", "Tier 1", 2046, 274, "N/D", 22, 53, "N/D", "CO CONTINÚA"),
    ("sgr-art.com", "Tier 1", 1217, 63, "N/D", 0, 5, "N/D", "CO CONTINÚA"),
    ("galeriacasacuadrada.com", "Tier 1", 398, 97, "N/D", 5, 4, "N/D", "CO CONTINÚA"),
    ("galeriafreites.com", "Tier 2", 184, 113, "N/D", 3, 15, "N/D", "CO CONTINÚA"),
    ("beatrizesguerra-art.com", "Tier 1", 184, 39, "N/D", 168, 77, "N/D", "CO CONTINÚA"),
    ("otros360grados.com", "Tier 1", 121, 56, "N/D", 0, 9, "N/D", "CO CONTINÚA"),
    ("operagallery.com", "Tier 2", 43, 44, "N/D", 7746, 2886, "N/D", "US CAMBIÓ ↓ -8,6%"),
    ("galeriaelsapineres.art", "Tier 1", 15, 30, "N/D", 0, 0, "N/D", "Sin datos US"),
    ("artoftheworldgallery.com", "Tier 2", 13, 26, "N/D", 1999, 834, "N/D", "US CAMBIÓ ↑ +8,6% ⚠️"),
    ("ascasogallery.com", "Tier 2", 0, 10, "N/D", 307, 160, "N/D", "CO Sin datos"),
    ("latinartcore.com", "Tier 2", 0, 0, "N/D", 147, 102, "N/D", "CO Sin datos"),
    ("cernudaarte.com", "Tier 2", 0, 5, "N/D", 543, 382, "N/D", "CO DESAPARECIÓ ⚠️ / US CAMBIÓ ↓ -62%"),
    ("forumgallery.com", "Tier 2", 0, 6, "N/D", 1929, 1683, "N/D", "US CONTINÚA +2,7%"),
    ("miguelabreugallery.com", "Tier 2", 0, 1, "N/D", 2889, 376, "N/D", "US CONTINÚA"),
]

duque_row = 2  # 0-indexed in data (row 4 in sheet = index 2)

for r, row in enumerate(bench_data):
    dom, tier, tr_co, kw_co, auth_co, tr_us, kw_us, auth_us, estado = row
    is_duque = dom == "galeriaduquearango.com"

    df = W(BOLD_DATA) if is_duque else W(DATA_FMT)
    dnf = W(BOLD_NUM) if is_duque else W(DATA_NUM)
    dcf = W(BOLD_CENTER) if is_duque else W(DATA_CENTER)

    ws2.write(2+r, 0, dom, df)
    ws2.write(2+r, 1, tier, dcf)
    ws2.write(2+r, 2, tr_co if isinstance(tr_co, int) else tr_co, dnf if isinstance(tr_co, int) else dcf)
    ws2.write(2+r, 3, kw_co if isinstance(kw_co, int) else kw_co, dnf if isinstance(kw_co, int) else dcf)
    ws2.write(2+r, 4, auth_co, dcf)
    ws2.write(2+r, 5, tr_us if isinstance(tr_us, int) else tr_us, dnf if isinstance(tr_us, int) else dcf)
    ws2.write(2+r, 6, kw_us if isinstance(kw_us, int) else kw_us, dnf if isinstance(kw_us, int) else dcf)
    ws2.write(2+r, 7, auth_us, dcf)
    ws2.write(2+r, 8, estado, W(estado_fmt(estado)))

# Posiciones CO y US
ws2.write(20, 0, "POSICIÓN DUQUE — Colombia (CO)", W(SECTION_FMT))
ws2.write(20, 1, "#1 de 17", W(GREEN_FMT))
ws2.write(20, 2, "Lider del mercado colombiano", W(DATA_FMT))

ws2.write(21, 0, "POSICIÓN DUQUE — Estados Unidos (US)", W(SECTION_FMT))
ws2.write(21, 1, "#5 de 17", W(ORANGE_FMT))
ws2.write(21, 2, "3ª semana consecutiva en posición #5. Fernando Botero sigue cayendo (pos.12).", W(DATA_FMT))

ws2.write(22, 0, "Líder CO", W(SECTION_FMT))
ws2.write(22, 1, "galeriaduquearango.com", W(GREEN_FMT))
ws2.write(22, 2, "20.778 visitas CO — domina el mercado colombiano con 5.7x más tráfico que el #2", W(DATA_FMT))

ws2.write(23, 0, "Líder US", W(SECTION_FMT))
ws2.write(23, 1, "operagallery.com", W(ORANGE_FMT))
ws2.write(23, 2, "7.746 visitas US — 4.4x más tráfico que Duque Arango en mercado estadounidense", W(DATA_FMT))

ws2.merge_range("A25:I25",
    "NOTAS: authority score = N/D (requiere API adicional de Semrush). "
    "Tráfico CO de artoftheworldgallery, cernudaarte, miguelabreugallery y forumgallery = 0 (sin presencia en CO). "
    "latinartcore.com sin datos CO ni presencia verificada.",
    W(NOTE_FMT))

# ── TAB 3: Brechas de Keywords ─────────────────────────────────────────────────
ws3 = wb.add_worksheet("Brechas de Keywords")
ws3.set_column("A:A", 28)
ws3.set_column("B:B", 16)
ws3.set_column("C:C", 22)
ws3.set_column("D:D", 28)
ws3.set_column("E:E", 20)
ws3.set_column("F:F", 16)
ws3.set_column("G:G", 42)
ws3.set_column("H:H", 16)

ws3.merge_range("A1:H1",
    f"BRECHAS DE KEYWORDS — {DATE_STR} | Prioridad para galeriaduquearango.com", W(TITLE_FMT))

brecha_headers = ["Keyword", "Volumen de Búsqueda", "Categoría SEO",
                  "Competidor que Rankea", "Posición Competidor",
                  "Nivel de Oportunidad", "Acción de Contenido Sugerida", "Estado (vs. semanas prev.)"]
for c, h in enumerate(brecha_headers):
    ws3.write(1, c, h, W(HEADER_DARK))

brechas = [
    # kw, vol, db, cat, competidor, pos, nivel, accion, estado
    ("wilfredo lam", "3.600 (US)", "Nombre Artista ROSTER ⚠️",
     "cernudaarte.com", "pos. 52", "ALTA",
     "Crear /en/artist/wifredo-lam/ con biography, notable works, market context + schema ArtistPage",
     "CONTINÚA — 3ª semana sin acción"),
    ("galerias bogotá", "2.900 (CO)", "Comercial",
     "casasriegner / galerialacometa / galeriaelmuseo", "pos. 11-23", "ALTA",
     "Crear /es/blog/galerias-de-bogota/ — guía con posicionamiento de Duque Arango Bogotá",
     "CONTINÚA — 3ª semana sin acción"),
    ("autorretrato", "2.400 (CO)", "Editorial",
     "galeriaelmuseo.com", "pos. 5", "ALTA",
     "Crear /es/blog/autorretrato-en-el-arte-colombiano/ con obras de artistas del roster (Obregón, Grau, Manzur)",
     "CONTINÚA — 3ª semana sin acción"),
    ("jesus abad colorado", "2.400 (CO)", "Editorial / Artista",
     "galeriaelmuseo.com", "pos. 3", "ALTA",
     "Artículo editorial sobre fotografía colombiana contemporánea — mencionar contexto de arte colombiano",
     "NUEVO esta semana"),
    ("emma reyes", "1.900 (CO)", "Editorial / Artista histórica",
     "galerialacometa.com", "pos. 4", "MEDIA",
     "Artículo editorial: 'Mujeres artistas colombianas del siglo XX' — incluye Emma Reyes en contexto",
     "CONTINÚA (galerialacometa la capitaliza 3ª semana)"),
    ("feliza bursztyn", "1.900 (CO)", "Editorial / Artista histórica",
     "galerialacometa.com", "pos. 32", "MEDIA",
     "Mismo artículo de mujeres artistas — incluir a Bursztyn como contexto histórico colombiano",
     "CONTINÚA (galerialacometa la capitaliza 3ª semana)"),
    ("beatriz gonzalez", "8.100 (CO)", "Editorial / Artista (no en roster)",
     "casasriegner.com", "pos. 5", "MEDIA",
     "Artículo editorial de contexto: 'Arte político colombiano — de Beatriz González a contemporáneos' — capturar tráfico informacional",
     "NUEVO — artista principal de casasriegner, volumen muy alto"),
    ("amelia pelaez", "1.000 (US)", "Nombre Artista (no en roster)",
     "cernudaarte.com", "pos. 6", "MEDIA",
     "Artículo editorial sobre arte cubano y latinoamericano — mencionar artistas afines del roster (Wifredo Lam)",
     "NUEVO esta semana"),
    ("wifredo lam paintings", "590 (US)", "Nombre Artista ROSTER ⚠️",
     "cernudaarte.com", "pos. 52", "MEDIA",
     "Incluir en página /en/artist/wifredo-lam/ — las obras principales con imágenes y contexto de mercado",
     "CONTINÚA — 3ª semana (relacionado con wilfredo lam)"),
    ("bienal de arte medellin", "880 (CO)", "Editorial / Evento",
     "casasriegner.com", "pos. 16", "MEDIA",
     "Artículo anual sobre la Bienal con perspectiva de artistas representados por la galería",
     "NUEVO esta semana"),
    ("tomas sanchez", "880 (US)", "Nombre Artista (no en roster)",
     "cernudaarte.com", "pos. 18", "MEDIA",
     "Artículo editorial sobre arte latinoamericano en US — mención en contexto con artistas del roster",
     "NUEVO esta semana"),
    ("carlos rojas", "1.300 (CO)", "Nombre Artista (no en roster)",
     "casasriegner.com", "pos. 1", "BAJA",
     "Contexto editorial sobre arte óptico y cinético colombiano — mencionar Omar Rayo y Edgar Negret del roster",
     "CONTINÚA — casasriegner lo usa como artista principal"),
    ("nc arte", "880 (CO)", "Editorial / Institución",
     "casasriegner.com", "pos. 33", "BAJA",
     "Mencionar NC Arte en contexto de guía de galerías de Bogotá (ver brecha 'galerias bogotá')",
     "NUEVO esta semana"),
    ("pedro ruiz", "720 (CO)", "Nombre Artista (no en roster)",
     "beatrizesguerra-art.com", "pos. 10", "BAJA",
     "Artículo editorial sobre arte geométrico y abstracción en Colombia — contexto histórico",
     "NUEVO esta semana"),
    ("ana mercedes hoyos en ingles", "3.600 (US)", "Nombre Artista ROSTER — OPORTUNIDAD",
     "Duque rankea pos.6", "pos. 6 (Duque)", "ALTA",
     "OPORTUNIDAD PROPIA: Crear /en/artist/ana-mercedes-hoyos/ con contenido en inglés — ya genera 79 visitas US/mes sin página dedicada en inglés",
     "NUEVO — keyword Duque ya rankea, necesita página dedicada EN"),
]

for r, (kw, vol, cat, comp, pos, nivel, accion, estado) in enumerate(brechas):
    ws3.write(2+r, 0, kw, W(DATA_FMT))
    ws3.write(2+r, 1, vol, W(DATA_CENTER))
    ws3.write(2+r, 2, cat, W(DATA_FMT))
    ws3.write(2+r, 3, comp, W(DATA_FMT))
    ws3.write(2+r, 4, pos, W(DATA_CENTER))
    # Color by level
    if nivel == "ALTA":
        ws3.write(2+r, 5, nivel, W(RED_FMT))
    elif nivel == "MEDIA":
        ws3.write(2+r, 5, nivel, W(ORANGE_FMT))
    else:
        ws3.write(2+r, 5, nivel, W(BLUE_FMT))
    ws3.write(2+r, 6, accion, W(DATA_FMT))
    ws3.write(2+r, 7, estado, W(DATA_FMT))

# ── TAB 4: Páginas Top de Competencia ─────────────────────────────────────────
ws4 = wb.add_worksheet("Páginas Top Competencia")
ws4.set_column("A:A", 24)
ws4.set_column("B:B", 10)
ws4.set_column("C:C", 58)
ws4.set_column("D:D", 16)
ws4.set_column("E:E", 22)
ws4.set_column("F:F", 22)
ws4.set_column("G:G", 38)

ws4.merge_range("A1:G1",
    f"PÁGINAS TOP DE COMPETENCIA — {DATE_STR}", W(TITLE_FMT))

pg_headers = ["Competidor", "Mercado", "URL de Página", "Tráfico Estimado",
              "Keyword Principal", "Tipo de Contenido", "Brecha para Duque Arango"]
for c, h in enumerate(pg_headers):
    ws4.write(1, c, h, W(HEADER_DARK))

top_pages = [
    # CO — Top 3 competidores
    ("galeriaelmuseo.com", "CO", "galeriaelmuseo.com/", 1081, "galeria el museo", "Homepage", "Duque supera con 20.778 vs 3.609 — sin acción urgente"),
    ("galeriaelmuseo.com", "CO", "galeriaelmuseo.com/archives/43710/", 866, "autorretrato (estimado)", "Artículo editorial", "BRECHA: 'autorretrato' (2.400/mes) — Duque no tiene página para este keyword"),
    ("galeriaelmuseo.com", "CO", "galeriaelmuseo.com/archives/450/", 304, "historia de arte CO (estimado)", "Artículo editorial", "Contenido de profundidad histórica que Duque no tiene en español"),
    ("galeriaelmuseo.com", "CO", "galeriaelmuseo.com/archives/2109/", 276, "jesús abad colorado (estimado)", "Artículo artista", "BRECHA: 'jesus abad colorado' (2.400/mes) — Duque sin página"),
    ("galeriaelmuseo.com", "CO", "galeriaelmuseo.com/archives/26977/", 196, "arte colombiano tema (estimado)", "Artículo editorial", "Modelo de blog con contenido de artistas/exposiciones — patrón a replicar"),

    ("galerialacometa.com", "CO", "galerialacometa.com/", 2285, "galeria la cometa", "Homepage (74% del tráfico)", "VULNERABILIDAD: 74% homepage-dependent. Duque tiene mayor diversificación."),
    ("galerialacometa.com", "CO", "galerialacometa.com/exhibiciones/bogota/gabriela-pinilla-clandestina", 197, "gabriela pinilla", "Página de exposición", "Modelo de exposición en Bogotá — Duque podría replicar para sus exhibiciones"),
    ("galerialacometa.com", "CO", "galerialacometa.com/artistas/carlos-castro-es", 95, "carlos castro artista", "Página de artista", "Patrón: páginas de artista en español generan tráfico sostenible"),
    ("galerialacometa.com", "CO", "galerialacometa.com/artistas/miguel-angel-rojas-es", 86, "miguel angel rojas", "Página de artista", "Artistas colombianos contemporáneos con buena tracción orgánica"),
    ("galerialacometa.com", "CO", "galerialacometa.com/exhibiciones/bogota/emma-reyes-las-caras-de-emma-reyes", 83, "emma reyes", "Página exposición/artista", "BRECHA: 'emma reyes' (1.900/mes) — Duque sin contenido sobre esta artista"),

    ("casasriegner.com", "CO", "casasriegner.com/artistas/carlos-rojas", 686, "carlos rojas", "Página de artista", "Modelo exitoso de página de artista dedicada — replicar para artistas roster Duque"),
    ("casasriegner.com", "CO", "casasriegner.com/artistas/beatriz-gonzalez", 456, "beatriz gonzalez", "Página de artista", "SEÑAL: 'beatriz gonzalez' (8.100/mes) — casasriegner capitaliza con página dedicada"),
    ("casasriegner.com", "CO", "casasriegner.com/", 381, "casa riegner galeria", "Homepage", "Solo 18% del tráfico viene del homepage — modelo de artista pages más robusto"),
    ("casasriegner.com", "CO", "casasriegner.com/artistas/maria-teresa-hincapie", 90, "maria teresa hincapie", "Página de artista", "Artistas colombianas históricas con búsquedas — nicho editorial"),
    ("casasriegner.com", "CO", "casasriegner.com/artistas/rosemberg-sandoval", 79, "rosemberg sandoval", "Página de artista", "Artistas colombianos contemporáneos — patrón de páginas dedicadas funciona"),

    # US — Top 3 competidores
    ("operagallery.com", "US", "operagallery.com/", 1540, "opera gallery", "Homepage", "Opera Gallery distribuye bien — solo 19% del tráfico en homepage"),
    ("operagallery.com", "US", "operagallery.com/artist/bernard-buffet", 894, "bernard buffet artist", "Página de artista", "Modelo Hub-and-Spoke: página de artista con 36 keywords y 894 visitas/mes"),
    ("operagallery.com", "US", "operagallery.com/artist/keith-haring", 794, "keith haring art", "Página de artista", "161 keywords por página de artista — contenido muy profundo"),
    ("operagallery.com", "US", "operagallery.com/artist/simon-hantai", 431, "simon hantai", "Página de artista", "Artista con alta demanda US — estrategia de cobertura amplia"),
    ("operagallery.com", "US", "operagallery.com/artist/fernando-botero", 310, "fernando botero gallery", "Página de artista ROSTER ⚠️", "AMENAZA: Opera Gallery tiene 310 visitas US/mes para Botero. Duque NO tiene página /en/artist/ para Botero"),

    ("miguelabreugallery.com", "US", "miguelabreugallery.com/", 1923, "miguel abreu gallery", "Homepage (67% del tráfico)", "66% homepage-dependent — vulnerable. Duque más diversificado."),
    ("miguelabreugallery.com", "US", "miguelabreugallery.com/artists/quaytman/", 155, "quaytman artist", "Página de artista", "Artistas de nicho con páginas dedicadas generan tráfico estable"),
    ("miguelabreugallery.com", "US", "miguelabreugallery.com/artists/francois-marie-banier/", 118, "francois marie banier", "Página de artista", "Artistas internacionales con búsquedas activas en US"),
    ("miguelabreugallery.com", "US", "miguelabreugallery.com/artists/rayne/", 89, "rayne artist", "Página de artista", "Modelo consistente: páginas de artistas con URLs limpias"),
    ("miguelabreugallery.com", "US", "miguelabreugallery.com/meanwhile-5-agematsu/", 68, "yuji agematsu artist", "Página exposición", "Exposiciones digitales con SEO generan tráfico adicional"),

    ("artoftheworldgallery.com", "US", "artoftheworldgallery.com/", 1617, "art of the world gallery", "Homepage (81% del tráfico)", "81% homepage-dependent — muy vulnerable pero creciendo rápido (+8,6%)"),
    ("artoftheworldgallery.com", "US", "artoftheworldgallery.com/represented-artists/bahk-seon-ghi/", 191, "bahk seon ghi sculptor", "Página de artista", "Artista escultora asiática con demanda US — diversificación geográfica"),
    ("artoftheworldgallery.com", "US", "artoftheworldgallery.com/blog/botero-blog/", 44, "botero art blog", "Blog / artista ROSTER ⚠️", "AMENAZA: Tienen blog de Botero generando tráfico US. Duque debe contrarrestar con página dedicada."),
    ("artoftheworldgallery.com", "US", "artoftheworldgallery.com/product-category/represented-artists/fernando-botero-cat/", 35, "fernando botero for sale", "Página comercial ROSTER ⚠️", "AMENAZA COMERCIAL: Página de venta de Botero con intent transaccional. Duque no tiene equivalente."),
    ("artoftheworldgallery.com", "US", "artoftheworldgallery.com/artists-2/", 20, "latin american artists gallery", "Catálogo de artistas", "Catálogo general de artistas con tracción de cola larga"),
]

for r, row in enumerate(top_pages):
    comp, mkt, url, traffic, kw_main, tipo, brecha = row
    ws4.write(2+r, 0, comp, W(DATA_FMT))
    ws4.write(2+r, 1, mkt, W(DATA_CENTER))
    ws4.write(2+r, 2, url, W(DATA_FMT))
    ws4.write(2+r, 3, traffic, W(DATA_NUM))
    ws4.write(2+r, 4, kw_main, W(DATA_FMT))
    ws4.write(2+r, 5, tipo, W(DATA_FMT))
    ws4.write(2+r, 6, brecha, W(DATA_FMT))

# ── TAB 5: Insights y Recomendaciones ─────────────────────────────────────────
ws5 = wb.add_worksheet("Insights y Recomendaciones")
ws5.set_column("A:A", 38)
ws5.set_column("B:B", 42)
ws5.set_column("C:C", 32)
ws5.set_column("D:D", 50)
ws5.set_column("E:E", 12)

ws5.merge_range("A1:E1",
    f"INSIGHTS Y RECOMENDACIONES — {DATE_STR}", W(TITLE_FMT))

ins_headers = ["Insight", "Evidencia", "Framework Aplicado", "Acción Recomendada", "Prioridad"]
for c, h in enumerate(ins_headers):
    ws5.write(1, c, h, W(HEADER_DARK))

insights = [
    ("'Fernando Botero' sigue cayendo en US (3ª semana consecutiva)",
     "pos.9 (Jun 12) → pos.11 (Jun 19) → pos.12 (Jun 26). artoftheworldgallery.com tiene /product-category/fernando-botero-cat/ con intent transaccional. operagallery.com/artist/fernando-botero genera 310 visitas/mes.",
     "SEO Audit — On-Page SEO: título/URL de blog vs. página de artista. Content Strategy — Hub and Spoke: blog post como spoke sin hub de artista.",
     "URGENTE: Crear /en/artist/fernando-botero/ como página pillar (hub) con: biography, notable works, market value, exhibition history. Aplicar schema ArtistPage + schema ItemList. Redirigir señales SEO del blog hacia esta URL.",
     "ALTA"),

    ("Wifredo Lam: 3ª semana sin presencia orgánica — brecha crítica de roster",
     "Keyword 'wilfredo lam' (3.600/mes US) + 'wifredo lam paintings' (590/mes US). cernudaarte.com rankea pos.52. galeriaduquearango.com = CERO presencia. Artista del roster sin página de artista.",
     "AI SEO — Pillar 2 (Authority): artista del roster sin contenido extractable. SEO Audit — Content Quality: brecha estructural de E-E-A-T.",
     "URGENTE: Crear /en/artist/wifredo-lam/ con biography, notable works (La Jungle, El tercero), market context, y FAQ schema. Potencial: +400-600 visitas/mes US en 60-90 días. Sin esta página, cernudaarte capturará posición más alta cuando recupere su tráfico.",
     "ALTA"),

    ("ana mercedes hoyos: oportunidad US no explotada — keyword de 3.600/mes sin página EN",
     "/artista/ana-mercedes-hoyos/ genera 79 visitas US/mes (4.44% del US total) con solo 1 keyword indexado y sin página en inglés. El keyword 'ana mercedes hoyos' tiene 3.600 búsquedas/mes en US.",
     "Content Strategy — Keyword Research by Buyer Stage: awareness stage con alta demanda. AI SEO — Pillar 1 (Structure): página en español no es extractable para AI en inglés.",
     "ALTA: Crear /en/artist/ana-mercedes-hoyos/ con contenido bilingüe. Potencial enorme dado que ya hay demanda comprobada sin página dedicada. Incluir: biography en inglés, obras representativas, contexto de mercado, FAQ schema.",
     "ALTA"),

    ("artoftheworldgallery.com acelera en US con targeting directo de Botero",
     "Jun 19: 1.840 US → Jun 26: 1.999 US (+8,6%). Tienen: /blog/botero-blog/ (44 vis) + /product-category/fernando-botero-cat/ (35 vis, intent transaccional). Su crecimiento es el más acelerado entre los competidores US.",
     "Competitor Alternatives — Research Process: competidor emergente con estrategia de artistas del roster. Marketing Psychology: ataque directo al keyword de mayor valor de Duque.",
     "Monitorear sus keywords semanalmente. Contraatacar con página /en/artist/fernando-botero/ que tenga mayor autoridad y profundidad de contenido. El galería debe recuperar posición pos.9-10 en 'fernando botero paintings' antes de que artoftheworldgallery escale más.",
     "ALTA"),

    ("CO Recovery (+9,7%) confirma liderazgo sostenible pero concentración en Edgar Negret es riesgo",
     "CO: 20.778 (+9,7% vs 18.939 de Jun 19). Pero /artista/edgar-negret/ genera 8.787 visitas (42,3% del total CO). Una sola página = 42% del tráfico total. Si Negret pierde posición, el tráfico CO colapsa.",
     "SEO Audit — Content Quality: concentración extrema en un artista. Content Strategy — Hub and Spoke: modelo dependiente de una sola spoke sin hub complementario.",
     "Diversificar tráfico CO: crear/fortalecer páginas de artistas con volumen medio (david manzur 4.400/mes, edgar negret obras 90.500 ya bien posicionado). Desarrollar contenido editorial de cola larga que reduzca dependencia en Negret.",
     "MEDIA"),

    ("cernudaarte.com colapsa en CO: de 1.203 a 0 visitas — oportunidad táctica",
     "Jun 19: 1.203 CO | Jun 26: 0 CO (5 keywords, sin tráfico). Caída de -100% en mercado colombiano. Posible penalización o restructuración del sitio. Siguen activos en US (543 visitas).",
     "Competitor Alternatives — Research Process: ventana táctica cuando competidor pierde posiciones. SEO Audit — Crawlability: señal de alerta en sitio competidor.",
     "Oportunidad inmediata: publicar contenido CO sobre artistas que cernudaarte cubría (arte cubano, Wifredo Lam en español). Si la caída es por penalización, Google puede redistribuir ese tráfico en 30-60 días.",
     "MEDIA"),

    ("galerialacometa.com 74% homepage-dependent — modelo frágil vs. Duque",
     "galerialacometa.com: 2.285 de 3.082 visitas CO desde homepage. Páginas de artista: max 197 visitas. Duque: homepage solo 3,75% (780 visitas) — mucho más diversificado.",
     "Content Strategy — Hub and Spoke: Duque tiene mejor arquitectura de contenido. AI SEO — Pillar 1 (Structure): páginas de artista de Duque son más extractables.",
     "Mantener y acelerar el modelo de páginas de artista. Añadir más artistas del roster al modelo (artistas contemporáneos como Javier Caraballo, Alejandra Aristizábal, Gustavo Vélez aún sin páginas con tracción). Esto amplía la ventaja estructural sobre galerialacometa.",
     "MEDIA"),

    ("Opera Gallery: primera caída en US (-8,6%) después de 2 semanas de crecimiento",
     "Jun 12: 7.958 → Jun 19: 8.471 (+6,5%) → Jun 26: 7.746 (-8,6%). Primera vez que baja. Su página /artist/bernard-buffet (894 vis) supera ahora a /artist/keith-haring (794 vis). Botero page: 310 vis/mes.",
     "Competitor Alternatives — Research Process: monitoreo de tendencia. Marketing Psychology: pullback puede ser temporal o señal de cambio algorítmico.",
     "No interpretar como desaceleración permanente. Monitorear próxima semana. La caída en Botero (si existe) podría ser oportunidad para que Duque recupere posiciones con página dedicada. Continuar vigilando /artist/fernando-botero en Opera Gallery.",
     "BAJA"),
]

for r, (insight, evidencia, framework, accion, prioridad) in enumerate(insights):
    ws5.write(2+r, 0, insight, W(DATA_FMT))
    ws5.write(2+r, 1, evidencia, W(DATA_FMT))
    ws5.write(2+r, 2, framework, W(DATA_FMT))
    ws5.write(2+r, 3, accion, W(DATA_FMT))
    if prioridad == "ALTA":
        ws5.write(2+r, 4, prioridad, W(RED_FMT))
    elif prioridad == "MEDIA":
        ws5.write(2+r, 4, prioridad, W(ORANGE_FMT))
    else:
        ws5.write(2+r, 4, prioridad, W(BLUE_FMT))

    ws5.set_row(2+r, 80)

# ── TAB 6: Evolución y Tendencias ─────────────────────────────────────────────
ws6 = wb.add_worksheet("Evolución y Tendencias")
ws6.set_column("A:A", 35)
ws6.set_column("B:B", 70)
ws6.set_column("C:C", 30)

ws6.merge_range("A1:C1",
    f"EVOLUCIÓN Y TENDENCIAS — Últimas 3 semanas (12 jun → 19 jun → 26 jun 2026)", W(TITLE_FMT))

evo_headers = ["Patrón / Tendencia", "Descripción multi-semana", "Clasificación"]
for c, h in enumerate(evo_headers):
    ws6.write(1, c, h, W(HEADER_DARK))

evolucion = [
    ("'Fernando Botero' US: 3ª semana cayendo (pos.9→11→12)",
     "Jun 12: pos.9 | Jun 19: pos.11 | Jun 26: pos.12. Caída sostenida sin reversión. Causa: todo el tráfico Botero US viene de blog posts, sin página de artista dedicada. Opera Gallery y artoftheworldgallery atacan con páginas de artista. La tendencia se AGRAVA cada semana.",
     "TENDENCIA NEGATIVA — Acción urgente"),

    ("Wifredo Lam: brecha activa 3 semanas sin acción",
     "Jun 12: detectado por primera vez (cernudaarte pos.54). Jun 19: 2ª semana alerta. Jun 26: 3ª semana — brecha confirmada ('wilfredo lam' 3.600/mes + 'wifredo lam paintings' 590/mes US). Galeriaduquearango.com = CERO presencia en todos los keywords del artista. La brecha estructural crece cada semana sin acción.",
     "TENDENCIA CRÍTICA — Sin acción = pérdida permanente"),

    ("Duque CO: fluctuación semana 2, rebote semana 3 (patrón estable)",
     "Jun 12: 19.511 | Jun 19: 18.939 (-3%) | Jun 26: 20.778 (+9,7%). La caída de semana 2 fue una fluctuación del índice Semrush, no una tendencia real. La galería mantiene #1 en CO con diferencia amplia vs #2 (galeriaelmuseo 3.609). Liderazgo CO es robusto.",
     "TENDENCIA POSITIVA — Monitorear"),

    ("artoftheworldgallery.com: aceleración US con targeting de artistas del roster",
     "Jun 12: sin datos | Jun 19: 1.840 US | Jun 26: 1.999 US (+8,6%). Tienen páginas de Botero con intent transaccional. Es el competidor de mayor crecimiento en US en 3 semanas consecutivas. Si siguen a este ritmo, podrían superar a Duque en 4-6 semanas.",
     "TENDENCIA ALERTA — Competidor emergente"),

    ("galerialacometa.com CO: 74% homepage-dependent, sin cambio en 3 semanas",
     "Jun 12: 3.326 | Jun 19: no reportado individualmente | Jun 26: 3.082. Tráfico estable pero 74% concentrado en homepage (2.285 visitas). Páginas de artista no crecen. Sin cambios en estrategia de contenido en 3 semanas. Modelo frágil que no escala.",
     "TENDENCIA ESTANCADA — Ventaja de Duque"),

    ("Opera Gallery US: crecimiento 2 semanas, primera caída en semana 3",
     "Jun 12: 7.958 | Jun 19: 8.471 (+6,5%) | Jun 26: 7.746 (-8,6%). Patrón de crecimiento interrumpido. Primera señal de que su dominio no es ilimitado. Página de Botero genera 310 visitas/mes — amenaza activa pero con señal mixta.",
     "TENDENCIA MIXTA — Monitorear"),
]

for r, (patron, desc, clasif) in enumerate(evolucion):
    ws6.write(2+r, 0, patron, W(DATA_FMT))
    ws6.write(2+r, 1, desc, W(DATA_FMT))
    ws6.write(2+r, 2, clasif, W(DATA_FMT))
    ws6.set_row(2+r, 80)

ws6.merge_range("A9:C9", "RESUMEN EJECUTIVO DE 3 SEMANAS", W(SUBTITLE_FMT))
ws6.merge_range("A10:C10",
    "Galería Duque Arango mantiene dominio absoluto en Colombia (#1 de 17, 5.7x más tráfico que el #2). "
    "El mercado US es el desafío: posición #5 se mantiene por tercera semana consecutiva, pero la caída de "
    "'fernando botero' (pos.9→12) indica que sin páginas de artista dedicadas en inglés, la posición seguirá deteriorándose. "
    "Los tres hallazgos que se repiten semana tras semana sin corrección — Fernando Botero US, Wifredo Lam, Galerías Bogotá — "
    "deben ser prioridad de contenido inmediata para la semana del 26 de junio.",
    W(DATA_FMT))
ws6.set_row(10, 80)

wb.close()
print(f"Excel creado: {OUTPUT_FILE}")
print(f"Tamaño: {os.path.getsize(OUTPUT_FILE):,} bytes")
