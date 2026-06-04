#!/usr/bin/env python3
"""
Construye el reporte Excel semanal de Inteligencia SEO
para Galería Duque Arango — semana del 2026-06-04
"""

import xlsxwriter
import os

DATE = "2026-06-04"
FILENAME = f"Inteligencia_SEO_{DATE}.xlsx"
FILEPATH = os.path.join(os.path.dirname(__file__), FILENAME)

wb = xlsxwriter.Workbook(FILEPATH)

# ── Formats ────────────────────────────────────────────────────────────────
def fmt(wb, bold=False, bg=None, font_color="#000000", size=10,
        align="left", valign="vcenter", border=1, num_format=None,
        italic=False, wrap=False):
    d = {
        "font_name": "Calibri", "font_size": size,
        "font_color": font_color, "bold": bold, "italic": italic,
        "align": align, "valign": valign, "border": border,
        "text_wrap": wrap,
    }
    if bg:
        d["bg_color"] = bg
    if num_format:
        d["num_format"] = num_format
    return wb.add_format(d)

DARK_BLUE   = "#1F3864"
MID_BLUE    = "#2F5496"
LIGHT_BLUE  = "#BDD7EE"
ACCENT_GOLD = "#C9A84C"
HIGHLIGHT   = "#FFEB9C"
GREEN_BG    = "#E2EFDA"
RED_BG      = "#FCE4D6"
GRAY_ALT    = "#F2F2F2"
WHITE       = "#FFFFFF"

f_title   = fmt(wb, bold=True, bg=DARK_BLUE, font_color=WHITE, size=13, align="center")
f_h1      = fmt(wb, bold=True, bg=MID_BLUE,  font_color=WHITE, size=11, align="center")
f_h2      = fmt(wb, bold=True, bg=LIGHT_BLUE, font_color=DARK_BLUE, size=10, align="center")
f_kpi_lbl = fmt(wb, bold=True, bg=GRAY_ALT,  font_color=DARK_BLUE, size=10)
f_kpi_val = fmt(wb, bold=True, bg=WHITE,     font_color=MID_BLUE,  size=11, align="center")
f_body    = fmt(wb, size=10)
f_body_alt= fmt(wb, size=10, bg=GRAY_ALT)
f_bold_body= fmt(wb, bold=True, size=10)
f_num     = fmt(wb, size=10, align="center")
f_num_alt = fmt(wb, size=10, align="center", bg=GRAY_ALT)
f_pct     = fmt(wb, size=10, align="center", num_format="0.00%")
f_pct_alt = fmt(wb, size=10, align="center", num_format="0.00%", bg=GRAY_ALT)
f_cur     = fmt(wb, size=10, align="center", num_format="#,##0")
f_cur_alt = fmt(wb, size=10, align="center", num_format="#,##0", bg=GRAY_ALT)
f_url     = fmt(wb, size=9,  font_color="#1155CC", italic=True)
f_url_alt = fmt(wb, size=9,  font_color="#1155CC", italic=True, bg=GRAY_ALT)
f_wrap    = fmt(wb, size=9,  wrap=True)
f_wrap_alt= fmt(wb, size=9,  wrap=True, bg=GRAY_ALT)
f_high    = fmt(wb, bold=True, bg=RED_BG, align="center", size=10)
f_med     = fmt(wb, bold=True, bg=HIGHLIGHT, align="center", size=10)
f_low     = fmt(wb, bold=True, bg=GREEN_BG, align="center", size=10)
f_our_row = fmt(wb, bold=True, bg=ACCENT_GOLD, font_color=DARK_BLUE, size=10, align="center")
f_our_lbl = fmt(wb, bold=True, bg=ACCENT_GOLD, font_color=DARK_BLUE, size=10)
f_note    = fmt(wb, italic=True, font_color="#595959", size=9, border=0)
f_section = fmt(wb, bold=True, bg=LIGHT_BLUE, font_color=DARK_BLUE, size=10)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — DESEMPEÑO PROPIO
# ══════════════════════════════════════════════════════════════════════════════
ws1 = wb.add_worksheet("Desempeño Propio")
ws1.set_zoom(90)
ws1.set_column("A:A", 42)
ws1.set_column("B:B", 16)
ws1.set_column("C:C", 16)
ws1.set_column("D:D", 50)

# Title
ws1.merge_range("A1:D1", "DESEMPEÑO PROPIO — galeriaduquearango.com", f_title)
ws1.merge_range("A2:D2", f"Datos al: mayo 2026  |  Fecha de reporte: {DATE}  |  Fuente: Semrush", f_note)
ws1.set_row(0, 22)

# ── KPI Overview ──
ws1.merge_range("A4:D4", "Resumen de Tráfico Orgánico", f_h1)
ws1.set_row(3, 18)

kpi_headers = ["Métrica", "Colombia (CO)", "Estados Unidos (US)", "Notas"]
for c, h in enumerate(kpi_headers):
    ws1.write(4, c, h, f_h2)

kpi_data = [
    ["Keywords orgánicas indexadas", "2.219", "1.905",
     "Mayor base de keywords en CO refleja contenido en español"],
    ["Tráfico orgánico estimado / mes", "21.243", "2.215",
     "CO lidera con 10x más tráfico que US"],
    ["Valor del tráfico orgánico (USD)", "$876", "$465",
     "Costo equivalente si se pagara con Google Ads"],
    ["Keywords en Top 3 (CO / US)", "Múltiples pos. #1", "Pos. #1 en 3 keywords",
     "Edgar Negret pos. 2 con 90.500 vol/mes CO es el mayor driver"],
    ["Ranking global Semrush", "4.976", "608.115",
     "Posición global del dominio — menor = más fuerte"],
]
for r, row in enumerate(kpi_data):
    bg = WHITE if r % 2 == 0 else GRAY_ALT
    f_lbl = fmt(wb, bold=True, bg=bg, size=10)
    f_val = fmt(wb, bg=bg, size=10, align="center")
    f_nt  = fmt(wb, bg=bg, size=9, italic=True)
    ws1.write(5 + r, 0, row[0], f_lbl)
    ws1.write(5 + r, 1, row[1], f_val)
    ws1.write(5 + r, 2, row[2], f_val)
    ws1.write(5 + r, 3, row[3], f_nt)

# ── Top 20 Keywords CO ──
ws1.merge_range("A12:D12", "Top 20 Keywords — Colombia (CO) | ordenados por contribución de tráfico", f_h1)
ws1.set_row(11, 18)

kw_heads = ["Keyword", "Posición", "Volumen de Búsqueda", "URL que Rankea"]
for c, h in enumerate(kw_heads):
    ws1.write(12, c, h, f_h2)

co_keywords = [
    ("edgar negret", 2, 90500, "galeriaduquearango.com/artista/edgar-negret/"),
    ("plaza botero medellín antioquia", 5, 22200, "galeriaduquearango.com/blog/plaza-botero-resignificar-el-espacio-publico-a-traves-del-arte-en-medellin/"),
    ("plaza botero", 5, 49500, "galeriaduquearango.com/blog/plaza-botero-resignificar-el-espacio-publico-a-traves-del-arte-en-medellin/"),
    ("galeria duque arango", 1, 720, "galeriaduquearango.com/"),
    ("obras de alejandro obregón", 1, 390, "galeriaduquearango.com/artista/alejandro-obregon/"),
    ("david manzur", 2, 4400, "galeriaduquearango.com/artista/david-manzur/"),
    ("edgar negret (blog)", 13, 90500, "galeriaduquearango.com/blog/edgar-negret-el-escultor-colombiano-que-transformo-el-metal-en-poesia/"),
    ("enrique grau obras", 1, 260, "galeriaduquearango.com/artista/enrique-grau/"),
    ("enrique grau", 2, 1600, "galeriaduquearango.com/artista/enrique-grau/"),
    ("ana mercedes hoyos obras", 1, 210, "galeriaduquearango.com/artista/ana-mercedes-hoyos/"),
    ("pinturas de alejandro obregón", 1, 210, "galeriaduquearango.com/artista/alejandro-obregon/"),
    ("alejandro obregon obras", 1, 210, "galeriaduquearango.com/artista/alejandro-obregon/"),
    ("obras de arte de david manzur", 1, 170, "galeriaduquearango.com/artista/david-manzur/"),
    ("botero cuadros", 1, 170, "galeriaduquearango.com/blog/las-obras-mejor-vendidas-de-fernando-botero/"),
    ("edgar negret obras", 1, 140, "galeriaduquearango.com/artista/edgar-negret/"),
    ("javier caraballo", 1, 140, "galeriaduquearango.com/artista/javier-caraballo/"),
    ("grau", 2, 1300, "galeriaduquearango.com/artista/enrique-grau/"),
    ("muralismo", 3, 1600, "galeriaduquearango.com/blog/muralismo-latinoamericano-el-arte-como-instrumento-politico/"),
    ("obregón pinturas", 2, 480, "galeriaduquearango.com/artista/alejandro-obregon/"),
    ("cuadros de botero", 1, 480, "galeriaduquearango.com/blog/las-obras-mejor-vendidas-de-fernando-botero/"),
]
for r, (kw, pos, vol, url) in enumerate(co_keywords):
    fb = f_body if r % 2 == 0 else f_body_alt
    fn = f_num  if r % 2 == 0 else f_num_alt
    fu = f_url  if r % 2 == 0 else f_url_alt
    ws1.write(13 + r, 0, kw, fb)
    ws1.write(13 + r, 1, pos, fn)
    ws1.write(13 + r, 2, vol, fn)
    ws1.write(13 + r, 3, url, fu)

# ── Top 20 Keywords US ──
start_us = 35
ws1.merge_range(start_us, 0, start_us, 3,
    "Top 20 Keywords — Estados Unidos (US) | ordenados por contribución de tráfico", f_h1)
ws1.set_row(start_us, 18)
for c, h in enumerate(kw_heads):
    ws1.write(start_us + 1, c, h, f_h2)

us_keywords = [
    ("botero's", 6, 3600, "galeriaduquearango.com/en/blog/all-you-need-to-know-to-understand-fernando-boteros-artworks/"),
    ("ana mercedes hoyos", 6, 3600, "galeriaduquearango.com/artista/ana-mercedes-hoyos/"),
    ("fernando botero", 8, 18100, "galeriaduquearango.com/en/blog/unmistakable-brand-boterism/"),
    ("colombian painters", 3, 720, "galeriaduquearango.com/en/blog/discovering-colombian-art-from-fernando-botero-to-contemporary-masters/"),
    ("plaza botero medellín colombia", 6, 1900, "galeriaduquearango.com/en/blog/plaza-plaza-botero-resignifying-public-space-through-art-in-medellin/"),
    ("oswaldo guayasamin", 7, 1600, "galeriaduquearango.com/en/blog/oswaldo-guayasamins-legacy-to-art/"),
    ("guayasamin", 5, 1300, "galeriaduquearango.com/en/blog/the-works-of-oswaldo-guayasamin/"),
    ("fernando botero artworks", 9, 1300, "galeriaduquearango.com/en/blog/all-you-need-to-know-to-understand-fernando-boteros-artworks/"),
    ("fernando botero style", 1, 170, "galeriaduquearango.com/en/blog/unmistakable-brand-boterism/"),
    ("boterismo", 3, 480, "galeriaduquearango.com/en/blog/unmistakable-brand-boterism/"),
    ("fernando botero art", 8, 3600, "galeriaduquearango.com/en/blog/unmistakable-brand-boterism/"),
    ("nombre de las 30 pinturas más vistas del fernando botero", 2, 210, "galeriaduquearango.com/blog/las-obras-mejor-vendidas-de-fernando-botero/"),
    ("botero", 12, 9900, "galeriaduquearango.com/en/blog/all-you-need-to-know-to-understand-fernando-boteros-artworks/"),
    ("picasso realistic art", 6, 480, "galeriaduquearango.com/en/blog/artistic-periods-of-pablo-picasso/"),
    ("fernando botero paintings", 13, 4400, "galeriaduquearango.com/en/blog/the-works-of-fernando-botero-and-their-significance/"),
    ("mexican art", 24, 8100, "galeriaduquearango.com/en/blog/mexican-art-history-evolution-and-two-key-figures-of-the-twentieth-century/"),
    ("fernando botero still life", 1, 50, "galeriaduquearango.com/en/blog/the-rich-volumes-of-fernando-botero-a-deep-dive-into-his-still-life-paintings/"),
    ("mona lisa age twelve", 7, 480, "galeriaduquearango.com/en/blog/fernando-botero-monalisa-at-12-why-did-he-paint-it/"),
    ("fernando de szyszlo", 5, 480, "galeriaduquearango.com/en/artist/fernando-de-szyszlo/"),
    ("oswaldo guayasamin legacy", 7, 1600, "galeriaduquearango.com/en/blog/oswaldo-guayasamins-legacy-to-art/"),
]
for r, (kw, pos, vol, url) in enumerate(us_keywords):
    fb = f_body if r % 2 == 0 else f_body_alt
    fn = f_num  if r % 2 == 0 else f_num_alt
    fu = f_url  if r % 2 == 0 else f_url_alt
    ws1.write(start_us + 2 + r, 0, kw, fb)
    ws1.write(start_us + 2 + r, 1, pos, fn)
    ws1.write(start_us + 2 + r, 2, vol, fn)
    ws1.write(start_us + 2 + r, 3, url, fu)

# ── Top 10 Páginas CO ──
start_pg = start_us + 24
ws1.merge_range(start_pg, 0, start_pg, 3,
    "Top 10 Páginas por Tráfico Orgánico — Colombia (CO)", f_h1)
ws1.set_row(start_pg, 18)
pg_heads = ["URL de la Página", "Tráfico Estimado/mes", "Nº Keywords", "Keyword Principal"]
for c, h in enumerate(pg_heads):
    ws1.write(start_pg + 1, c, h, f_h2)

top_pages_co = [
    ("galeriaduquearango.com/artista/edgar-negret/", 8782, 27, "edgar negret"),
    ("galeriaduquearango.com/artista/david-manzur/", 2697, 49, "david manzur"),
    ("galeriaduquearango.com/blog/plaza-botero-resignificar-...", 1899, 33, "plaza botero"),
    ("galeriaduquearango.com/blog/las-obras-de-fernando-botero-y-su-significado/", 1473, 138, "obras de fernando botero"),
    ("galeriaduquearango.com/ (homepage)", 808, 109, "galeria duque arango"),
    ("galeriaduquearango.com/artista/alejandro-obregon/", 795, 69, "obras de alejandro obregón"),
    ("galeriaduquearango.com/blog/artistas-colombianos-que-debes-conocer/", 531, 184, "artistas colombianos"),
    ("galeriaduquearango.com/blog/las-obras-mejor-vendidas-de-fernando-botero/", 473, 87, "cuadros de botero"),
    ("galeriaduquearango.com/artista/ana-mercedes-hoyos/", 452, 19, "ana mercedes hoyos obras"),
    ("galeriaduquearango.com/artista/enrique-grau/", 371, 19, "enrique grau obras"),
]
for r, (url, tr, kw_count, kw_top) in enumerate(top_pages_co):
    fb = f_body if r % 2 == 0 else f_body_alt
    fn = f_num  if r % 2 == 0 else f_num_alt
    fu = f_url  if r % 2 == 0 else f_url_alt
    ws1.write(start_pg + 2 + r, 0, url, fu)
    ws1.write(start_pg + 2 + r, 1, tr,  fn)
    ws1.write(start_pg + 2 + r, 2, kw_count, fn)
    ws1.write(start_pg + 2 + r, 3, kw_top, fb)

# ── Top 10 Páginas US ──
start_pg2 = start_pg + 14
ws1.merge_range(start_pg2, 0, start_pg2, 3,
    "Top 10 Páginas por Tráfico Orgánico — Estados Unidos (US)", f_h1)
ws1.set_row(start_pg2, 18)
for c, h in enumerate(pg_heads):
    ws1.write(start_pg2 + 1, c, h, f_h2)

top_pages_us = [
    ("galeriaduquearango.com/artista/ana-mercedes-hoyos/ (*)", 561, 2, "ana mercedes hoyos"),
    ("galeriaduquearango.com/en/blog/all-you-need-to-know-to-understand-fernando-boteros-artworks/", 234, 99, "botero's"),
    ("galeriaduquearango.com/en/blog/unmistakable-brand-boterism/", 218, 74, "boterismo"),
    ("galeriaduquearango.com/en/blog/discovering-colombian-art-from-fernando-botero-to-contemporary-masters/", 169, 76, "colombian painters"),
    ("galeriaduquearango.com/en/blog/the-works-of-oswaldo-guayasamin/", 99, 39, "guayasamin"),
    ("galeriaduquearango.com/en/blog/artistic-periods-of-pablo-picasso/", 93, 151, "picasso realistic art"),
    ("galeriaduquearango.com/en/artist/luis-caballero/", 88, 13, "luis caballero"),
    ("galeriaduquearango.com/en/blog/the-works-of-fernando-botero-and-their-significance/", 75, 67, "fernando botero paintings"),
    ("galeriaduquearango.com/blog/las-obras-de-fernando-botero-y-su-significado/", 61, 33, "obras de fernando botero"),
    ("galeriaduquearango.com/blog/las-obras-mejor-vendidas-de-fernando-botero/", 58, 29, "cuadros de botero"),
]
for r, (url, tr, kw_count, kw_top) in enumerate(top_pages_us):
    fb = f_body if r % 2 == 0 else f_body_alt
    fn = f_num  if r % 2 == 0 else f_num_alt
    fu = f_url  if r % 2 == 0 else f_url_alt
    ws1.write(start_pg2 + 2 + r, 0, url, fu)
    ws1.write(start_pg2 + 2 + r, 1, tr,  fn)
    ws1.write(start_pg2 + 2 + r, 2, kw_count, fn)
    ws1.write(start_pg2 + 2 + r, 3, kw_top, fb)

ws1.write(start_pg2 + 13, 0,
    "(*) Página en español captando tráfico US — oportunidad para crear versión en inglés /en/artist/ana-mercedes-hoyos/",
    f_note)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — BENCHMARK COMPETITIVO
# ══════════════════════════════════════════════════════════════════════════════
ws2 = wb.add_worksheet("Benchmark Competitivo")
ws2.set_zoom(90)
ws2.set_column("A:A", 32)
ws2.set_column("B:B", 8)
ws2.set_column("C:C", 14)
ws2.set_column("D:D", 12)
ws2.set_column("E:E", 16)
ws2.set_column("F:F", 14)
ws2.set_column("G:G", 12)
ws2.set_column("H:H", 16)

ws2.merge_range("A1:H1", "BENCHMARK COMPETITIVO — 17 Galerías | Colombia (CO) y Estados Unidos (US)", f_title)
ws2.merge_range("A2:H2", f"Datos al: mayo 2026  |  Fuente: Semrush  |  Ordenado por Tráfico CO descendente", f_note)
ws2.set_row(0, 22)

headers2 = [
    "Dominio", "Tier",
    "Tráfico CO", "Keywords CO",
    "Ranking Global CO",
    "Tráfico US", "Keywords US",
    "Ranking Global US",
]
for c, h in enumerate(headers2):
    ws2.write(3, c, h, f_h2)

# Data: sorted by CO traffic desc. Duque Arango highlighted separately.
# Format: (domain, tier, traffic_co, kw_co, rank_co, traffic_us, kw_us, rank_us)
benchmark_data = [
    ("galeriaduquearango.com ★", "—", 21243, 2219, 4976,      2215,  1905, 608115),
    ("galeriaelmuseo.com",        "1", 4275,  535,  16716,     28,    44,   6039103),
    ("galerialacometa.com",       "1", 3252,  357,  20476,     521,   138,  1637342),
    ("casasriegner.com",          "1", 1660,  231,  33420,     29,    54,   5977430),
    ("sgr-art.com",               "1", 674,   64,   60849,     0,     3,    24335668),
    ("galeriacasacuadrada.com",   "1", 512,   75,   72119,     0,     2,    27271677),
    ("beatrizesguerra-art.com",   "1", 169,   39,   136779,    137,   93,   3286318),
    ("galeriafreites.com",        "2", 163,   108,  139544,    13,    19,   7545957),
    ("otros360grados.com",        "1", 123,   56,   161710,    0,     9,    17370514),
    ("operagallery.com",          "2", 56,    49,   237163,    8102,  3437, 207074),
    ("galeriaelsapineres.art",    "1", 15,    23,   425793,    0,     0,    0),
    ("artoftheworldgallery.com",  "2", 13,    22,   459856,    1750,  755,  727360),
    ("ascasogallery.com",         "2", 0,     11,   1324190,   319,   156,  2162366),
    ("cernudaarte.com",           "2", 0,     5,    1588126,   574,   404,  1545447),
    ("forumgallery.com",          "2", 0,     5,    1624241,   1924,  1546, 677249),
    ("miguelabreugallery.com",    "2", 0,     3,    2146049,   3332,  441,  439894),
    ("latinartcore.com",          "2", 0,     0,    0,         146,   117,  3189351),
]

for r, row in enumerate(benchmark_data):
    domain, tier, tr_co, kw_co, rk_co, tr_us, kw_us, rk_us = row
    is_ours = "★" in domain

    if is_ours:
        ws2.write(4 + r, 0, domain, f_our_lbl)
        for c, val in enumerate([tier, tr_co, kw_co, rk_co, tr_us, kw_us, rk_us]):
            ws2.write(4 + r, c + 1, val if val != 0 else "Sin datos", f_our_row)
    else:
        fb = f_body if r % 2 == 0 else f_body_alt
        fn = f_num  if r % 2 == 0 else f_num_alt
        ws2.write(4 + r, 0, domain, fb)
        ws2.write(4 + r, 1, tier,   fn)
        for c_off, val in enumerate([tr_co, kw_co, rk_co, tr_us, kw_us, rk_us]):
            disp = val if val > 0 else ("Sin datos" if domain in ["latinartcore.com","galeriaelsapineres.art"] and c_off in [0,1,2] else ("Sin datos" if val == 0 and c_off in [2,5] else 0))
            ws2.write(4 + r, c_off + 2, val if val > 0 else "Sin datos" if str(val)=="0" else val, fn)

# Re-write cleaner
for r, row in enumerate(benchmark_data):
    domain, tier, tr_co, kw_co, rk_co, tr_us, kw_us, rk_us = row
    is_ours = "★" in domain
    bg = ACCENT_GOLD if is_ours else (WHITE if r % 2 == 0 else GRAY_ALT)
    fc = DARK_BLUE   if is_ours else "#000000"
    bl = True        if is_ours else False

    def mf(align="center"):
        return fmt(wb, bold=bl, bg=bg, font_color=fc, size=10, align=align)

    ws2.write(4 + r, 0, domain, mf("left"))
    ws2.write(4 + r, 1, tier, mf())
    ws2.write(4 + r, 2, tr_co if tr_co > 0 else "Sin datos", mf())
    ws2.write(4 + r, 3, kw_co if kw_co > 0 else "Sin datos", mf())
    ws2.write(4 + r, 4, rk_co if rk_co > 0 else "Sin datos", mf())
    ws2.write(4 + r, 5, tr_us if tr_us > 0 else "Sin datos", mf())
    ws2.write(4 + r, 6, kw_us if kw_us > 0 else "Sin datos", mf())
    ws2.write(4 + r, 7, rk_us if rk_us > 0 else "Sin datos", mf())

ws2.write(22, 0,
    "★ Duque Arango destacado en dorado  |  Tier 1 = Comp. directa Colombia  |  Tier 2 = Mercado LATAM/Miami/NY",
    f_note)
ws2.write(23, 0,
    "Ranking Global Semrush: posición más baja = dominio más fuerte en esa base de datos",
    f_note)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — BRECHAS DE KEYWORDS
# ══════════════════════════════════════════════════════════════════════════════
ws3 = wb.add_worksheet("Brechas de Keywords")
ws3.set_zoom(90)
ws3.set_column("A:A", 34)
ws3.set_column("B:B", 14)
ws3.set_column("C:C", 18)
ws3.set_column("D:D", 28)
ws3.set_column("E:E", 18)
ws3.set_column("F:F", 16)
ws3.set_column("G:G", 40)

ws3.merge_range("A1:G1", "BRECHAS DE KEYWORDS — Keywords donde competidores rankean y Duque Arango NO", f_title)
ws3.merge_range("A2:G2",
    "Fuente: Semrush domain_domains | Bases: CO y US | Ordenado por Volumen de Búsqueda descendente",
    f_note)
ws3.set_row(0, 22)

gap_heads = [
    "Keyword", "Volumen de Búsqueda", "Categoría SEO",
    "Competidor que Rankea", "Posición Competidor",
    "Nivel de Oportunidad", "Acción de Contenido Sugerida",
]
for c, h in enumerate(gap_heads):
    ws3.write(3, c, h, f_h2)

# Categorías: Nombre de Artista / Marca / Editorial / Comercial
# Nivel: Alta (>5.000) / Media (1.000–5.000) / Baja (<1.000)
gap_data = [
    # keyword, vol, categoria, competitor, pos_comp, nivel, accion
    ("museo de arte", 8100, "Editorial",
     "galeriaelmuseo.com (CO)", 15, "Alta",
     "Post editorial 'Museos y Galerías de Arte en Colombia': qué ver, cómo visitar, qué comprar"),
    ("wifredo lam", 6600, "Nombre de Artista ⚠",
     "cernudaarte.com (US)", 54, "Alta",
     "Crear/expandir página /en/artist/wifredo-lam/ con obra, legado y disponibilidad — artista en roster"),
    ("beatriz gonzalez", 3600, "Nombre de Artista",
     "casasriegner.com (CO)", 4, "Alta",
     "Post editorial sobre Beatriz González: contexto histórico y relación con el arte colombiano moderno"),
    ("galerias bogotá", 2900, "Comercial",
     "galeriaelmuseo.com pos 11, galerialacometa.com pos 30, casasriegner.com pos 19 (CO)", 11, "Alta",
     "Landing page 'Galerías de Arte en Bogotá' posicionando la sala Bogotá de Duque Arango"),
    ("description of art gallery", 2900, "Editorial",
     "beatrizesguerra-art.com (US)", 34, "Alta",
     "Post en inglés 'What Is an Art Gallery? A Complete Guide' — captura intención informacional US"),
    ("autorretrato", 2400, "Editorial",
     "galeriaelmuseo.com (CO)", 5, "Alta",
     "Post 'El autorretrato en el arte latinoamericano' con obras de artistas del roster"),
    ("jesus abad colorado", 2400, "Nombre de Artista",
     "galeriaelmuseo.com (CO)", 3, "Alta",
     "Post editorial sobre fotografía documental colombiana con mención a fotógrafos del entorno"),
    ("galerias bogota", 1900, "Comercial",
     "galeriaelmuseo.com (CO)", 5, "Alta",
     "Igual que 'galerías bogotá' — misma landing page cubre ambas variantes (con/sin tilde)"),
    ("emma reyes", 1900, "Nombre de Artista",
     "galerialacometa.com (CO)", 4, "Media",
     "Post editorial 'Emma Reyes: la artista colombiana que pintó desde París' — posiciona curador"),
    ("feliza bursztyn", 1900, "Nombre de Artista",
     "galerialacometa.com (CO)", 32, "Media",
     "Post editorial sobre escultoras colombianas del siglo XX (Bursztyn, Negret, etc.)"),
    ("define gallery / what is a gallery", 1600, "Editorial",
     "beatrizesguerra-art.com (US)", 13, "Media",
     "Post en inglés 'What Is an Art Gallery?' con FAQ schema markup para visibilidad en AI"),
    ("pedro ruiz", 1600, "Nombre de Artista",
     "beatrizesguerra-art.com (US)", 8, "Media",
     "Post editorial sobre pintores colombianos de galería que incluya a Pedro Ruiz y artistas del roster"),
    ("beatriz gonzalez artista", 1600, "Nombre de Artista",
     "casasriegner.com (CO)", 3, "Media",
     "Ver acción de 'beatriz gonzalez' — mismo contenido cubre ambas variantes"),
    ("amelia pelaez", 1000, "Nombre de Artista",
     "cernudaarte.com (US)", 6, "Media",
     "Expandir contenido sobre arte cubano moderno en inglés; incluir a Wifredo Lam y artistas caribeños del roster"),
    ("carlos rojas", 1000, "Nombre de Artista",
     "casasriegner.com (CO)", 2, "Media",
     "Post sobre geometrismo y abstracción en Colombia — relacionar con Omar Rayo y Carlos Cruz-Diez del roster"),
]

for r, (kw, vol, cat, comp, pos_c, nivel, accion) in enumerate(gap_data):
    bg = WHITE if r % 2 == 0 else GRAY_ALT
    fb = fmt(wb, bg=bg, size=10)
    fn = fmt(wb, bg=bg, size=10, align="center")
    fw = fmt(wb, bg=bg, size=9, wrap=True)
    ws3.set_row(4 + r, 42)

    if nivel == "Alta":
        f_niv = f_high
    elif nivel == "Media":
        f_niv = f_med
    else:
        f_niv = f_low

    ws3.write(4 + r, 0, kw, fb)
    ws3.write(4 + r, 1, vol, fn)
    ws3.write(4 + r, 2, cat, fn)
    ws3.write(4 + r, 3, comp, fb)
    ws3.write(4 + r, 4, pos_c, fn)
    ws3.write(4 + r, 5, nivel, f_niv)
    ws3.write(4 + r, 6, accion, fw)

ws3.write(20, 0, "Categorías SEO: Nombre de Artista / Editorial / Comercial / Marca  |  "
    "Nivel: Alta >5.000 búsquedas/mes · Media 1.000–5.000 · Baja <1.000  |  "
    "⚠ = artista en roster con amenaza directa de competidor", f_note)
ws3.write(21, 0,
    "Nota: 'museo de arte' (8.100 vol) es la mayor oportunidad editorial sin explotar en CO.",
    f_note)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — PÁGINAS TOP DE COMPETENCIA
# ══════════════════════════════════════════════════════════════════════════════
ws4 = wb.add_worksheet("Páginas Top Competencia")
ws4.set_zoom(90)
ws4.set_column("A:A", 22)
ws4.set_column("B:B", 8)
ws4.set_column("C:C", 54)
ws4.set_column("D:D", 14)
ws4.set_column("E:E", 26)
ws4.set_column("F:F", 20)
ws4.set_column("G:G", 38)

ws4.merge_range("A1:G1",
    "PÁGINAS TOP DE COMPETENCIA — Top 5 páginas de los 3 competidores con más tráfico por mercado",
    f_title)
ws4.merge_range("A2:G2",
    "CO: galeriaelmuseo.com · galerialacometa.com · casasriegner.com  |  "
    "US: operagallery.com · miguelabreugallery.com · forumgallery.com",
    f_note)
ws4.set_row(0, 22)

pg4_heads = [
    "Competidor", "Mercado", "URL de Página",
    "Tráfico Estimado", "Keyword Principal",
    "Tipo de Contenido", "Brecha para Duque Arango",
]
for c, h in enumerate(pg4_heads):
    ws4.write(3, c, h, f_h2)

comp_pages = [
    # CO competitors
    ("galeriaelmuseo.com", "CO", "galeriaelmuseo.com/",
     1476, "galeria el museo", "Página de Inicio / Marca",
     "Nuestra homepage CO tiene 808 vs 1.476 de ellos — fortalecer CTR con schema Organization"),
    ("galeriaelmuseo.com", "CO", "galeriaelmuseo.com/archives/43710/",
     719, "autorretrato", "Post de blog — Artista/Obra",
     "Crear post sobre 'autorretrato en el arte latinoamericano' — keyword 2.400 vol/mes"),
    ("galeriaelmuseo.com", "CO", "galeriaelmuseo.com/archives/2084/",
     503, "jesus abad colorado", "Post de blog — Artista",
     "Post editorial sobre fotografía documental colombiana para capturar búsquedas editoriales"),
    ("galeriaelmuseo.com", "CO", "galeriaelmuseo.com/archives/450/",
     303, "galería el museo", "Archivo de exposiciones",
     "Crear sección de exposiciones pasadas indexable con keywords de artistas representados"),
    ("galeriaelmuseo.com", "CO", "galeriaelmuseo.com/archives/26977/",
     196, "jacanamijoy", "Post de blog — Artista",
     "Artistas colombianos contemporáneos — reforzar /blog/artistas-colombianos-que-debes-conocer/"),
    ("galerialacometa.com", "CO", "galerialacometa.com/",
     1860, "galeria la cometa", "Página de Inicio / Marca",
     "Su homepage supera a la nuestra en CO — fortalecer con schema Organization y local SEO"),
    ("galerialacometa.com", "CO", "galerialacometa.com/exhibiciones/bogota/gabriela-pinilla-clandestina-es",
     182, "exhibición bogotá arte", "Página de Exposición Individual",
     "Crear páginas de exposición individual indexables con keywords de artistas y temas"),
    ("galerialacometa.com", "CO", "galerialacometa.com/artistas/miguel-angel-rojas-es",
     86, "miguel ángel rojas", "Página de Artista",
     "Nuestras páginas de artistas dominan CO — continuar modelo con artistas con menos contenido"),
    ("galerialacometa.com", "CO", "galerialacometa.com/exhibiciones/bogota/emma-reyes-las-caras-de-emma-reyes-es",
     83, "emma reyes", "Exposición / Nombre de Artista",
     "Post editorial sobre Emma Reyes captura búsquedas de artistas históricos colombianos"),
    ("galerialacometa.com", "CO", "galerialacometa.com/artistas/miguel-angel-rojas-es",
     80, "la chiqui m19", "Exposición Histórica",
     "Exposiciones con contexto histórico colombiano generan tráfico — oportunidad editorial"),
    ("casasriegner.com", "CO", "casasriegner.com/",
     994, "casas riegner", "Página de Inicio / Marca",
     "Su homepage es fuerte en CO — diferencial: nuestra artista Edgar Negret supera todo su tráfico"),
    ("casasriegner.com", "CO", "casasriegner.com/artistas/beatriz-gonzalez",
     179, "beatriz gonzalez", "Página de Artista",
     "Beatriz González (3.600 vol/mes) — crear contenido editorial sobre ella para competir"),
    ("casasriegner.com", "CO", "casasriegner.com/artistas/carlos-rojas",
     116, "carlos rojas", "Página de Artista",
     "Carlos Rojas (1.000 vol/mes) — post sobre geometrismo vinculando con Omar Rayo y Cruz-Diez"),
    ("casasriegner.com", "CO", "casasriegner.com/artistas/antonio-caro",
     66, "antonio caro", "Página de Artista",
     "Arte conceptual colombiano — post editorial vinculado a artistas modernos del roster"),
    ("casasriegner.com", "CO", "casasriegner.com/artistas/rosemberg-sandoval",
     61, "rosemberg sandoval", "Página de Artista",
     "Arte performativo colombiano — diferenciación de contenido editorial por movimiento artístico"),
    # US competitors
    ("operagallery.com", "US", "operagallery.com/",
     1384, "opera gallery", "Página de Inicio / Marca",
     "Opera Gallery domina US — fortalecer homepage en inglés con schema Organization y local NY/Miami"),
    ("operagallery.com", "US", "operagallery.com/artist/keith-haring",
     849, "keith haring", "Página de Artista — Alta Demanda",
     "Artistas con alta demanda global generan enorme tráfico US — replicar con Fernando Botero página en inglés"),
    ("operagallery.com", "US", "operagallery.com/artist/bernard-buffet",
     611, "bernard buffet", "Página de Artista — Internacional",
     "Artistas europeos en galería latinoamericana capturan tráfico US — Igor Mitoraj y Jaume Plensa en roster"),
    ("operagallery.com", "US", "operagallery.com/viewing-rooms/botero-2023",
     476, "botero viewing room", "Viewing Room Virtual — Botero",
     "Viewing rooms de Botero generan 476+452 visitas US — crear /en/viewing-room/botero/ como página dedicada"),
    ("operagallery.com", "US", "operagallery.com/viewing-rooms/fernando-botero",
     452, "fernando botero gallery", "Viewing Room Virtual — Botero",
     "Duplicar estrategia: landing page en inglés de Botero con viewing room y obras disponibles"),
    ("miguelabreugallery.com", "US", "miguelabreugallery.com/",
     2402, "miguel abreu gallery", "Página de Inicio / Marca",
     "72% de su tráfico en homepage — altamente dependiente de marca; aprovechar con artista pages"),
    ("miguelabreugallery.com", "US", "miguelabreugallery.com/artists/quaytman/",
     158, "quaytman artist", "Página de Artista Contemporáneo",
     "Artistas contemporáneos menos conocidos pueden rankear — crear páginas /en/artist/ para todos los contemporáneos"),
    ("miguelabreugallery.com", "US", "miguelabreugallery.com/artists/francois-marie-banier/",
     142, "francois marie banier", "Página de Artista Internacional",
     "Artistas internacionales en roster (Mitoraj, Plensa, Rondinone) tienen potencial US sin explotar"),
    ("forumgallery.com", "US", "forumgallery.com/",
     864, "forum gallery", "Página de Inicio / Marca",
     "Homepage fuerte — diferenciarnos con contenido de artistas latinoamericanos únicos"),
    ("forumgallery.com", "US", "forumgallery.com/artists/holly-lane/biography",
     188, "holly lane artist biography", "Página de Biografía de Artista",
     "Páginas de BIOGRAPHY generan tráfico US — agregar sección /biography a páginas de artistas en inglés"),
    ("forumgallery.com", "US", "forumgallery.com/artists/claudio-bravo/videos",
     114, "claudio bravo videos", "Página de Videos de Artista",
     "Contenido de VIDEO por artista rankea en US — crear sección /videos o /works en páginas de artistas"),
    ("forumgallery.com", "US", "forumgallery.com/artists/norman-rockwell/biography",
     100, "norman rockwell biography", "Biografía de Artista Clásico",
     "Artistas históricos con páginas de biografía detallada dominan en US — reproducir con artistas modernos del roster"),
]

for r, (comp, mkt, url, tr, kw_top, tipo, brecha) in enumerate(comp_pages):
    bg = WHITE if r % 2 == 0 else GRAY_ALT
    fb = fmt(wb, bg=bg, size=10, bold=True)
    fm = fmt(wb, bg=bg, size=10, align="center")
    fu2= fmt(wb, bg=bg, size=9, font_color="#1155CC", italic=True)
    fn2= fmt(wb, bg=bg, size=10, align="center")
    fw2= fmt(wb, bg=bg, size=9, wrap=True)
    ws4.set_row(4 + r, 40)
    ws4.write(4 + r, 0, comp,   fb)
    ws4.write(4 + r, 1, mkt,    fm)
    ws4.write(4 + r, 2, url,    fu2)
    ws4.write(4 + r, 3, tr,     fn2)
    ws4.write(4 + r, 4, kw_top, fn2)
    ws4.write(4 + r, 5, tipo,   fw2)
    ws4.write(4 + r, 6, brecha, fw2)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — INSIGHTS Y RECOMENDACIONES
# ══════════════════════════════════════════════════════════════════════════════
ws5 = wb.add_worksheet("Insights y Recomendaciones")
ws5.set_zoom(90)
ws5.set_column("A:A", 36)
ws5.set_column("B:B", 36)
ws5.set_column("C:C", 28)
ws5.set_column("D:D", 40)
ws5.set_column("E:E", 10)

ws5.merge_range("A1:E1",
    "INSIGHTS Y RECOMENDACIONES — Semana del 2026-06-04",
    f_title)
ws5.merge_range("A2:E2",
    "Cada insight cita el framework de skill que lo respalda. Prioridad: Alta / Media / Baja.",
    f_note)
ws5.set_row(0, 22)

ins_heads = [
    "Insight", "Evidencia (datos Semrush)", "Framework Aplicado",
    "Acción Recomendada", "Prioridad",
]
for c, h in enumerate(ins_heads):
    ws5.write(3, c, h, f_h2)

insights = [
    (
        "Liderazgo absoluto en Colombia: Duque Arango es #1 de 17 galerías en CO con 21.243 visitas/mes orgánicas",
        "Tráfico CO: Duque Arango 21.243 vs. 2º lugar galeriaelmuseo.com 4.275 (5x de ventaja). Keywords CO: 2.219 vs. máximo competidor 535.",
        "seo-audit: E-E-A-T Signals — Authoritativeness demostrada en resultados CO",
        "Mantener cadencia de publicación de páginas de artistas. Replicar el modelo Edgar Negret (41% del tráfico CO desde una sola página) en artistas del roster con menos contenido: Gustavo Vélez, Tomás Ochoa, Carlos Salas.",
        "Alta",
    ),
    (
        "Edgar Negret: dependencia crítica — una sola página genera el 41% del tráfico orgánico CO",
        "Página /artista/edgar-negret/ = 8.782 visitas/mes CO. Segunda página más visitada: /artista/david-manzur/ = 2.697. Concentración de riesgo alta.",
        "seo-audit: Keyword Targeting — No keyword cannibalization; content-strategy: Pillar 1 diversificación",
        "Publicar en los próximos 30 días páginas de artista para: Alejandra Aristizábal, Reynier Ferrer, Nadín Ospina y Álvaro Barrios. Formato: misma estructura que Edgar Negret (obras, biografía, contexto, FAQ).",
        "Alta",
    ),
    (
        "Wifredo Lam (6.600 búsquedas/mes US): competidor cernudaarte.com amenaza posición en artista de nuestro roster",
        "cernudaarte.com rankea pos. 54 para 'wifredo lam' en US (6.600 vol/mes). galeriaduquearango.com no aparece en ninguna posición para este keyword. Wifredo Lam está en el roster.",
        "competitor-alternatives: Research Process — Threat identification; seo-audit: Keyword Targeting — Artist name gaps",
        "Crear o expandir /en/artist/wifredo-lam/ con: definición extractable del artista en primer párrafo, estadísticas de ventas en subasta, tabla comparativa de períodos, FAQ schema, y sección de disponibilidad de obras. Publicar antes de 2 semanas.",
        "Alta",
    ),
    (
        "'Galerías Bogotá' (2.900 búsquedas/mes CO): brecha comercial sin explotar — 3 competidores rankean, Duque Arango no",
        "casasriegner.com pos 19, galerialacometa.com pos 30, galeriaelmuseo.com pos 11 para 'galerias bogotá'. Duque Arango ausente. Intención: COMERCIAL con alta conversión.",
        "content-strategy: Keyword Research by Buyer Stage — Decision Stage (modifiers: 'galerías', 'bogotá'); competitor-alternatives: SEO Considerations",
        "Crear landing page /galerias-bogota/ con: descripción de la sala Bogotá, artistas expuestos, horarios, dirección y FAQ local. Incluir schema LocalBusiness. También captura la variante 'galerias bogota' (1.900 vol) sin tilde.",
        "Alta",
    ),
    (
        "'Museo de arte' (8.100 búsquedas/mes CO): mayor oportunidad editorial — galeriaelmuseo.com rankea pos. 15, Duque Arango ausente",
        "galeriaelmuseo.com en pos. 15 para 'museo de arte' (8.100 vol/mes CO). Tráfico potencial estimado: 200–400 visitas/mes si se alcanza top 10. Intención: informacional.",
        "ai-seo: Pillar 1 Structure — Content Extractability (definition blocks, comparison tables); content-strategy: Hub and Spoke",
        "Publicar post 'Museos y Galerías de Arte en Colombia: guía completa' con definiciones extractables, tabla comparativa de instituciones, sección de galerías privadas destacando Duque Arango, y FAQ schema. Target: AI Overviews + tráfico orgánico CO.",
        "Alta",
    ),
    (
        "Ana Mercedes Hoyos genera el 25% del tráfico US desde una página en ESPAÑOL — sin versión en inglés",
        "Página /artista/ana-mercedes-hoyos/ (en español) = 561 visitas/mes US con solo 2 keywords indexados. Enorme potencial si existe versión /en/artist/ con contenido profundo en inglés.",
        "ai-seo: Pillar 1 — Content Extractability Check; seo-audit: On-Page SEO — International SEO & hreflang",
        "Crear /en/artist/ana-mercedes-hoyos/ con: párrafo definitorio en inglés (40-60 palabras), estadísticas de sus obras más vendidas, comparación con artistas similares (tabla), FAQ, y sección de disponibilidad. Enlazar desde página en español con hreflang. Prioridad máxima en US.",
        "Alta",
    ),
    (
        "Opera Gallery domina US con viewing rooms de Botero: 476+452 visitas/mes desde páginas dedicadas",
        "operagallery.com/viewing-rooms/botero-2023 = 476 visitas; /viewing-rooms/fernando-botero = 452 visitas. Botero genera tráfico US de alta intención sin que Duque Arango lo capte directamente.",
        "competitor-alternatives: Format 3 You vs Competitor — positioning; ai-seo: Pillar 2 Authority — citable statistics",
        "Crear /en/artist/fernando-botero/ o /en/viewing-room/botero/ como página dedicada en inglés, con: contexto del mercado secundario de Botero, obras históricas, precios de subasta documentados (con fuente), y disponibilidad en galería. Schema ArtGallery + ArtWork.",
        "Media",
    ),
    (
        "Forum Gallery captura tráfico US con páginas de BIOGRAPHY y VIDEOS de artistas — formato ausente en Duque Arango",
        "forumgallery.com/artists/holly-lane/biography = 188 visitas; /artists/claudio-bravo/videos = 114 visitas. Formato de contenido multimedia por artista genera tráfico orgánico en US sin competir directamente.",
        "ai-seo: Content Types That Get Cited Most — Definitive guides (15% citation share); content-strategy: Content Pillars",
        "Agregar a cada página /en/artist/ una sección de Biography estructurada (nacimiento, formación, estilo, impacto) y, donde exista contenido audiovisual, una sección /videos o /works. Empezar con Julio Larraz, Óscar Murillo y Olga de Amaral — artistas con mayor demanda US.",
        "Media",
    ),
]

for r, (insight, evidencia, framework, accion, prioridad) in enumerate(insights):
    bg = WHITE if r % 2 == 0 else GRAY_ALT
    fw3 = fmt(wb, bg=bg, size=9, wrap=True)
    fw3b= fmt(wb, bg=bg, size=9, wrap=True, bold=True)
    if prioridad == "Alta":
        f_p = f_high
    elif prioridad == "Media":
        f_p = f_med
    else:
        f_p = f_low
    ws5.set_row(4 + r, 90)
    ws5.write(4 + r, 0, insight,   fw3b)
    ws5.write(4 + r, 1, evidencia, fw3)
    ws5.write(4 + r, 2, framework, fw3)
    ws5.write(4 + r, 3, accion,    fw3)
    ws5.write(4 + r, 4, prioridad, f_p)

ws5.write(13, 0,
    "Frameworks: seo-audit v1.2 · ai-seo v1.2 · content-strategy v1.1 · competitor-alternatives v1.1",
    f_note)

wb.close()

size = os.path.getsize(FILEPATH)
print(f"✓ Archivo creado: {FILEPATH}")
print(f"✓ Tamaño: {size:,} bytes")
