#!/usr/bin/env python3
"""Build Gallery Intelligence Excel for 2026-06-15."""

import xlsxwriter
import os

OUTPUT = "/home/user/da-routines/Gallery_Intelligence_2026-06-15.xlsx"

workbook = xlsxwriter.Workbook(OUTPUT)

# ── Formats ─────────────────────────────────────────────────────────────────
hdr = workbook.add_format({
    "bold": True, "bg_color": "#1A1A2E", "font_color": "#FFFFFF",
    "border": 1, "text_wrap": True, "valign": "vcenter"
})
hdr2 = workbook.add_format({
    "bold": True, "bg_color": "#16213E", "font_color": "#FFFFFF",
    "border": 1, "text_wrap": True, "valign": "vcenter"
})
cell = workbook.add_format({
    "border": 1, "text_wrap": True, "valign": "top"
})
cell_nuevo = workbook.add_format({
    "border": 1, "text_wrap": True, "valign": "top",
    "bg_color": "#E8F5E9"
})
cell_continua = workbook.add_format({
    "border": 1, "text_wrap": True, "valign": "top",
    "bg_color": "#FFF9C4"
})
cell_cambio = workbook.add_format({
    "border": 1, "text_wrap": True, "valign": "top",
    "bg_color": "#FFF3E0"
})
cell_alert = workbook.add_format({
    "border": 1, "text_wrap": True, "valign": "top",
    "bg_color": "#FFEBEE", "font_color": "#B71C1C", "bold": True
})
url_fmt = workbook.add_format({
    "border": 1, "text_wrap": True, "valign": "top",
    "font_color": "#1565C0", "underline": True
})

# ── TAB 1: Changes & Findings ────────────────────────────────────────────────
ws1 = workbook.add_worksheet("Hallazgos y cambios")

cols = [
    "Galería", "Tier", "Sección", "Tipo de cambio", "Estado",
    "Descripción del hallazgo", "Nota estratégica para Duque Arango", "URL"
]
col_widths = [22, 6, 18, 22, 12, 60, 55, 40]

for i, (name, w) in enumerate(zip(cols, col_widths)):
    ws1.write(0, i, name, hdr)
    ws1.set_column(i, i, w)

ws1.set_row(0, 30)
ws1.freeze_panes(1, 0)

# Data rows — each is (gallery, tier, section, type, status, description, strategic_note, url, fmt_key)
# fmt_key: "nuevo", "continua", "cambio", "alert", "normal"

rows = [
    # ─── TIER 1 ───────────────────────────────────────────────────────────────
    (
        "Galería La Cometa", "1",
        "Exposiciones", "New Exhibition",
        "CONTINÚA",
        "Cuatro exposiciones activas en Madrid y Medellín (todas hasta el 5 de julio):\n"
        "• Madrid: 'Parar el Mundo' – Adam Goldstein (9 mayo–5 jul)\n"
        "• Madrid: 'Pensamiento mágico / El año entrante' – Alejandro Ospina (9 mayo–5 jul)\n"
        "• Medellín: 'Envoltorios' – Asicaz Monzón (14 mayo–5 jul)\n"
        "• Medellín: 'Nada Es Lo Que Parece, Al Parecer Desaparece II' – Luisa Aristizábal (14 mayo–5 jul)",
        "La Cometa opera con programa activo en dos ciudades simultáneamente. "
        "Duque Arango puede diferenciarse con contenido editorial sobre sus artistas propios "
        "mientras La Cometa no produce nuevas aperturas esta semana.",
        "https://galerialacometa.com",
        "continua"
    ),
    (
        "Galería Casa Cuadrada", "1",
        "General", "Operational",
        "—",
        "No se detectó actividad nueva esta semana.",
        "Sin señal competitiva directa. Galería en modo mantenimiento o transición de programa.",
        "http://galeriacasacuadrada.com",
        "normal"
    ),
    (
        "Casas Riegner", "1",
        "Ferias de arte", "Fair Announcement",
        "CONTINÚA",
        "Art Basel Basilea 2026 — Preview VIP mañana, 16 de junio; días públicos 18–21 de junio. "
        "Casas Riegner es la única galería colombiana en el sector Galleries del evento más importante "
        "del circuito A internacional (290 expositores de 43 países). Presencia confirmada desde semanas previas.",
        "🚨 Esta semana es el momento de mayor visibilidad de Casas Riegner en 2026. "
        "Duque Arango debe activar comunicaciones con coleccionistas internacionales "
        "y producir contenido editorial que aproveche el contexto de la feria aunque no esté presente en ella.",
        "https://www.artbasel.com/catalog/gallery/1082/Casas-Riegner",
        "alert"
    ),
    (
        "Galería El Museo", "1",
        "General", "Operational",
        "—",
        "No se encontró actividad nueva en el período 8–15 de junio de 2026. "
        "El sitio web lista exposiciones anteriores.",
        "Competidor en silencio esta semana. Ventana para Duque Arango de generar presencia digital.",
        "https://www.galeriaelmuseo.com",
        "normal"
    ),
    # ─── TIER 2 ───────────────────────────────────────────────────────────────
    (
        "Latin Art Core", "2",
        "Exposiciones", "New Exhibition",
        "NUEVO",
        "'The Optical Experience' — abrió el 12 de junio de 2026 en Miami (Little Havana / Calle Ocho). "
        "La muestra explora el movimiento del Arte Concreto cubano del siglo XX (grupo 10 Pintores Concretos, "
        "1958–1961) y su legado en el arte óptico latinoamericano. "
        "Obras con títulos como 'Circulos' (1971) y 'Fragment of a Symphony #2' (1970).",
        "El movimiento cinético-óptico latinoamericano recibe atención institucional simultáneamente con "
        "la retrospectiva de Le Parc en Tate y la presencia de Olga de Amaral en Art Basel. "
        "Duque Arango, que tiene a Carlos Cruz-Diez en su roster, puede publicar contenido editorial "
        "sobre este legado y posicionarse como referente en el tema.",
        "https://latinartcore.com/exhibition/the-optical-experience/",
        "nuevo"
    ),
    (
        "Art of the World Gallery", "2",
        "General", "Operational",
        "—",
        "La exposición 'Mimetism: Echoes that Breathe' de Karla de Lara cerró el 6 de junio. "
        "No se detectó nuevo programa para la semana del 8–15 de junio.",
        "Sin actividad nueva. La galería comparte con Duque Arango el mercado de coleccionistas latinoamericanos en Houston.",
        "https://www.artoftheworldgallery.com",
        "normal"
    ),
    (
        "Ascaso Gallery", "2",
        "Exposiciones / Roster", "New Exhibition + Institutional",
        "CONTINÚA",
        "Ascaso Gallery anuncia 'Reviver', exposición individual de Andrew Hem (Los Ángeles), apertura 30 de junio. "
        "La galería continúa representando activamente a Julio Larraz, Fernando Botero y Carlos Cruz-Diez "
        "en el mercado estadounidense — amenaza de roster confirmada desde semana del 1 de junio.",
        "🚨 3ª semana consecutiva sin respuesta visible de Duque Arango ante la competencia de Ascaso "
        "sobre artistas compartidos. Prioridad: reforzar contenido SEO y presencia digital para "
        "Botero, Cruz-Diez y Larraz antes del segundo semestre.",
        "https://www.ascasogallery.com",
        "alert"
    ),
    (
        "Galería Freites", "2",
        "General", "Operational",
        "—",
        "No se detectó actividad nueva esta semana en el período 8–15 de junio.",
        "Galería venezolana con sedes en Caracas, Miami y Madrid. Sin señal competitiva directa esta semana.",
        "https://galeriafreites.com",
        "normal"
    ),
    (
        "Opera Gallery", "2",
        "Exposiciones", "New Exhibition",
        "CONTINÚA",
        "Exposiciones activas en múltiples sedes globales (todas iniciadas antes del 8 de junio):\n"
        "• Londres: 'Pieter Obels | Feng Xiao-Min' (4 jun–5 jul)\n"
        "• París: 'Regards sur l'art espagnol' (22 may–17 jun)\n"
        "• Singapur: (22 may–7 jul)\n"
        "• Madrid: (8 may–20 jun)\n"
        "Sin lanzamientos nuevos esta semana.",
        "Opera Gallery mantiene presencia global activa pero sin novedad directa sobre artistas del roster "
        "de Duque Arango esta semana. Su espacio en Houston (abierto marzo 2026) sigue activo.",
        "https://www.operagallery.com",
        "continua"
    ),
    # ─── TIER 3 ───────────────────────────────────────────────────────────────
    (
        "Gagosian", "3",
        "General", "Operational",
        "—",
        "Sin actividad relacionada con artistas latinoamericanos detectada esta semana. "
        "Exhibitions activas: Anselm Kiefer (hasta 27 jun), Frank Gehry (hasta 27 jun). "
        "Summer 2026 issue de Gagosian Quarterly disponible.",
        "Sin impacto directo sobre artistas del roster de Duque Arango esta semana.",
        "https://gagosian.com",
        "normal"
    ),
    (
        "David Zwirner", "3",
        "Exposiciones", "New Exhibition",
        "CONTINÚA",
        "Óscar Murillo — 'Collective Osmosis' (con Claude Monet). DAS MINSK Kunsthaus y Museum Barberini, "
        "Potsdam, Alemania. Sigue en curso hasta agosto 2026. "
        "Primera reportada semana del 1 de junio.",
        "Óscar Murillo (artista colombiano del roster de Duque Arango) mantiene visibilidad institucional "
        "sostenida en Europa. Duque Arango puede publicar contenido que amplíe esta narrativa.",
        "https://www.davidzwirner.com/artists/oscar-murillo",
        "continua"
    ),
    (
        "Hauser & Wirth", "3",
        "Ferias de arte", "Fair Announcement",
        "CONTINÚA",
        "Hauser & Wirth participa en Art Basel Basel 2026 (18–21 jun) con Louise Bourgeois, "
        "Picasso, Philip Guston, Amy Sherald entre otros. Forman parte de 'Basel Exclusive'. "
        "Sin artistas latinoamericanos específicos identificados en su booth esta semana.",
        "Sin impacto directo sobre artistas del roster de Duque Arango.",
        "https://www.hauserwirth.com",
        "normal"
    ),
    (
        "Galerie Lelong", "3",
        "Exposiciones", "New Exhibition",
        "CONTINÚA",
        "Lucia Laguna (Río de Janeiro, 1966) — 'Apenas meus cabelos são brancos..' "
        "[Only my hair is white…]. Galerie Lelong & Co., Nueva York. 14 mayo–27 junio 2026. "
        "Pinturas de las series 'Pequenos formatos' y 'Paisagem'.",
        "Artista brasileña contemporánea con galería de primer nivel en NY. "
        "Recordatorio del potencial de posicionamiento internacional de artistas latinoamericanos "
        "del calibre de las artistas del roster de Duque Arango.",
        "https://galerielelong.com/exhibitions/",
        "continua"
    ),
    (
        "Lisson Gallery", "3",
        "Ferias de arte / Artista roster", "Fair Announcement",
        "CONTINÚA ↑ PICO",
        "Art Basel Basel 2026 — VIP Preview MAÑANA (16 jun), público 18–21 jun.\n"
        "Olga de Amaral incluida en el booth principal de Lisson con obras nuevas e históricas, "
        "incluyendo 'Rojo y oro' (2016). También participan Carmen Herrera, Hélio Oiticica, Dalton Paula.\n"
        "Arte Unlimited: Ryan Gander y Wael Shawky.\n"
        "NOTA: Art Basel Qatar 2026 con solo show de Olga de Amaral ya ocurrió (feb 2026).",
        "🚨 Esta es la semana de MÁXIMA visibilidad de Olga de Amaral en el mercado internacional en 2026. "
        "Duque Arango debe publicar contenido editorial sobre de Amaral HOY o mañana para aprovechar "
        "el pico de búsquedas e interés mediático generado por Art Basel.",
        "https://www.lissongallery.com/news/lisson-at-art-basel-2026",
        "alert"
    ),
    (
        "Lehmann Maupin", "3",
        "Exposiciones / Ferias", "New Exhibition",
        "CAMBIÓ",
        "OSGEMEOS (São Paulo) — 'The Open Window', Nueva York. Cerró el 6 de junio (antes del período reportado).\n"
        "Art Basel Basel 2026 (18–21 jun): Lehmann Maupin participa pero sin artistas latinoamericanos "
        "específicos identificados para su booth en Basel.\n"
        "Art Basel París: presentarán OSGEMEOS y Cecilia Vicuña (chilena) — feria diferente.",
        "OSGEMEOS (artistas brasileños) tuvo exposición de alto perfil en Nueva York que acaba de cerrar. "
        "Lehmann Maupin consolida su programa latinoamericano con Cecilia Vicuña para Art Basel París.",
        "https://www.lehmannmaupin.com",
        "cambio"
    ),
    (
        "Perrotin", "3",
        "Exposiciones", "New Exhibition",
        "CONTINÚA",
        "Gabriel Rico (México) — 'Gabrielinos (I Am You And What I See Is Me)'. "
        "Perrotin Los Ángeles. 6 junio–11 julio 2026. "
        "También: JR 'Les esquisses de la Caverne' en París Marais (5 jun–25 jul).",
        "Gabriel Rico es artista mexicano en galería de primer nivel en LA. "
        "Sin impacto directo sobre artistas del roster de Duque Arango.",
        "https://www.perrotin.com",
        "continua"
    ),
    # ─── BONUS: Le Parc update ────────────────────────────────────────────────
    (
        "Tate Modern (referencia)", "3",
        "Institucional / Artista roster", "Institutional",
        "CAMBIÓ → ABIERTO",
        "Julio Le Parc: 'Light. Colour. Action.' — RETROSPECTIVA INAUGURADA EL 11 DE JUNIO DE 2026.\n"
        "Más de 60 obras de 70 años de trayectoria. Corre hasta el 3 de mayo de 2027.\n"
        "Primera retrospectiva de un museo británico para el artista. "
        "Organizada en colaboración con el artista y su Atelier antes de su fallecimiento (mayo 2026, 97 años).",
        "🚨 La ventana de mercado secundario para Julio Le Parc (roster Duque Arango) está ahora ACTIVA. "
        "La retrospectiva en Tate legitimará precios y aumentará demanda. "
        "El comunicado de homenaje recomendado la semana pasada sigue siendo urgente. "
        "Revisar inventario de obras de Le Parc para posicionamiento comercial.",
        "https://www.tate.org.uk/whats-on/tate-modern/julio-le-parc",
        "alert"
    ),
]

fmt_map = {
    "nuevo": cell_nuevo,
    "continua": cell_continua,
    "cambio": cell_cambio,
    "alert": cell_alert,
    "normal": cell,
}

for r_idx, row in enumerate(rows, start=1):
    gallery, tier, section, change_type, status, desc, note, url, fmt_key = row
    fmt = fmt_map[fmt_key]
    ws1.write(r_idx, 0, gallery, fmt)
    ws1.write(r_idx, 1, tier, fmt)
    ws1.write(r_idx, 2, section, fmt)
    ws1.write(r_idx, 3, change_type, fmt)
    ws1.write(r_idx, 4, status, fmt)
    ws1.write(r_idx, 5, desc, fmt)
    ws1.write(r_idx, 6, note, fmt)
    ws1.write(r_idx, 7, url, url_fmt)
    ws1.set_row(r_idx, 100)

# Legend
ws1.write(len(rows) + 2, 0, "LEYENDA DE COLORES:", hdr)
ws1.write(len(rows) + 3, 0, "Verde claro = NUEVO (no reportado en últimas 4 semanas)", cell_nuevo)
ws1.write(len(rows) + 4, 0, "Amarillo = CONTINÚA (ya reportado, sigue activo)", cell_continua)
ws1.write(len(rows) + 5, 0, "Naranja claro = CAMBIÓ (estaba presente, evolucionó)", cell_cambio)
ws1.write(len(rows) + 6, 0, "Rojo / Alerta = Requiere acción inmediata", cell_alert)

# ── TAB 2: Cross-Gallery Patterns ────────────────────────────────────────────
ws2 = workbook.add_worksheet("Patrones cruzados")

cols2 = [
    "Patrón observado", "Galerías involucradas",
    "Lo que puede señalar", "Acción recomendada para Duque Arango"
]
col_widths2 = [35, 30, 45, 55]

for i, (name, w) in enumerate(zip(cols2, col_widths2)):
    ws2.write(0, i, name, hdr2)
    ws2.set_column(i, i, w)

ws2.set_row(0, 30)
ws2.freeze_panes(1, 0)

patterns = [
    (
        "Art Basel Basel como eje central de la semana — múltiples galerías presentes simultáneamente",
        "Casas Riegner (T1), Lisson Gallery (T3), Hauser & Wirth (T3), Lehmann Maupin (T3)",
        "La semana del 16–21 de junio concentra el mayor flujo de compradores, prensa y coleccionistas "
        "del año. Quienes no están en la feria quedan en silencio mediático. Duque Arango no está presente.",
        "Producir contenido editorial propio esta semana (newsletter, redes, website) sobre artistas del roster "
        "conectados al ecosistema de Art Basel (Olga de Amaral, Le Parc). "
        "Aprovechar el tráfico de búsquedas sobre arte que se dispara cada junio."
    ),
    (
        "Olga de Amaral: campaña de posicionamiento multi-fair completando su arco anual",
        "Lisson Gallery (Art Basel HK, Qatar, Basel) — 3 ferias en 2026",
        "Lisson Gallery ejecuta una estrategia sistemática de posicionamiento de de Amaral "
        "a través de tres ediciones de Art Basel en un mismo año — estrategia sin precedentes "
        "para la artista. Esto maximiza visibilidad, precios y legitimidad institucional.",
        "Esta semana es el pico: publicar hoy/mañana contenido editorial sobre Olga de Amaral "
        "para Duque Arango. Si la galería tiene inventario de obras de de Amaral, activar CTA. "
        "Posicionar a Duque Arango como referente en Colombia sobre su trayectoria."
    ),
    (
        "Renacimiento institucional del Arte Cinético/Op Art latinoamericano",
        "Latin Art Core (exposición 12 jun), Tate Modern (Le Parc retro), Ascaso Gallery (Cruz-Diez)",
        "Tres señales simultáneas en la misma semana convergen alrededor del movimiento cinético "
        "y óptico latinoamericano. Este micro-ciclo de atención institucional es temporal "
        "pero puede generar demanda de coleccionistas y cobertura editorial.",
        "Duque Arango tiene a Carlos Cruz-Diez y Julio Le Parc en su roster. "
        "Publicar un artículo sobre el legado del arte cinético latinoamericano esta semana "
        "captura tráfico real y refuerza autoridad editorial en el tema."
    ),
    (
        "Ascaso Gallery: presión competitiva sostenida sobre artistas compartidos (semana 3)",
        "Ascaso Gallery vs. Duque Arango (Botero, Cruz-Diez, Larraz)",
        "Ascaso Gallery mantiene representación activa de tres artistas del roster de Duque Arango "
        "en el mercado estadounidense durante al menos 3 semanas consecutivas sin respuesta SEO visible "
        "de Duque Arango. La brecha digital puede traducirse en pérdida de leads de coleccionistas.",
        "Prioridad alta: crear o actualizar páginas de artista para Botero, Cruz-Diez y Larraz "
        "en galeriaduquearango.com. Reforzar SEO local en búsquedas de Miami y NY. "
        "Publicar contenido de autoridad (ensayos, novedades de mercado) sobre estos artistas."
    ),
    (
        "Galerías colombianas Tier 1: brecha creciente entre Casas Riegner y el resto",
        "Casas Riegner (Art Basel Basel), La Cometa, Casa Cuadrada, El Museo, Duque Arango",
        "Por cuarta semana consecutiva, Casas Riegner opera en un nivel diferente al de los demás. "
        "Esta semana está en la feria más importante del mundo mientras el resto mantiene programas locales. "
        "La brecha de presencia internacional se está institucionalizando.",
        "Considerar estrategia de feria para segundo semestre 2026 o 2027. "
        "Mientras tanto, competir en el territorio digital: Duque Arango puede tener mayor "
        "autoridad de contenido en artistas del roster que Casas Riegner no representa."
    ),
    (
        "Julio Le Parc: ventana de mercado secundario activa",
        "Tate Modern (Londres), Duque Arango (roster)",
        "La retrospectiva en Tate Modern (inaugurada 11 jun, corre hasta mayo 2027) convierte "
        "a Le Parc en el artista latinoamericano más visible institucionalmente en Europa en este momento. "
        "Las retrospectivas en Tate tienen históricamente un efecto directo sobre el mercado secundario.",
        "Revisar inventario de obras de Julio Le Parc. Publicar comunicado de homenaje/editorial "
        "que posicione a Duque Arango como galería de referencia para su mercado en Colombia. "
        "Esta ventana dura mientras corra la exposición (hasta mayo 2027)."
    ),
]

for r_idx, row in enumerate(patterns, start=1):
    for c_idx, val in enumerate(row):
        ws2.write(r_idx, c_idx, val, cell)
    ws2.set_row(r_idx, 90)

# ── TAB 3: Evolución y tendencias (4 semanas) ────────────────────────────────
ws3 = workbook.add_worksheet("Evolución 4 semanas")

ws3.set_column(0, 0, 40)
ws3.set_column(1, 1, 75)
ws3.set_column(2, 2, 30)

ws3.write(0, 0, "Tendencia", hdr)
ws3.write(0, 1, "Evolución (semanas 25-may → 15-jun)", hdr)
ws3.write(0, 2, "Dirección", hdr)
ws3.set_row(0, 30)

tendencias = [
    (
        "Olga de Amaral — posicionamiento multi-fair",
        "Sem 1 (25-may): Lisson la lleva a Art Basel HK. "
        "Sem 2 (1-jun): Momentum institucional en NY destacado. "
        "Sem 3 (8-jun): 3 ediciones Art Basel + Akris. "
        "Sem 4 (15-jun): Art Basel Basel abre MAÑANA con de Amaral en el booth principal. "
        "Pico de 4 semanas de visibilidad acumulada.",
        "↑↑ ACELERANDO"
    ),
    (
        "Julio Le Parc — de fallecimiento a retrospectiva activa",
        "Sem 1 (25-may): Sin señal. "
        "Sem 2 (1-jun): Sin señal. "
        "Sem 3 (8-jun): Fallecimiento 30 mayo reportado; Tate retro anunciada para 11 jun. "
        "Sem 4 (15-jun): Retrospectiva ABIERTA desde el 11 jun. Mercado secundario activo.",
        "↑ CONSOLIDANDO"
    ),
    (
        "Casas Riegner — única galería colombiana circuito A",
        "Sem 1 (25-may): Confirmada en Art Basel Basel 2025 y 2026; oficina Madrid. "
        "Sem 2 (1-jun): Art Basel Basel 2026 a 15 días, activaciones recomendadas. "
        "Sem 3 (8-jun): A una semana de la feria. "
        "Sem 4 (15-jun): VIP MAÑANA. Pico anual. Brecha consolidad 4ª semana consecutiva.",
        "↑ PICO ANUAL"
    ),
    (
        "Ascaso Gallery — amenaza de roster sostenida",
        "Sem 1 (25-may): Formalizan relación con Fundación Julio Larraz (riesgo exclusividad). "
        "Sem 2 (1-jun): ALERTA — marketing activo de Botero y Cruz-Diez en EE.UU. "
        "Sem 3 (8-jun): Continúa sin respuesta visible de Duque Arango. "
        "Sem 4 (15-jun): Nuevo show anunciado (Hem, 30 jun). Amenaza permanente activa.",
        "→ SOSTENIDA (sin respuesta)"
    ),
    (
        "Arte Cinético/Op Art latinoamericano — narrativa en auge",
        "Sem 1-2 (25-may a 1-jun): Sin señal específica. "
        "Sem 3 (8-jun): Le Parc (artista cinético) en foco por fallecimiento. "
        "Sem 4 (15-jun): Convergencia: Le Parc en Tate + 'Optical Experience' en Latin Art Core "
        "+ Cruz-Diez en Ascaso. Tres señales en la misma semana.",
        "↑ NUEVO MOMENTUM"
    ),
    (
        "Galería Casa Cuadrada y El Museo — silencio prolongado",
        "Sem 1 (25-may): Sin actividad nueva. "
        "Sem 2 (1-jun): Sin actividad nueva. "
        "Sem 3 (8-jun): Sin actividad nueva. "
        "Sem 4 (15-jun): Sin actividad nueva. 4ª semana consecutiva sin señal.",
        "↓ DORMIDA"
    ),
]

for r_idx, row in enumerate(tendencias, start=1):
    ws3.write(r_idx, 0, row[0], cell)
    ws3.write(r_idx, 1, row[1], cell)
    direction = row[2]
    if "↑↑" in direction or "PICO" in direction:
        ws3.write(r_idx, 2, direction, cell_alert)
    elif "↑" in direction:
        ws3.write(r_idx, 2, direction, cell_nuevo)
    elif "SOSTENIDA" in direction:
        ws3.write(r_idx, 2, direction, cell_continua)
    else:
        ws3.write(r_idx, 2, direction, cell_cambio)
    ws3.set_row(r_idx, 80)

workbook.close()

size = os.path.getsize(OUTPUT)
print(f"File created: {OUTPUT}")
print(f"File size: {size} bytes")
if size > 0:
    print("VALIDATION: OK")
else:
    print("VALIDATION: FAILED — file is empty")
