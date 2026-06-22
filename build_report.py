#!/usr/bin/env python3
"""Build Gallery_Intelligence_2026_06_22.xlsx"""
import xlsxwriter
import os

OUT = "/home/user/da-routines/Gallery_Intelligence_2026_06_22.xlsx"

wb = xlsxwriter.Workbook(OUT)

# ── Formats ──────────────────────────────────────────────────────────────────
hdr = wb.add_format({
    "bold": True, "bg_color": "#1A1A2E", "font_color": "#FFFFFF",
    "border": 1, "text_wrap": True, "valign": "vcenter", "align": "center",
    "font_size": 11
})
cell = wb.add_format({
    "border": 1, "text_wrap": True, "valign": "top", "font_size": 10
})
nuevo = wb.add_format({
    "border": 1, "text_wrap": True, "valign": "top", "font_size": 10,
    "bg_color": "#D5F5E3"   # green
})
continua = wb.add_format({
    "border": 1, "text_wrap": True, "valign": "top", "font_size": 10,
    "bg_color": "#EBF5FB"   # light blue
})
cambio = wb.add_format({
    "border": 1, "text_wrap": True, "valign": "top", "font_size": 10,
    "bg_color": "#FEF9E7"   # yellow
})
desap = wb.add_format({
    "border": 1, "text_wrap": True, "valign": "top", "font_size": 10,
    "bg_color": "#FDEDEC"   # light red
})
no_act = wb.add_format({
    "border": 1, "text_wrap": True, "valign": "top", "font_size": 10,
    "font_color": "#999999", "italic": True
})
bold_cell = wb.add_format({
    "border": 1, "text_wrap": True, "valign": "top", "font_size": 10, "bold": True
})
pattern_hdr = wb.add_format({
    "bold": True, "bg_color": "#2C3E50", "font_color": "#FFFFFF",
    "border": 1, "text_wrap": True, "valign": "vcenter", "align": "center",
    "font_size": 11
})
trend_hdr = wb.add_format({
    "bold": True, "bg_color": "#154360", "font_color": "#FFFFFF",
    "border": 1, "text_wrap": True, "valign": "vcenter", "align": "center",
    "font_size": 11
})
trend_cell = wb.add_format({
    "border": 1, "text_wrap": True, "valign": "top", "font_size": 10
})

STATUS_FMT = {
    "NUEVO": nuevo,
    "CONTINÚA": continua,
    "CAMBIÓ": cambio,
    "DESAPARECIÓ": desap,
    "N/A": no_act,
}

# ═══════════════════════════════════════════════════════════════════════════
# TAB 1 — Changes & Findings
# ═══════════════════════════════════════════════════════════════════════════
ws1 = wb.add_worksheet("Hallazgos y Cambios")
ws1.set_zoom(85)

cols = [
    "Galería", "Tier", "Sección", "Tipo de Cambio",
    "Descripción del Hallazgo",
    "Nota Estratégica para Duque Arango",
    "Estado vs. 4 semanas", "URL"
]
col_w = [22, 6, 16, 20, 55, 42, 14, 32]
for i, (c, w) in enumerate(zip(cols, col_w)):
    ws1.set_column(i, i, w)
    ws1.write(0, i, c, hdr)
ws1.set_row(0, 36)
ws1.freeze_panes(1, 0)

rows = [
    # Gallery | Tier | Section | Type | Description | Strategic note | Status | URL
    (
        "Casas Riegner", "1",
        "Feria / Art Basel Basel",
        "Fair Announcement / Presencia activa",
        "Art Basel Basel 2026 (Jun 18–21) cerró ayer con 90.000 visitantes y ventas fuertes. Casas Riegner participó en el Booth L17 — única galería colombiana en el sector principal Galleries. El mercado $200k–$2M fue el más activo del año. Post-feria: recapitulaciones editoriales y búsquedas de coleccionistas crecerán esta semana.",
        "La ventana post-Basel es ahora mismo: coleccionistas recién llegados de Basilea buscan arte colombiano online. Publicar hoy: reseña del mercado latinoamericano en Art Basel, artistas del roster con mención editorial. Casas Riegner amplía cada semana su brecha de visibilidad internacional.",
        "CONTINÚA",
        "https://www.artbasel.com/catalog/gallery/1082/Casas-Riegner"
    ),
    (
        "Galería La Cometa", "1",
        "Exposiciones en curso",
        "New Content / Programa activo",
        "Exposiciones vigentes hasta 5 jul: 'Envoltorios' (Asicaz Monzón) y 'Nada Es Lo Que Parece, Al Parecer Desaparece II' (Luisa Aristizábal) en Medellín; 'Parar el Mundo' (Adam Goldstein) y 'Pensamiento mágico / El año entrante' (Alejandro Ospina) en Madrid. Todas iniciaron sem. del 9–14 may. No se detectó actividad nueva esta semana.",
        "La Cometa mantiene programación activa en 4 ciudades (Bogotá, Medellín, Madrid, Miami). Duque Arango debería tener una presencia editorial similar durante la semana de Art Basel para no perder visibilidad comparativa.",
        "CONTINÚA",
        "https://galerialacometa.com/exhibiciones/"
    ),
    (
        "Galería Casa Cuadrada", "1",
        "Sitio / Actividad",
        "Operational",
        "No se detectó actividad nueva esta semana. Quinta semana consecutiva sin contenido nuevo publicado o exposición anunciada.",
        "Cinco semanas de silencio digital durante el período de mayor tráfico del año (Art Basel). Oportunidad para Duque Arango de ganar posicionamiento en búsquedas de arte colombiano contemporáneo sin competencia directa de este actor.",
        "CONTINÚA",
        "http://galeriacasacuadrada.com"
    ),
    (
        "Galería El Museo", "1",
        "Exposiciones",
        "Operational",
        "Las exposiciones previas ('Cartografías del cuerpo' de Aurora Lario, 'Negro humo' de Fredy Alzate, 'Tintín y Milú en Macondo' de Gabriel Ortega) cerraron el 6 de junio. No se detectaron nuevas exposiciones anunciadas para después del 15 de junio.",
        "Galería El Museo continúa en pausa de programación pública. Si Duque Arango publica contenido sobre arte colombiano esta semana, no encontrará competencia de dos de sus cuatro competidores Tier 1 directos.",
        "DESAPARECIÓ",
        "https://www.galeriaelmuseo.com"
    ),
    (
        "Latin Art Core", "2",
        "Exposición",
        "New Exhibition (CONTINÚA)",
        "'The Optical Experience': muestra sobre Arte Concreto cubano y legado Op Art latinoamericano — sigue en cartel desde el 12 de junio. Contexto de visibilidad sostenida: Tate Modern tiene activa la retrospectiva de Julio Le Parc (roster Duque Arango) hasta mayo 2027.",
        "El movimiento cinético/Op Art latinoamericano recibe atención sostenida de múltiples frentes (Latin Art Core + Tate/Le Parc + Ascaso/Cruz-Diez). Duque Arango tiene a Cruz-Diez y Le Parc en su roster: es el momento de publicar un artículo de autoridad sobre este movimiento.",
        "CONTINÚA",
        "https://latinartcore.com"
    ),
    (
        "Art of the World Gallery", "2",
        "Exposición",
        "Operational",
        "La exposición anterior ('Mimetism: Echoes that Breathe' de Karla de Lara) cerró el 6 de junio. No se detectaron nuevas exposiciones o actividad después del 15 de junio.",
        "Sin actividad detectable esta semana.",
        "N/A",
        "https://www.artoftheworldgallery.com"
    ),
    (
        "Ascaso Gallery", "2",
        "Nueva exposición + Roster",
        "New Exhibition",
        "NUEVO: Anuncia 'Reviver', primera exposición individual de Andrew Hem (Los Ángeles), con apertura el 30 de junio. CONTINÚA: la galería sigue representando activamente a Fernando Botero, Carlos Cruz-Diez y Julio Larraz en el mercado estadounidense. Cuarta semana consecutiva sin respuesta SEO visible de Duque Arango.",
        "🚨 Amenaza de roster en semana 4. Ascaso no solo representa a artistas compartidos — también profundiza su programación con nuevas exposiciones individuales, aumentando su autoridad de galería de destino en EE.UU. Crear/actualizar páginas de artista y contenido de autoridad para Botero, Cruz-Diez y Larraz en galeriaduquearango.com es ya urgente.",
        "NUEVO",
        "https://www.ascasogallery.com"
    ),
    (
        "Galería Freites", "2",
        "Actividad semanal",
        "Operational",
        "No se detectó nueva actividad esta semana. La galería (Caracas, Miami, Madrid) representa a Fernando Botero y Manolo Valdés, entre otros artistas del roster de Duque Arango.",
        "Galería Freites representa a Botero y Manolo Valdés (ambos en el roster de Duque Arango) sin actividad visible esta semana. Sin embargo, la presencia establecida de Freites en mercados de Miami y Madrid con estos artistas es una competencia estructural latente.",
        "CONTINÚA",
        "https://galeriafreites.com"
    ),
    (
        "Opera Gallery", "2",
        "Nueva exposición",
        "New Exhibition",
        "NUEVO: Anuncia exposición individual de Gustavo Nazareno en Madrid, 25 jun – 29 ago. Adicionalmente: nueva exposición en París (26 jun – 15 jul), Dubai (1 jul – 31 ago) y Mónaco (3 jul – 31 ago). Xevi Solà en Seúl desde 10 julio.",
        "Opera Gallery activa su red global de espacios simultáneamente post-Art Basel. No hay artistas del roster de Duque Arango en las exposiciones anunciadas esta semana, pero el modelo de activación global post-feria es una práctica a replicar.",
        "NUEVO",
        "https://www.operagallery.com/exhibitions"
    ),
    (
        "Gagosian", "3",
        "Art Basel Basel",
        "Fair Announcement",
        "Presencia destacada en Art Basel Basel (Jun 16–21) con obras de postguerra y contemporáneas. No se detectaron artistas latinoamericanos específicos del roster de Duque Arango en el booth. Fuera de foco Tier 3 esta semana.",
        "Sin actividad relevante para el roster de Duque Arango esta semana.",
        "N/A",
        "https://gagosian.com/fairs-and-collecting/fairs/art-basel-2026/"
    ),
    (
        "David Zwirner", "3",
        "Exposición institucional",
        "Institutional",
        "Óscar Murillo: 'Collective Osmosis' en Das Minsk (Potsdam, Alemania) continúa hasta el 9 de agosto de 2026. Murillo participa en conversación con curador Bonaventure Soh Bejeng Ndikung. Murillo es artista del roster de Duque Arango.",
        "Óscar Murillo tiene visibilidad institucional sostenida en Europa a través de David Zwirner. Duque Arango debe amplificar con contenido propio: artículo editorial sobre Murillo, link a la exposición, posicionamiento como la galería colombiana que lo conoce mejor.",
        "CONTINÚA",
        "https://www.davidzwirner.com/artists/oscar-murillo"
    ),
    (
        "Hauser & Wirth", "3",
        "Exposiciones en curso",
        "New Content",
        "Dos exposiciones con artistas latinoamericanos/caribeños activas: (1) Firelei Báez 'feet squelching on wet grass, nourished by uncertainty' hasta 31 jul 2026; (2) Angel Otero 'Agua Salada' (puertorriqueño) hasta 17 oct 2026. En Art Basel Basel: ventas totales superaron $65M (Picasso $35M, Richter $20M, Bourgeois $2.5M).",
        "Hauser & Wirth consolida su liderazgo en arte latinoamericano/caribeño con programación sostenida. Otero y Báez no están en el roster de Duque Arango, pero señalan que la demanda por arte latinoamericano en espacios top sigue activa post-Art Basel.",
        "CONTINÚA",
        "https://www.hauserwirth.com/hauser-wirth-exhibitions/"
    ),
    (
        "Galerie Lelong", "3",
        "Actividad semanal",
        "Operational",
        "No se detectó actividad nueva con artistas latinoamericanos esta semana. La galería participó en ARCOmadrid 2026 (marzo). Programa latinoamericano histórico incluye Alfredo Jaar, Cildo Meireles, Ana Mendieta, Zilia Sánchez.",
        "Sin actividad relevante para el roster de Duque Arango esta semana.",
        "N/A",
        "https://www.galerie-lelong.com/en/"
    ),
    (
        "Lisson Gallery", "3",
        "Post-Art Basel / Roster",
        "Institutional",
        "Concluye ciclo de 4 semanas con Olga de Amaral (roster Duque Arango): Rojo y oro (2016) en booth principal de Art Basel Basel cerrado ayer. Post-feria: Shirazeh Houshiary en National Gallery London (19 jun); Lisson en Royal Academy Summer Exhibition 2026; Ha Chong-Hyun exhibición anunciada para otoño en Londres. Olga de Amaral sigue como figura central del programa.",
        "El ciclo Lisson/Olga de Amaral (Art Basel HK → Qatar → Basel) cerró ayer. La demanda de búsqueda por Olga de Amaral permanecerá elevada esta semana. Publicar contenido editorial de autoridad sobre de Amaral HOY captura tráfico orgánico de alta intención sin competencia directa del momento.",
        "CAMBIÓ",
        "https://www.lissongallery.com/news/lisson-at-art-basel-2026"
    ),
    (
        "Lehmann Maupin", "3",
        "Actividad semanal",
        "Operational",
        "30° aniversario de la galería en 2026. Cecilia Vicuña (chilena) representada. No se detectó nueva exposición o contenido latinoamericano específico esta semana.",
        "Sin actividad relevante para el roster de Duque Arango esta semana.",
        "N/A",
        "https://www.lehmannmaupin.com"
    ),
    (
        "Perrotin", "3",
        "Exposición en curso",
        "New Exhibition (CONTINÚA)",
        "Hugo Toro 'Ojo de Agua' sigue en Perrotin New York (10 jun – 31 jul). Toro es artista franco-mexicano, debut en EE.UU. La muestra explora identidad dual, memoria y agua. No hay artistas del roster de Duque Arango.",
        "Perrotin invierte en artistas latinoamericanos/híbridos con identidad compleja en Nueva York. Señal de demanda sostenida por narrativas de identidad latinoamericana en el mercado norteamericano.",
        "CONTINÚA",
        "https://www.perrotin.com/artists/hugo-toro"
    ),
]

for r_idx, row_data in enumerate(rows, start=1):
    status = row_data[6]
    fmt = STATUS_FMT.get(status, cell)
    for c_idx, val in enumerate(row_data):
        ws1.write(r_idx, c_idx, val, fmt)
    ws1.set_row(r_idx, 72)

# Legend
ws1.write(len(rows) + 2, 0, "Leyenda de colores:", bold_cell)
for i, (label, fmt_key) in enumerate([
    ("NUEVO — No reportado en últimas 4 semanas", "NUEVO"),
    ("CONTINÚA — Activo y reportado previamente", "CONTINÚA"),
    ("CAMBIÓ — Evolucionó respecto a semanas anteriores", "CAMBIÓ"),
    ("DESAPARECIÓ — Estaba activo, ahora sin actividad", "DESAPARECIÓ"),
]):
    ws1.write(len(rows) + 3 + i, 0, label, STATUS_FMT[fmt_key])

# ═══════════════════════════════════════════════════════════════════════════
# TAB 2 — Cross-Gallery Patterns
# ═══════════════════════════════════════════════════════════════════════════
ws2 = wb.add_worksheet("Patrones Cruzados")
ws2.set_zoom(85)

p_cols = [
    "Patrón Observado",
    "Galerías involucradas",
    "Qué puede señalar",
    "Acción recomendada para Duque Arango"
]
p_widths = [35, 28, 38, 42]
for i, (c, w) in enumerate(zip(p_cols, p_widths)):
    ws2.set_column(i, i, w)
    ws2.write(0, i, c, pattern_hdr)
ws2.set_row(0, 36)
ws2.freeze_panes(1, 0)

patterns = [
    (
        "Vacío digital colombiano durante Art Basel: cuatro galerías Tier 1 silenciosas durante la semana más importante del año",
        "Galería La Cometa (sin nueva actividad), Casa Cuadrada (sem. 5 de silencio), Galería El Museo (sem. 3 sin exposición), Duque Arango",
        "El mercado digital del arte colombiano está desatendido exactamente cuando el mundo busca 'arte colombiano'. Casas Riegner es la única que capitaliza esto. La brecha de visibilidad digital se ensancha cada año de Art Basel.",
        "Publicar esta semana: artículo editorial 'Lo que ver en el mercado del arte colombiano post-Art Basel 2026'. Incluir artistas del roster, piezas destacadas y por qué coleccionistas deben mirar a galeriaduquearango.com. La ventana es esta semana — después, el tráfico baja."
    ),
    (
        "Arte Cinético/Op Art latinoamericano: narrativa sostenida en múltiples galerías simultáneamente (2ª semana)",
        "Latin Art Core ('The Optical Experience'), Ascaso Gallery (Cruz-Diez), Lisson Gallery (Le Parc/Oiticica/Tunga en Art Basel Basel)",
        "El movimiento cinético y concretista latinoamericano recibe atención institucional y comercial sincronizada. No es coincidencia: es una narrativa de mercado activa que múltiples actores están amplificando. La retrospectiva de Le Parc en Tate (hasta may 2027) es el catalizador.",
        "Duque Arango tiene a Carlos Cruz-Diez y Julio Le Parc en su roster — dos artistas centrales de este movimiento. Publicar: 'El legado del Arte Cinético Latinoamericano: Cruz-Diez y Le Parc'. Optimizar estas páginas de artista para búsquedas relacionadas con la retrospectiva de Tate."
    ),
    (
        "Amenaza de roster por Ascaso Gallery: 4ª semana consecutiva sin respuesta",
        "Ascaso Gallery (Fernando Botero, Carlos Cruz-Diez, Julio Larraz en Miami); Galería Freites (Botero, Manolo Valdés en Miami/Madrid/Caracas)",
        "Dos galerías competidoras representan activamente a 4 artistas del roster de Duque Arango en mercados de EE.UU. Sin contenido de autoridad de Duque Arango, los coleccionistas que buscan estas obras en Google llegan a los competidores primero. Cada semana de inacción cede terreno de búsqueda permanente.",
        "🚨 Prioridad 1: Crear/actualizar páginas de artista para Botero, Cruz-Diez, Larraz y Manolo Valdés en galeriaduquearango.com con contenido de autoridad (biografía extendida, obras representativas, trayectoria, contexto del mercado). Target: posicionar en búsquedas '[artista] galería Colombia', '[artista] obra en venta'."
    ),
    (
        "Activación post-Art Basel: galerías internacionales anuncian programación simultáneamente",
        "Opera Gallery (Madrid 25 jun, París 26 jun, Dubai 1 jul, Mónaco 3 jul), Lisson Gallery (Ra Summer Exhib., Houshiary talk), Hauser & Wirth (programación en curso)",
        "Las galerías de referencia usan el post-Art Basel como plataforma de lanzamiento de su siguiente ciclo de programación. Anuncian shows simultáneamente para mantener el momentum mediático. Es una práctica de comunicación sistemática, no reactiva.",
        "Duque Arango debería tener un calendario editorial post-Basel: anunciar la próxima exposición, publicar un artículo de opinión sobre el mercado del arte latinoamericano, o activar en redes sociales con contenido de artistas del roster. El silencio post-feria desperdicia el pico de atención del año."
    ),
    (
        "Olga de Amaral: cierre del ciclo Lisson 2026 — demanda de búsqueda en pico esta semana",
        "Lisson Gallery (Art Basel HK, Qatar, Basel — ciclo de 4 semanas concluido), galeriaduquearango.com (sin contenido editorial sobre de Amaral)",
        "El ciclo de presencia institucional de Olga de Amaral en tres ediciones de Art Basel + Akris generó meses de cobertura mediática. Los coleccionistas que vieron su obra en Basel esta semana buscarán más información online en los próximos 7–10 días. La demanda de búsqueda está en su pico anual.",
        "Publicar esta semana: artículo de autoridad sobre Olga de Amaral — su obra, su legado, su relación con el textil latinoamericano, y el contexto de su momento de mercado. Objetivo: aparecer en búsquedas de coleccionistas post-Basel que buscan 'Olga de Amaral obra', 'Olga de Amaral galería Colombia'."
    ),
]

for r_idx, row_data in enumerate(patterns, start=1):
    for c_idx, val in enumerate(row_data):
        ws2.write(r_idx, c_idx, val, cell)
    ws2.set_row(r_idx, 90)

# ═══════════════════════════════════════════════════════════════════════════
# TAB 3 — Evolución y Tendencias (4 semanas)
# ═══════════════════════════════════════════════════════════════════════════
ws3 = wb.add_worksheet("Evolución 4 Semanas")
ws3.set_zoom(85)

t_cols = [
    "Tendencia / Actor",
    "25 may",
    "1 jun",
    "8 jun",
    "15 jun",
    "22 jun (esta semana)",
    "Dirección",
    "Implicación para Duque Arango"
]
t_widths = [26, 20, 20, 20, 20, 28, 12, 38]
for i, (c, w) in enumerate(zip(t_cols, t_widths)):
    ws3.set_column(i, i, w)
    ws3.write(0, i, c, trend_hdr)
ws3.set_row(0, 36)
ws3.freeze_panes(1, 0)

trends = [
    (
        "Casas Riegner — presencia internacional",
        "Confirmada en Art Basel Basel 2026",
        "Única galería Col. en Art Basel",
        "VIP preview inminente",
        "VIP 16–17 jun, feria abierta 18–21 jun",
        "Feria cerró ayer (21 jun). Booth L17. Post-feria: recaps y tráfico de búsqueda",
        "PICO → descenso",
        "Publicar editorial post-Basel HOY para capturar tráfico que Casas Riegner atraerá a la narrativa colombiana."
    ),
    (
        "Ascaso Gallery — amenaza de roster (Botero, Cruz-Diez, Larraz)",
        "Semana 1: alerta inicial detectada",
        "Semana 2: confirmado activo",
        "Semana 3: continúa sin respuesta Duque Arango",
        "Semana 4: nuevo show 'Reviver' anunciado (30 jun)",
        "Semana 5: 'Reviver' abre en 8 días. Profundidad de programación aumenta.",
        "↗ ESCALANDO",
        "🚨 4 semanas sin respuesta = 4 semanas de terreno cedido en búsquedas de EE.UU. Acción urgente en SEO de páginas de artista."
    ),
    (
        "Olga de Amaral — momentum institucional (Lisson)",
        "Presencia en Art Basel HK anunciada",
        "Art Basel Qatar + Akris activos",
        "Ciclo completo: HK + Qatar + Basel confirmado. Akris cobertura cruzada.",
        "PICO: booth principal Art Basel Basel con Rojo y oro (2016)",
        "Ciclo Lisson concluido ayer. Búsquedas post-Basel en pico esta semana.",
        "PICO → plateau",
        "Publicar artículo editorial sobre de Amaral esta semana: demanda de búsqueda es máxima ahora. Capturar tráfico orgánico de coleccionistas post-Basel."
    ),
    (
        "Julio Le Parc — retrospectiva Tate Modern",
        "Fallecimiento (30 may), retrospectiva preparada",
        "Retrospectiva póstumo abierta (11 jun)",
        "1ª semana completa en Tate",
        "2ª semana activa, nueva cobertura en medios",
        "3ª semana. Tate corre hasta may 2027. Interés sostenido.",
        "→ SOSTENIDO",
        "La ventana editorial es larga (hasta 2027). Publicar homenaje y contenido de autoridad sobre Le Parc (roster Duque Arango). Optimizar para búsqueda 'Julio Le Parc galería'."
    ),
    (
        "Arte Cinético/Op Art latinoamericano — narrativa de mercado",
        "Sin señal específica",
        "Ninguna señal",
        "Tate/Le Parc + Ascaso/Cruz-Diez: 2 señales",
        "Latin Art Core 'The Optical Experience' + Tate + Ascaso: 3 señales simultáneas",
        "Narrativa consolidada: 3 frentes activos en semana 2. Perrotin/Lisson también con artistas de la región.",
        "↗ CRECIENDO",
        "Momento editorial ideal: 'El renacimiento del Arte Cinético Latinoamericano'. Duque Arango tiene Cruz-Diez y Le Parc. Oportunidad de posicionarse como voz de autoridad en esta narrativa."
    ),
    (
        "Casa Cuadrada + El Museo — silencio digital",
        "Sin actividad semana 1",
        "Sin actividad semana 2",
        "Sin actividad semana 3",
        "Sin actividad semana 4",
        "Sem. 5 de silencio. El Museo sin exposición nueva desde 6 jun.",
        "→ PERSISTENTE",
        "Vacío de contenido competitivo en Tier 1 Colombia sigue. Duque Arango puede ganar visibilidad digital en este espacio con bajo esfuerzo relativo."
    ),
]

for r_idx, row_data in enumerate(trends, start=1):
    for c_idx, val in enumerate(row_data):
        ws3.write(r_idx, c_idx, val, trend_cell)
    ws3.set_row(r_idx, 72)

wb.close()

# Validate
size = os.path.getsize(OUT)
print(f"File: {OUT}")
print(f"Size: {size} bytes")
print("VALID" if size > 0 else "INVALID — empty file")
