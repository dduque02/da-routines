#!/usr/bin/env python3
"""Build Gallery Intelligence Excel report for week of 2026-06-29."""

import xlsxwriter
import os

DATE = "2026-06-29"
FILENAME = f"Gallery_Intelligence_{DATE.replace('-', '_')}.xlsx"
OUTPUT_PATH = os.path.join("/home/user/da-routines", FILENAME)

wb = xlsxwriter.Workbook(OUTPUT_PATH)

# ── Formats ────────────────────────────────────────────────────────────────
hdr = wb.add_format({
    "bold": True, "font_color": "#FFFFFF", "bg_color": "#1a1a2e",
    "border": 1, "text_wrap": True, "valign": "top", "align": "center"
})
nuevo = wb.add_format({"bg_color": "#d4edda", "border": 1, "text_wrap": True, "valign": "top"})
continua = wb.add_format({"bg_color": "#cce5ff", "border": 1, "text_wrap": True, "valign": "top"})
cambio = wb.add_format({"bg_color": "#fff3cd", "border": 1, "text_wrap": True, "valign": "top"})
desaparece = wb.add_format({"bg_color": "#f8d7da", "border": 1, "text_wrap": True, "valign": "top"})
noact = wb.add_format({"bg_color": "#f8f9fa", "font_color": "#6c757d", "border": 1, "text_wrap": True, "valign": "top"})
wrap = wb.add_format({"border": 1, "text_wrap": True, "valign": "top"})
bold_wrap = wb.add_format({"bold": True, "border": 1, "text_wrap": True, "valign": "top"})
pattern_hdr = wb.add_format({
    "bold": True, "font_color": "#FFFFFF", "bg_color": "#2c3e50",
    "border": 1, "text_wrap": True, "valign": "top", "align": "center"
})
trend_hdr = wb.add_format({
    "bold": True, "font_color": "#FFFFFF", "bg_color": "#8e44ad",
    "border": 1, "text_wrap": True, "valign": "top", "align": "center"
})
green_bold = wb.add_format({"bold": True, "font_color": "#27ae60", "border": 1, "text_wrap": True, "valign": "top"})
blue_bold = wb.add_format({"bold": True, "font_color": "#2980b9", "border": 1, "text_wrap": True, "valign": "top"})
yellow_bold = wb.add_format({"bold": True, "font_color": "#e67e22", "border": 1, "text_wrap": True, "valign": "top"})
red_bold = wb.add_format({"bold": True, "font_color": "#e74c3c", "border": 1, "text_wrap": True, "valign": "top"})

status_fmt = {
    "NUEVO": nuevo,
    "CONTINÚA": continua,
    "CAMBIÓ": cambio,
    "DESAPARECIÓ": desaparece,
    "Sin actividad": noact,
}

# ─────────────────────────────────────────────────────────────────────────────
# TAB 1 — Changes & Findings
# ─────────────────────────────────────────────────────────────────────────────
ws1 = wb.add_worksheet("Hallazgos y Cambios")
ws1.set_zoom(90)
ws1.freeze_panes(1, 0)

cols = ["Galería", "Tier", "Estado", "Sección", "Tipo de Cambio",
        "Descripción", "Nota Estratégica para Duque Arango", "URL"]
widths = [22, 6, 12, 14, 20, 50, 45, 35]
for i, (c, w) in enumerate(zip(cols, widths)):
    ws1.set_column(i, i, w)
    ws1.write(0, i, c, hdr)

rows = [
    # Gallery | Tier | Status | Section | Type | Description | Strategic Note | URL
    (
        "Lisson Gallery", "T3", "CAMBIÓ",
        "Programación / Exhibiciones",
        "Transición post-feria",
        "Tras 5 semanas de campaña intensa sobre Olga de Amaral (Art Basel HK → Qatar → Basel), Lisson pivota a nueva programación: Spencer Finch abre 26 jun en LA, Daniel Buren en Londres (11 jun–5 sep), Ha Chong-Hyun en otoño. La campaña de Olga de Amaral concluye con Basel.",
        "ALTA — Vacío de contenido post-Lisson sobre Olga de Amaral. DA representa a la artista y puede capturar el tráfico residual de búsqueda publicando editorial propio ESTA SEMANA: entrevista, obras disponibles, nota de proceso.",
        "https://www.lissongallery.com/exhibitions"
    ),
    (
        "Tate Modern (referencia)", "T3", "NUEVO",
        "Evento institucional",
        "Nueva exposición latinoamericana",
        "'Frida: The Making of an Icon' abre 25 jun 2026 en Tate Modern, Londres (hasta 3 ene 2027). Vendió 41,000 tickets anticipados — récord institucional de Tate. Simultáneo con Julio Le Parc (desde 11 jun, hasta may 2027). Dos exposiciones latinoamericanas en Tate al mismo tiempo.",
        "ALTA — DA representa a Julio Le Parc y Leonora Carrington. El verano 2026 es el peak de atención mediática e institucional sobre arte latinoamericano en Europa. Publicar contenido de autoridad sobre Le Parc y Carrington AHORA, mientras el interés está en su cima.",
        "https://www.tate.org.uk/whats-on/tate-modern/frida-kahlo-the-making-of-an-icon"
    ),
    (
        "Perrotin", "T3", "CONTINÚA",
        "Exhibiciones activas",
        "Programación latinoamericana",
        "Hugo Toro 'Ojo de Agua' en Perrotin New York (10 jun–31 jul 2026) — primera exposición del artista franco-mexicano en EEUU. Gabriel Rico 'Gabrielinos (I Am You And What I See Is Me)' en Perrotin LA (6 jun–11 jul 2026) — primera individual en LA. Dos muestras de artistas mexicanos activas simultáneamente en NY y LA.",
        "MEDIA — Perrotin apuesta fuerte por artistas mexicanos en EEUU. Señal de demanda de coleccionistas. DA tiene a Rufino Tamayo en su roster — oportunidad de posicionar contenido editorial sobre maestros mexicanos mientras el mercado está activo.",
        "https://www.perrotin.com/exhibitions/current"
    ),
    (
        "Casas Riegner", "T1", "CONTINÚA",
        "Exhibición institucional",
        "Retrospectiva internacional",
        "Beatriz González (1932–2026, falleció 9 ene 2026) — retrospectiva póstumo en Astrup Fearnley Museet, Oslo (12 jun–11 oct 2026). Final stop del tour internacional (también en Pinacoteca SP y Barbican Londres). 150+ obras. Casas Riegner confirmado en ArtBo 2026 (24–27 sep, no agosto como se reportó la semana pasada).",
        "ALTA — Casas Riegner mantiene presencia internacional activa con la retrospectiva González durante todo el verano. CORRECCIÓN: ArtBo 2026 es 24–27 septiembre, no agosto. DA debe confirmar participación con ~13 semanas de anticipación.",
        "https://www.casasriegner.com"
    ),
    (
        "Galerie Lelong", "T3", "CAMBIÓ",
        "Exhibición / Transición",
        "Cierre de muestra latinoamericana",
        "Lucia Laguna 'Apenas meus cabelos são brancos...' (primera individual en EEUU de la artista brasileña) cierra el 27 jun 2026 — esta semana. También representa a Elda Cerrato (argentina, 1930–2019) y Cildo Meireles (brasileño). Galerie Lelong es una de las principales galerías NY para arte latinoamericano.",
        "MEDIA — El cierre de Lucia Laguna esta semana libera espacio editorial. DA puede publicar contenido sobre artistas latinoamericanos en NY aprovechando el ciclo de cobertura post-Lelong. Monitorear qué programa Galerie Lelong como próxima muestra latinoamericana.",
        "https://galerielelong.com/exhibitions/"
    ),
    (
        "Ascaso Gallery", "T2", "CONTINÚA",
        "Programación / Actividad",
        "5ª semana activa consecutiva",
        "5ª semana consecutiva con actividad sostenida en el segmento latinoamericano Miami. Anuncio de show 'Reviver' (Andrew Hem, apertura 30 jun). También registrada actividad previa relacionada con Julio Larraz. Yelp actualizado junio 2026. Competidor más activo en el mercado latinoamericano de Miami.",
        "ALTA — Ascaso sigue siendo el competidor más constante en Miami Latin American. Overlap con Julio Larraz (DA roster). 5 semanas sin respuesta digital de DA en este segmento = terreno cedido en búsquedas US de coleccionistas.",
        "https://www.ascasogallery.com"
    ),
    (
        "Galería La Cometa", "T1", "CONTINÚA",
        "Exhibiciones en curso",
        "Programación activa multi-sede",
        "Exposiciones vigentes hasta 5 jul: 'Envoltorios' (Asicaz Monzón, Medellín), 'Nada Es Lo Que Parece II' (Luisa Aristizábal, Medellín), 'Parar el Mundo' (Adam Goldstein, Madrid), 'Pensamiento mágico' (Alejandro Ospina, Madrid). 4 sedes activas: Bogotá, Medellín, Madrid, Miami. NOTA: cambio de director artístico reportado semana pasada NO PUDO SER VERIFICADO esta semana.",
        "MEDIA — La Cometa mantiene programación activa multi-sede. Nada radicalmente nuevo esta semana. El cambio de director anterior no confirmado — monitorear.",
        "https://galerialacometa.com/exhibiciones/"
    ),
    (
        "Latin Art Core", "T2", "NUEVO",
        "Nueva exposición",
        "Individual de maestro cubano",
        "Manuel Mendive 'With the new day the sun shines and leads us' — individual del maestro cubano, abrió 19 jun 2026, activa esta semana. Mendive es una figura central del arte cubano contemporáneo. Show enfocado en obras 2023–2025.",
        "BAJA-MEDIA — Mendive no está en el roster de DA. Sin embargo, Latin Art Core continúa presentando maestros latinoamericanos con rigor institucional. Monitorear si presentan algún artista que compita con el roster de DA.",
        "https://latinartcore.com/exhibition/"
    ),
    (
        "Musée du Luxembourg (referencia)", "T3", "CONTINÚA",
        "Evento institucional",
        "Retrospectiva Leonora Carrington",
        "Leonora Carrington: primera gran retrospectiva en Francia (18 feb–19 jul 2026, 126 obras). Cierra en 3 semanas. Freud Museum Londres 'The Symptomatic Surreal' (hasta 10 ago). El interés institucional europeo por Carrington está en su pico.",
        "ALTA — DA tiene obras de Leonora Carrington. La retrospectiva de Luxemburgo cierra en 3 semanas — ventana final para publicar contenido editorial de autoridad y capturar búsquedas mientras el interés está máximo. Después del 19 jul la atención mediática disminuirá.",
        "https://museeduluxembourg.fr/en/agenda/evenement/leonora-carrington"
    ),
    (
        "Mercado post-Basel (señal)", "T3", "NUEVO",
        "Contexto de mercado",
        "Señal de mercado",
        "Art Basel 2026 cerró 21 jun con ventas sólidas. Latin American art market: galerias brasileñas +20% en SP-Arte YoY. Coleccionistas europeos activamente moviéndose hacia arte latinoamericano como mercado con precios accesibles pre-burbuja. Art Basel/UBS Market Report 2026: mercado global $59.6B.",
        "ALTA — El contexto de mercado es el más favorable para arte latinoamericano en años recientes. DA debe comunicar disponibilidad de obras a coleccionistas activos ESTA SEMANA — el post-Basel es el momento pico de interés.",
        "https://www.artbasel.com/stories/art-basel-2026-closes-with-strong-sales-and-global-engagement"
    ),
    (
        "David Zwirner", "T3", "CONTINÚA",
        "Post-feria",
        "Seguimiento post-Basel",
        "David Zwirner activo con seguimiento post-Art Basel Basel 2026. 'Art Basel Plus' — programa de obras adicionales (16–21 jun). Ventas destacadas en la feria. No se detectó programación específica latinoamericana post-22 jun.",
        "BAJA — Sin actividad latinoamericana específica esta semana. Señal de mercado activo general.",
        "https://www.davidzwirner.com/fairs/2026/art-basel"
    ),
    (
        "Galería Casa Cuadrada", "T1", "CONTINÚA",
        "Silencio digital",
        "Inactividad",
        "5ª semana consecutiva sin actividad digital detectable. Sin nuevas exhibiciones, noticias, ni actualizaciones en web o redes. Posible cierre temporal, renovación, o pausa programática.",
        "INFO — El silencio de Casa Cuadrada continúa. DA puede capturar su audiencia en Bogotá con contenido activo. Investigar si hay alguna comunicación oficial sobre su estado.",
        "http://galeriacasacuadrada.com"
    ),
    (
        "Galería El Museo", "T1", "CONTINÚA",
        "Silencio digital",
        "Inactividad",
        "3ª semana sin actividad digital detectable. El Museo ha representado históricamente a Fernando Botero, Alejandro Obregón, y Nadín Ospina — artistas que también están en el roster de DA.",
        "MEDIA — Sin actividad, pero el overlap de artistas (Nadín Ospina especialmente) requiere monitoreo. Si El Museo reduce su presencia, DA puede posicionarse más agresivamente en búsquedas de esos artistas.",
        "https://www.galeriaelmuseo.com"
    ),
    (
        "Galería Freites", "T2", "CONTINÚA",
        "Actividad digital mínima",
        "Actividad puntual",
        "Galería Freites (Caracas) sin nueva actividad detectable post-22 jun. La semana pasada retomó brevemente actividad digital tras 2 semanas en silencio. Sin eventos o exposiciones nuevas confirmadas esta semana.",
        "BAJA — Sin actividad relevante esta semana.",
        "https://galeriafreites.com"
    ),
    (
        "Hauser & Wirth", "T3", "CONTINÚA",
        "Contexto institucional",
        "Momentum Leonora Carrington",
        "No se detectó nueva actividad específica de H&W post-22 jun relacionada con artistas latinoamericanos. El momentum institucional de Leonora Carrington es vía museos (Luxemburgo, Freud Museum), no H&W directamente. Ha estado activo editorialmente sobre Carrington en semanas previas.",
        "MEDIA — H&W sigue siendo punto de referencia para coleccionistas de Carrington. DA debe publicar su propio contenido editorial sobre Carrington antes del cierre de la retrospectiva de Luxemburgo (19 jul).",
        "https://www.hauserwirth.com"
    ),
    (
        "Opera Gallery", "T2", "CONTINÚA",
        "Programación general",
        "Actividad no-latinoamericana",
        "Opera Gallery con muestra activa jun 26–jul 15. No se detectó programación latinoamericana específica esta semana. Multi-sede global activa.",
        "BAJA — Sin actividad latinoamericana relevante detectada.",
        "https://www.operagallery.com/exhibitions"
    ),
    (
        "Lehmann Maupin", "T3", "CONTINÚA",
        "Programación futura",
        "Sin actividad actual latinoamericana",
        "Cecilia Vicuña (Chile) planeada para Art Basel Miami Beach diciembre 2026. Loriel Beltrán (Venezuela) en roster. Sin actividad latinoamericana activa post-22 jun.",
        "BAJA — Sin actividad esta semana. Monitorear Cecilia Vicuña para ABMB en diciembre.",
        "https://www.lehmannmaupin.com/exhibitions"
    ),
    (
        "Art of the World Gallery", "T2", "Sin actividad",
        "—",
        "—",
        "No activity detected this week. La exposición de Karla de Lara cerró el 6 jun. Sin nueva programación detectada.",
        "—",
        "https://www.artoftheworldgallery.com/exhibitions/"
    ),
    (
        "Gagosian", "T3", "Sin actividad",
        "—",
        "—",
        "No activity detected this week for Latin American artists (Tier 3 filter applied).",
        "—",
        "https://gagosian.com/exhibitions/"
    ),
]

for i, row in enumerate(rows, start=1):
    status = row[2]
    fmt = status_fmt.get(status, wrap)
    for j, val in enumerate(row):
        if j == 2:
            ws1.write(i, j, val, fmt)
        else:
            ws1.write(i, j, val, wrap)

ws1.autofilter(0, 0, len(rows), len(cols) - 1)

# ─────────────────────────────────────────────────────────────────────────────
# TAB 2 — Cross-Gallery Patterns
# ─────────────────────────────────────────────────────────────────────────────
ws2 = wb.add_worksheet("Patrones Cruzados")
ws2.set_zoom(90)
ws2.freeze_panes(1, 0)

p_cols = ["Patrón Observado", "Galerías Involucradas", "Lo Que Puede Señalar", "Acción Recomendada para DA"]
p_widths = [32, 35, 40, 45]
for i, (c, w) in enumerate(zip(p_cols, p_widths)):
    ws2.set_column(i, i, w)
    ws2.write(0, i, c, pattern_hdr)

patterns = [
    (
        "Verano latinoamericano en Europa — pico sin precedentes",
        "Tate Modern (Frida Kahlo + Julio Le Parc), Musée du Luxembourg (Carrington), Freud Museum (Carrington)",
        "Verano 2026 = mayor concentración de atención institucional y mediática europea sobre arte latinoamericano en décadas. Coleccionistas europeos están buscando activamente arte latinoamericano con precios accesibles.",
        "URGENTE: DA tiene a Julio Le Parc y Leonora Carrington en su roster. Publicar contenido editorial de calidad esta semana — nota sobre Le Parc (Tate retrospectiva), ensayo sobre Carrington (cierre Luxemburgo en 3 semanas). Contactar coleccionistas europeos activos."
    ),
    (
        "Vacío post-Lisson sobre Olga de Amaral",
        "Lisson Gallery (campaña 5 semanas concluida), DA (representa exclusivamente a de Amaral en Colombia)",
        "Lisson terminó su campaña de Art Basel. El interés en Olga de Amaral sigue alto pero sin contenido fresco de su galería principal. La demanda de búsqueda no desaparece inmediatamente tras el fin de la campaña.",
        "Publicar THIS WEEK contenido editorial sobre Olga de Amaral: obras disponibles, nueva pieza editorial, statement de la artista. DA puede capturar el tráfico orgánico residual de la campaña Lisson. Ventana de 2–3 semanas antes de que el interés decaiga."
    ),
    (
        "Ascaso Gallery — 5ª semana activa sin respuesta de DA",
        "Ascaso Gallery (Miami), DA (Julio Larraz en roster compartido)",
        "Ascaso ha estado activo 5 semanas consecutivas con programación y presencia digital en el mercado latinoamericano de Miami. Cada semana adicional consolida su posicionamiento ante coleccionistas de ese mercado.",
        "Crear o actualizar páginas de Julio Larraz en galeriaduquearango.com con contenido SEO de autoridad. Activar presencia en redes para Miami. Considerar outreach directo a coleccionistas de Miami que podrían estar viendo ambas galerías."
    ),
    (
        "Perrotin apuesta por artistas mexicanos en EEUU",
        "Perrotin (Hugo Toro NY + Gabriel Rico LA simultáneos)",
        "Perrotin — la galería más commercially attuned de las T3 — está haciendo una apuesta doble por artistas mexicanos en sus dos sedes principales de EEUU. Señal de que hay demanda de coleccionistas para este perfil.",
        "DA tiene a Rufino Tamayo (maestro mexicano moderno) en su roster. Oportunidad de publicar contenido sobre Tamayo como referente histórico del arte mexicano del que Perrotin es el 'sucesor contemporáneo'. Conectar narrativa pasado-presente."
    ),
    (
        "Casas Riegner + Instituto de Visión → ArtBo 2026 (24–27 sep)",
        "Casas Riegner (confirmado ArtBo), mercado colombiano activo",
        "ArtBo 2026 es 24–27 septiembre (CORRECCIÓN: no agosto como se reportó la semana pasada). Las galerías colombianas más activas ya están confirmadas. Hay 13 semanas para preparación.",
        "DA debe confirmar participación en ArtBo 2026 ESTA SEMANA si aún no lo ha hecho. Con 13 semanas de anticipación hay tiempo para preparar booth de calidad. Artistas a priorizar: Olga de Amaral, Óscar Murillo, Javier Caraballo, Alejandra Aristizábal."
    ),
    (
        "Silencio digital sostenido de competidores colombianos",
        "Casa Cuadrada (5 semanas), Galería El Museo (3 semanas), Galería Freites (actividad mínima)",
        "Tres competidores con presencia histórica relevante están en pausa digital. Esto crea un vacío de contenido en el mercado colombiano que DA puede llenar.",
        "Publicar contenido activo y consistente esta semana y las próximas: exposiciones actuales, artistas del roster, contexto de mercado. El silencio de Casa Cuadrada y El Museo es una oportunidad de SEO y redes para DA."
    ),
]

for i, row in enumerate(patterns, start=1):
    for j, val in enumerate(row):
        ws2.write(i, j, val, wrap)

# ─────────────────────────────────────────────────────────────────────────────
# TAB 3 — Evolución y Tendencias (4 semanas)
# ─────────────────────────────────────────────────────────────────────────────
ws3 = wb.add_worksheet("Evolución 4 Semanas")
ws3.set_zoom(90)
ws3.freeze_panes(1, 0)

t_cols = ["Tendencia", "01 jun", "08 jun", "15 jun", "22 jun", "29 jun (hoy)", "Dirección"]
t_widths = [30, 22, 22, 22, 22, 28, 15]
for i, (c, w) in enumerate(zip(t_cols, t_widths)):
    ws3.set_column(i, i, w)
    ws3.write(0, i, c, trend_hdr)

trends = [
    (
        "Amenaza Ascaso Gallery (Miami)",
        "Semana 1 — Debut nueva actividad",
        "Semana 2 — Activa",
        "Semana 3 — Sostenida con show 'Reviver' anunciado",
        "Semana 4 — Escalada con nueva muestra",
        "Semana 5 — CONTINÚA. Show 'Reviver' (Andrew Hem) abre 30 jun",
        "↑ Escalando"
    ),
    (
        "Campaña Lisson / Olga de Amaral",
        "Inicio campaña Art Basel HK",
        "Art Basel Qatar + Akris collab activa",
        "VIP Art Basel Basel — pico de campaña",
        "Post-Basel — pico final de visibilidad",
        "CAMBIÓ — Lisson pivota a nueva programación. Vacío se abre",
        "↓ Concluye → oportunidad DA"
    ),
    (
        "Arte latinoamericano en instituciones europeas",
        "—",
        "Leonora Carrington en Luxemburgo (feb–jul)",
        "Julio Le Parc abre en Tate Modern",
        "Semana 2 Tate/Le Parc + Carrington activa",
        "NUEVO: Frida Kahlo en Tate (25 jun, récord 41K tickets). Pico máximo del verano",
        "↑↑ Peak histórico"
    ),
    (
        "Casas Riegner presencia internacional",
        "4ª semana activa Art Basel Basel",
        "Preparación Basel activa",
        "VIP Art Basel Basel",
        "Post-Basel + Oslo retrospectiva González iniciada",
        "CONTINÚA — Oslo en semana 3. ArtBo sep confirmado",
        "→ Estable, sólida"
    ),
    (
        "Silencio Casa Cuadrada",
        "Activa",
        "Semana 1 silencio",
        "Semana 2 silencio",
        "Semana 4 silencio",
        "Semana 5 — sin señales de reapertura",
        "→ Persistente"
    ),
    (
        "ArtBo 2026 activación",
        "—",
        "—",
        "—",
        "NUEVO: Instituto de Visión confirma (se reportó como agosto — INCORRECTO)",
        "CORRECCIÓN: ArtBo es 24–27 SEP. Casas Riegner confirmado. 13 semanas.",
        "↑ Activándose"
    ),
    (
        "Perrotin / artistas mexicanos en EEUU",
        "—",
        "—",
        "—",
        "Gabriel Rico LA + Hugo Toro NY iniciados",
        "CONTINÚA — Ambas muestras activas. Señal de demanda coleccionistas",
        "→ Activo"
    ),
    (
        "Beatriz González (legado Casas Riegner)",
        "—",
        "Fallecimiento anunciado (9 ene 2026)",
        "Oslo retrospectiva confirmada",
        "Oslo inaugura 12 jun",
        "Semana 3 Oslo. Tour póstumo: SP → Barbican → Oslo (hasta oct)",
        "→ Estable, largo plazo"
    ),
]

for i, row in enumerate(trends, start=1):
    for j, val in enumerate(row):
        ws3.write(i, j, val, wrap)

wb.close()

# Verify
size = os.path.getsize(OUTPUT_PATH)
print(f"File created: {OUTPUT_PATH}")
print(f"File size: {size} bytes")
assert size > 0, "File is empty!"
print("Validation passed.")
