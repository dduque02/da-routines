import xlsxwriter

FILENAME = "Gallery_Intelligence_2026-07-13.xlsx"

wb = xlsxwriter.Workbook(FILENAME)

# ---------- formats ----------
header_fmt = wb.add_format({
    "bold": True, "bg_color": "#1a1a2e", "font_color": "#ffffff",
    "border": 1, "valign": "vcenter", "text_wrap": True,
})
cell_fmt = wb.add_format({"border": 1, "valign": "top", "text_wrap": True})
cell_fmt_bold = wb.add_format({"border": 1, "valign": "top", "text_wrap": True, "bold": True})

nuevo_fmt = wb.add_format({"border": 1, "valign": "top", "text_wrap": True, "bg_color": "#d4f7dc"})
cambio_fmt = wb.add_format({"border": 1, "valign": "top", "text_wrap": True, "bg_color": "#fff3cd"})
continua_fmt = wb.add_format({"border": 1, "valign": "top", "text_wrap": True, "bg_color": "#dbe9ff"})
desaparecio_fmt = wb.add_format({"border": 1, "valign": "top", "text_wrap": True, "bg_color": "#f0f0f0"})
sinactividad_fmt = wb.add_format({"border": 1, "valign": "top", "text_wrap": True, "bg_color": "#fafafa"})
correccion_fmt = wb.add_format({"border": 1, "valign": "top", "text_wrap": True, "bg_color": "#ffd9d9"})

CLASS_FMT = {
    "NUEVO": nuevo_fmt,
    "CAMBIÓ": cambio_fmt,
    "CONTINÚA": continua_fmt,
    "DESAPARECIÓ": desaparecio_fmt,
    "SIN ACTIVIDAD": sinactividad_fmt,
    "CORRECCIÓN": correccion_fmt,
    "N/A": cell_fmt,
}

def fmt_for(classification):
    return CLASS_FMT.get(classification, cell_fmt)

# ================== TAB 1: Changes & Findings ==================
ws1 = wb.add_worksheet("Changes & Findings")
cols1 = ["Gallery", "Tier", "Section", "Type of Change", "Description",
         "Strategic Note for Duque Arango", "Clasificación (4 sem.)", "URL"]
widths1 = [20, 8, 16, 18, 46, 42, 16, 30]
for i, w in enumerate(widths1):
    ws1.set_column(i, i, w)
ws1.set_row(0, 34)
for c, h in enumerate(cols1):
    ws1.write(0, c, h, header_fmt)

rows1 = [
# --- Tier 1 Colombia ---
["Galería La Cometa", "T1", "Exhibiciones", "Operational",
 "Las 4 exposiciones vigentes (Medellín: Aristizábal, Monzón; Madrid: Goldstein, Ospina) cerraron el 5 de julio. No se ha anunciado ninguna muestra de reemplazo; footer del sitio sigue con copyright '2024'.",
 "Vacío de programación visible en un competidor directo de Medellín/Bogotá. Ventana para que DA luzca más activa con contenido o programación propia esta semana.",
 "NUEVO", "https://galerialacometa.com/exhibiciones/"],
["Galería La Cometa", "T1", "La Galería (equipo)", "Operational",
 "Roster directivo confirmado estable (Esteban, Nicolás, Lucas y Paloma Jaramillo; Andrés Córdoba; Jorge Stefanell). No hay evidencia nueva del cambio de director artístico reportado el 22 de junio.",
 "El rumor de cambio de dirección sigue sin confirmarse — no actuar todavía, seguir monitoreando.",
 "CONTINÚA", "https://galerialacometa.com/la-galeria/"],
["Galería Casa Cuadrada", "T1", "Blog / Noticias", "Other",
 "Vía API de WordPress: el post más reciente sigue fechado en octubre 2025. 7ª semana consecutiva sin contenido nuevo confirmado.",
 "Silencio digital prolongado se mantiene. Sin amenaza inmediata; posible reactivación en agosto según patrón histórico.",
 "CONTINÚA", "https://galeriacasacuadrada.com/blog-fullscreen/"],
["Galería Casa Cuadrada", "T1", "Exposiciones (página)", "Operational",
 "La página de exposiciones tiene un timestamp de modificación del 9 de julio (dentro de la ventana), pero el contenido visible no cambió — ninguna muestra nueva aparece. Podría ser un touch técnico (plugin/caché) y no contenido real.",
 "Señal ambigua — no confirmar como reactivación todavía; revisar de nuevo la próxima semana para ver si se convierte en una muestra real.",
 "CAMBIÓ", "https://galeriacasacuadrada.com/portfolio-grid/"],
["Casas Riegner", "T1", "Artistas / Corrección de base", "Other",
 "El roster actual de 18 artistas verificado directamente NO incluye a Óscar Murillo; fuentes externas confirman que su representación actual es David Zwirner, Kurimanzutto y otras — no Casas Riegner.",
 "Se recomienda corregir el registro base: el supuesto solapamiento DA–Riegner vía Óscar Murillo no está sustentado esta semana. Dejar de monitorear esta amenaza específica salvo nueva evidencia.",
 "CORRECCIÓN", "https://www.casasriegner.com/artistas/"],
["Casas Riegner", "T1", "Exhibiciones / Noticias", "No Activity",
 "Exposición actual 'Fruto animal' (Liliana Sánchez) sigue sin cambios (5 jun–5 ago). Ciclo de prensa post-Art Basel confirmado cerrado, sin sucesor.",
 "Sin movimiento; próximo punto de control es ArtBo (24–27 sep 2026).",
 "CONTINÚA", "https://www.casasriegner.com/noticias/"],
["Galería El Museo", "T1", "Disponibilidad del sitio", "Operational",
 "Verificado en su totalidad esta semana: el sitio está en línea y funcionando con normalidad. El fallo 'no verificable' de la semana pasada fue un artefacto de la herramienta de rastreo (bloqueo de bot), no una caída real.",
 "Se resuelve la incógnita abierta la semana pasada — no hubo disrupción operativa real en El Museo.",
 "CAMBIÓ", "https://www.galeriaelmuseo.com/"],
["Galería El Museo", "T1", "Contenido sobre Edgar Negret", "Other",
 "Verificado vía API de WordPress: el sitio no ha publicado ni modificado contenido desde el 17–23 de junio (3 semanas). No existe ningún post sobre Negret del 29 de junio en el sitio — la 'efeméride' reportada la semana pasada probablemente fue solo una publicación en redes sociales, no contenido del sitio web.",
 "La amenaza de Negret no está escalando en el sitio propio de El Museo. Se puede bajar la urgencia de esta vía específica, aunque vale la pena capturar un pantallazo de la publicación en redes si aún existe, ya que Negret es palabra clave de DA.",
 "DESAPARECIÓ", "https://www.galeriaelmuseo.com/"],
["Galería El Museo", "T1", "Tienda online — Negret", "Inventory Update",
 "La tienda en línea de El Museo vende una escultura de Edgar Negret ('Sin título') con precio visible ($4.522.000 COP), dentro de un catálogo de +60 artistas con filtro por categoría.",
 "Infraestructura preexistente (no confirmada como nueva esta semana), pero relevante: El Museo vende obra de Negret con precio visible online, algo que DA no ofrece actualmente. Vale una revisión estratégica fuera del ciclo semanal.",
 "N/A", "https://galeriaelmuseo.com/tienda/categoria-producto/negret-edgar/"],

# --- Tier 2 ---
["Ascaso Gallery", "T2", "Disponibilidad del sitio / Corrección", "Operational",
 "Causa raíz del fallo de 2 semanas identificada: la home de ascasogallery.com no renderiza vía fetch automatizado (plantilla JS), pero TODAS las páginas internas (/exhibition/, /artist/, /fair/, /blog/) cargan con normalidad. El sitio está confirmado accesible y navegable esta semana.",
 "Se resuelve el vacío de monitoreo de 2 semanas. A futuro, monitorear vía subpáginas directas, no la raíz del dominio.",
 "CAMBIÓ", "https://www.ascasogallery.com/exhibition/"],
["Ascaso Gallery", "T2", "Muestra 'Reviver' — corrección", "Other",
 "Confirmado: 'Reviver' NO está activa (/exhibition/reviver/ da 404) y no es una muestra de artistas venezolanos. Es una muestra cerrada de 2022 del artista Andrew Hem (LA), sin relación con Botero, Cruz-Diez o Larraz. El encuadre de reportes anteriores estaba basado en información obsoleta o incorrecta.",
 "Corrección importante al registro de amenazas: retirar 'Reviver' como amenaza activa — es una muestra de hace 4 años, sin relevancia actual.",
 "CORRECCIÓN", "https://www.ascasogallery.com/exhibition/"],
["Ascaso Gallery", "T2", "Botero / Cruz-Diez / Larraz", "No Activity",
 "Sin contenido nuevo esta semana sobre estos tres artistas. Toda la cobertura encontrada es archivada (Art Miami dic-2025, 'Reciprocity' sep-2025, 'Forms in Space' dic-2025–feb-2026, todas ya cerradas). El blog no publica desde el 15 de octubre de 2025 (~9 meses).",
 "La caracterización de '5 semanas consecutivas de amenaza activa' en reportes previos no se sostiene con la evidencia actual — se recomienda bajar el nivel de amenaza de Ascaso hasta que aparezca actividad nueva confirmada.",
 "CORRECCIÓN", "https://www.ascasogallery.com/blog/"],
["Latin Art Core", "T2", "Sitio completo", "Operational",
 "Sitio devuelve HTTP 403 en todas las rutas probadas (raíz, /artists, /exhibitions, /news; con y sin www; http y https) de forma consistente — no es un fallo puntual, parece protección WAF/anti-bot.",
 "Bloqueo de acceso persistente — se recomienda un método de verificación alternativo (fetch autenticado o navegador) en el próximo ciclo en lugar de reintentos repetidos.",
 "N/A", "https://latinartcore.com"],
["Art of the World Gallery", "T2", "Homepage", "Operational",
 "La portada sigue destacando 'Mimetism: Echoes that Breathe' (Karla de Lara), muestra que cerró el 6 de junio — más de 5 semanas sin reemplazo visible.",
 "Portada desactualizada indica baja cadencia de contenido — señal de baja presión competitiva, pero también una oportunidad de posicionamiento SEO para DA en los 11 artistas que comparten roster (incl. Botero, Julio Larraz).",
 "CONTINÚA", "https://www.artoftheworldgallery.com"],
["Galería Freites", "T2", "Blog / Noticias", "No Activity",
 "Último post de blog fechado el 28 de mayo de 2026 — más de 6 semanas sin publicaciones nuevas pese a presencia activa en redes sociales.",
 "Freites comparte a Botero y Cruz-Diez con DA (mercado Caracas/Miami/Madrid) — vale mantenerla en watch permanente aunque esta semana esté en silencio.",
 "CONTINÚA", "https://galeriafreites.com/blog"],
["Opera Gallery", "T2", "Noticias / Ferias", "No Activity",
 "Nada cae estrictamente dentro de la ventana 06-jul–13-jul (instalación de Kenny Scharf en Mónaco y post de Kusama son del 2–3 de julio, justo antes del corte). Sala de exhibición virtual 'Botero' sigue activa como capacidad permanente, sin evidencia de lanzamiento nuevo esta semana.",
 "Sin actividad relevante esta semana. La sala de exhibición virtual dedicada a Botero de Opera es una referencia útil si DA evalúa construir su propia OVR.",
 "SIN ACTIVIDAD", "https://www.operagallery.com/news"],

# --- Tier 3 group 1 ---
["Gagosian", "T3", "Noticias / Exposiciones (LatAm)", "No Activity",
 "Sitio bloqueado para fetch directo (403 persistente); verificado por búsqueda. Sin exposiciones, incorporaciones de roster o prensa sobre artistas latinoamericanos en la ventana. Única artista LatAm relevante (Adriana Varejão) sin actividad — su muestra en Hispanic Society cerró en junio de 2025.",
 "Baja presión competitiva de Gagosian en posicionamiento latinoamericano esta semana.",
 "SIN ACTIVIDAD", "https://gagosian.com/news/"],
["David Zwirner", "T3", "Óscar Murillo (roster)", "No Activity",
 "Único artista latinoamericano del roster de Zwirner es Óscar Murillo; sin noticias nuevas en la ventana. Su muestra 'Collective Osmosis' (Potsdam) sigue en curso hasta el 9 de agosto (no es nueva). Muestra en París (19 oct) está anunciada pero sin fecha de publicación confirmada dentro de la ventana.",
 "Sin señal competitiva nueva esta semana; vigilar el anuncio formal de la muestra de París cuando se confirme.",
 "SIN ACTIVIDAD", "https://www.davidzwirner.com/artists/oscar-murillo/news"],
["Hauser & Wirth", "T3", "Leonora Carrington", "Other",
 "Sitio con rate-limiting persistente (429) — verificado por búsqueda. Sin contenido nuevo de Hauser & Wirth sobre Carrington esta semana, pero SÍ hay noticias de terceros: pintura redescubierta en Freud Museum Londres (hasta 10 ago), caso de robo/arresto de escultura en Ciudad de México (8 jul), 'Shape of Dreams' en L'Space NY (hasta 25 jul) — ninguna es programa de Hauser & Wirth.",
 "El momentum hacia el centenario 2027 de Carrington sigue construyéndose en el mercado, pero no lo está liderando Hauser & Wirth esta semana — ventana para que DA publique contenido editorial propio sobre Carrington y capture visibilidad antes de que la campaña de H&W se reactive.",
 "SIN ACTIVIDAD", "https://www.hauserwirth.com/artists/2707-leonora-carrington/"],
["Galerie Lelong", "T3", "Alfredo Jaar — evento Arles", "New Content",
 "Alfredo Jaar (representado por Lelong) participó en una conversación pública en Rencontres d'Arles el 9 de julio, dentro de la programación del 30º aniversario de la Maison Européenne de la Photographie, junto a Ming Smith y Park Chan-wook.",
 "Confirma que Lelong sigue posicionando activamente a su artista latinoamericano en programación de festivales europeos de primer nivel — referencia útil para DA al hablar de artistas latinoamericanos conceptuales/fotográficos en contenido editorial.",
 "NUEVO", "https://www.rencontres-arles.com/en/agenda/2026/alfredo-jaar-ming-smith-park-chan-wook-the-maison-europeenne-de-la-photographie-turns-30"],
["Galerie Lelong", "T3", "Exposición NY — reemplazo", "New Exhibition",
 "La muestra de Lucia Laguna (cerrada el 27 de junio) fue reemplazada el 9 de julio por 'For it is our battle', una exposición colectiva sin artistas latinoamericanos (Coyne, Tianmiao, Malani, Ono, Rosler, Sanpitak, Spero).",
 "Lelong no continuó su programación de artistas latinoamericanos en Nueva York — sugiere que su enfoque LatAm es intermitente, no una campaña sostenida. DA puede posicionarse como una fuente más consistente de atención de mercado secundario latinoamericano.",
 "CAMBIÓ", "https://www.galerie-lelong.com/en/expositions/"],

# --- Tier 3 group 2 ---
["Lisson Gallery", "T3", "Olga de Amaral — vacío SEO", "No Activity",
 "Sin contenido nuevo sobre Olga de Amaral en la ventana. Los últimos 10 ítems del feed de noticias de Lisson (10-jul a 23-jun) no mencionan a Amaral. El vacío de contenido desde el pivote de finales de junio entra en su 4ª semana consecutiva sin que nadie lo ocupe.",
 "Ventana SEO amplia y aún sin reclamar — oportunidad fuerte y sostenida para que DA publique contenido fresco sobre Olga de Amaral y capture cuota de búsqueda mientras el líder de categoría permanece en silencio.",
 "CONTINÚA", "https://www.lissongallery.com/artists/olga-de-amaral"],
["Lisson Gallery", "T3", "Julio Le Parc — ausencia total", "Other",
 "Julio Le Parc ya no aparece en ningún lugar del sitio de Lisson (ni directorio de artistas, ni exposiciones, ni página 'Artists in Venice 2026'). Confirmado: Le Parc falleció en París el 30 de mayo de 2026 (97 años), días antes de que abriera su retrospectiva póstuma en Tate Modern (11 jun 2026–3 may 2027). Su representación de mercado está con otras galerías (Perrotin, Galleria Continua, Sicardi, The Mayor Gallery, Nara Roesler), no con Lisson.",
 "Le Parc también es artista DA. Su fallecimiento y la retrospectiva de Tate (activa hasta mayo 2027) son un gancho noticioso vigente sin una voz de galería dominante en línea — junto con el vacío de Amaral, DA tiene DOS oportunidades SEO simultáneas sobre artistas compartidos sin reclamar.",
 "NUEVO", "https://www.lissongallery.com/exhibitions/venice-2026"],
["Lehmann Maupin", "T3", "Roster LatAm", "No Activity",
 "Artistas LatAm del roster (Beltrán, Fernández, OSGEMEOS, Vicuña) sin contenido dentro de la ventana estricta. Ítems más cercanos caen fuera del corte (muestra 'Second Nature' abrió 1 jul; retrospectiva de Cecilia Vicuña en Whitechapel anunciada 12 jun).",
 "Ninguno de los artistas LatAm de Lehmann Maupin se solapa con el roster de DA — baja relevancia competitiva directa esta semana. El perfil institucional creciente de Cecilia Vicuña (Whitechapel, oct 2026) vale un seguimiento a mediano plazo.",
 "SIN ACTIVIDAD", "https://www.lehmannmaupin.com/news"],
["Perrotin", "T3", "Roster LatAm", "No Activity",
 "Dos muestras de artistas latinoamericanos en curso (Hugo Toro en NY hasta 31 jul; Christiane Pooley en Shanghái hasta 15 ago) pero ambas abrieron antes de la ventana, sin prensa nueva esta semana. Ninguno de estos artistas se solapa con el roster de DA.",
 "Sin amenazas competitivas nuevas esta semana desde Perrotin.",
 "SIN ACTIVIDAD", "https://www.perrotin.com/en"],
]

r = 1
for row in rows1:
    classification = row[6]
    f = fmt_for(classification)
    for c, val in enumerate(row):
        ws1.write(r, c, val, f)
    r += 1

ws1.freeze_panes(1, 0)
ws1.autofilter(0, 0, r - 1, len(cols1) - 1)

# ================== TAB 2: Cross-Gallery Patterns ==================
ws2 = wb.add_worksheet("Cross-Gallery Patterns")
cols2 = ["Pattern Observed", "Galleries Involved", "What It May Signal", "Recommended Action for Duque Arango"]
widths2 = [34, 26, 42, 42]
for i, w in enumerate(widths2):
    ws2.set_column(i, i, w)
ws2.set_row(0, 30)
for c, h in enumerate(cols2):
    ws2.write(0, c, h, header_fmt)

rows2 = [
["Doble vacío SEO en artistas compartidos: Olga de Amaral (4ª semana) + Julio Le Parc (ausencia tras fallecimiento)",
 "Lisson Gallery",
 "El artista de mayor peso institucional para ambos temas (Lisson) se retiró de los dos frentes casi al mismo tiempo, dejando espacio de búsqueda desocupado en dos artistas DA simultáneamente — una coincidencia poco común.",
 "Publicar esta semana contenido editorial en inglés y español sobre Olga de Amaral y sobre Julio Le Parc (homenaje/retrospectiva de su obra, disponibilidad de piezas) para capturar ambas ventanas antes de que otro competidor las ocupe."],
["Corrección de amenaza sobrestimada: Ascaso Gallery",
 "Ascaso Gallery",
 "Una amenaza reportada como '5 semanas consecutivas de actividad activa' resultó estar basada en una muestra cerrada de 2022 y prensa de hace ~9 meses. Indica que el proceso de verificación de esta rutina puede propagar errores si no se revisan las fuentes primarias con suficiente frecuencia.",
 "Bajar el nivel de amenaza de Ascaso en el registro de seguimiento; mantenerla en monitoreo estándar (no prioritario) hasta que aparezca evidencia de actividad realmente nueva sobre Botero, Cruz-Diez o Larraz."],
["Vacíos de programación en competidores colombianos directos",
 "Galería La Cometa, Galería Casa Cuadrada, Casas Riegner (parcial)",
 "Tres de los cuatro competidores Tier 1 colombianos muestran baja actividad esta semana (La Cometa sin muestra de reemplazo, Casa Cuadrada en su 7ª semana de silencio, Riegner en fase de espera pre-ArtBo). El mercado local está relativamente tranquilo.",
 "Aprovechar la baja actividad competitiva local publicando programación o contenido editorial propio esta semana para ganar cuota de visibilidad en Medellín/Bogotá mientras los competidores directos están inactivos."],
["Corrección de solapamiento de roster: Óscar Murillo no aparece representado por Casas Riegner",
 "Casas Riegner",
 "El supuesto solapamiento DA–Riegner vía Óscar Murillo (citado en reportes anteriores) no se sustenta con el roster público actual ni con fuentes externas de representación del artista.",
 "Actualizar el registro base de amenazas: retirar a Casas Riegner como competidor directo en Óscar Murillo salvo que aparezca evidencia nueva."],
["Contenido latinoamericano intermitente, no sostenido, en galerías Tier 3",
 "Galerie Lelong, Perrotin, Lehmann Maupin, David Zwirner",
 "Ninguna galería internacional de referencia mantiene una campaña continua de contenido latinoamericano esta semana — Lelong reemplazó su única muestra LatAm (Lucia Laguna) por una colectiva sin artistas latinoamericanos; el resto está en pausa sobre sus artistas LatAm.",
 "DA puede diferenciarse posicionándose como la fuente más consistente y sostenida de contenido y mercado secundario latinoamericano, en contraste con la atención intermitente de las galerías internacionales."],
["Acceso bloqueado de forma persistente en dos sitios (403/429)",
 "Latin Art Core, Gagosian, Hauser & Wirth",
 "Protección anti-bot cada vez más agresiva en varios sitios está reduciendo la cobertura directa de la rutina; Latin Art Core lleva dos ciclos sin poder verificarse por ningún método probado.",
 "Evaluar un método de verificación alternativo (navegador headless autenticado o revisión manual puntual) para estos tres sitios en el próximo ciclo, priorizando Latin Art Core."],
]

r = 1
for row in rows2:
    for c, val in enumerate(row):
        ws2.write(r, c, val, cell_fmt)
    r += 1

ws2.freeze_panes(1, 0)
ws2.autofilter(0, 0, r - 1, len(cols2) - 1)

# ================== TAB 3: Evolución y tendencias (últimas 4 semanas) ==================
ws3 = wb.add_worksheet("Evolución 4 semanas")
cols3 = ["Tendencia", "Resumen de evolución (últimas 4 semanas)", "Estado esta semana (13 jul)"]
widths3 = [26, 60, 30]
for i, w in enumerate(widths3):
    ws3.set_column(i, i, w)
ws3.set_row(0, 30)
for c, h in enumerate(cols3):
    ws3.write(0, c, h, header_fmt)

rows3 = [
["Ascaso Gallery — amenaza revisada",
 "22 jun: semana 4 de actividad sostenida reportada. 29 jun: semana 5, cobertura de Cruz-Diez en inglés/español. 06 jul: sitio NO verificable por fallo técnico — amenaza posiblemente en semana 6. 13 jul: verificación completa revela que 'Reviver' es una muestra de 2022 y que no hay prensa nueva sobre Botero/Cruz-Diez/Larraz desde hace ~9 meses.",
 "CORREGIDO — la racha de '5 semanas de amenaza activa' no se sostiene con evidencia primaria. Bajar prioridad hasta nueva evidencia real."],
["Olga de Amaral / Lisson Gallery — ventana SEO",
 "15 jun: pico de campaña (3 ediciones de Art Basel + colaboración Akris). 22–29 jun: Lisson pivota a Julio Le Parc, vacío se abre. 06 jul: 3ª semana de vacío sin ocupar. 13 jul: confirmado 4ª semana consecutiva sin contenido nuevo de Amaral en Lisson ni en ningún competidor.",
 "ABIERTO — 4 semanas consecutivas de oportunidad de contenido sin reclamar por nadie."],
["Julio Le Parc — nueva capa de la historia",
 "15 jun: retrospectiva abre en Tate tras su fallecimiento reportado semanas antes. 22–29 jun: Lisson activa contenido de homenaje. 13 jul: Le Parc desaparece por completo del sitio de Lisson (retrospectiva de Tate sigue activa hasta mayo 2027); ninguna galería comercial domina su narrativa de búsqueda actualmente.",
 "NUEVO HALLAZGO — segunda ventana SEO simultánea a la de Amaral, sobre un artista también representado por DA."],
["Galería El Museo / Edgar Negret",
 "29 jun: publicó contenido biográfico sobre Negret (efeméride), flagged como amenaza directa. 06 jul: sitio no verificable, amenaza sin confirmar ni descartar. 13 jul: verificación vía API confirma que el sitio no ha publicado nada desde el 17–23 de junio — el contenido de Negret probablemente fue solo una publicación de redes sociales, no una campaña editorial en el sitio.",
 "DESESCALADO — la amenaza en el sitio propio de El Museo no se materializó; capturar evidencia de redes sociales para el registro."],
["Casa Cuadrada — silencio digital prolongado",
 "22 jun: 4ª semana de silencio. 29 jun: 5ª semana, se anticipa reactivación en agosto. 06 jul: 6ª semana confirmada. 13 jul: 7ª semana — sin contenido real nuevo, solo un touch de metadatos sin cambio visible el 9 de julio.",
 "CONTINÚA — vigilar el touch de metadatos de esta semana; podría anticipar una reactivación real la próxima semana."],
["Casas Riegner / Óscar Murillo",
 "22 jun–06 jul: reportado como solapamiento directo de roster con DA vía Óscar Murillo, con foco en el ciclo de prensa post-Art Basel. 13 jul: verificación directa del roster (18 artistas) y de fuentes externas no confirma a Murillo como artista representado por Riegner actualmente.",
 "CORREGIDO — se recomienda retirar este solapamiento del registro de amenazas activo salvo nueva evidencia."],
]

r = 1
for row in rows3:
    for c, val in enumerate(row):
        ws3.write(r, c, val, cell_fmt)
    r += 1

ws3.freeze_panes(1, 0)
ws3.autofilter(0, 0, r - 1, len(cols3) - 1)

wb.close()
print("OK", FILENAME)
