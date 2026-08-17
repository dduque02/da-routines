#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Weekly Gallery Intelligence report builder — Galería Duque Arango
Report window: 2026-08-10 to 2026-08-17
"""
import xlsxwriter

OUT_PATH = "/home/user/da-routines/reports/Gallery_Intelligence_2026-08-17.xlsx"

wb = xlsxwriter.Workbook(OUT_PATH)

# ---------- Formats ----------
FONT = "Arial"

title_fmt = wb.add_format({
    "font_name": FONT, "font_size": 14, "bold": True, "font_color": "#1F2933",
})
subtitle_fmt = wb.add_format({
    "font_name": FONT, "font_size": 10, "italic": True, "font_color": "#52606D",
})
header_fmt = wb.add_format({
    "font_name": FONT, "font_size": 10, "bold": True, "font_color": "white",
    "bg_color": "#1F2933", "border": 1, "border_color": "#CBD2D9",
    "align": "left", "valign": "top", "text_wrap": True,
})
cell_fmt = wb.add_format({
    "font_name": FONT, "font_size": 10, "align": "left", "valign": "top",
    "text_wrap": True, "border": 1, "border_color": "#E4E7EB",
})
cell_fmt_alt = wb.add_format({
    "font_name": FONT, "font_size": 10, "align": "left", "valign": "top",
    "text_wrap": True, "border": 1, "border_color": "#E4E7EB", "bg_color": "#F5F7FA",
})
tag_new = wb.add_format({
    "font_name": FONT, "font_size": 10, "align": "center", "valign": "top",
    "text_wrap": True, "border": 1, "border_color": "#E4E7EB",
    "bold": True, "font_color": "#0B6E4F", "bg_color": "#E3F9E5",
})
tag_new_alt = wb.add_format({
    "font_name": FONT, "font_size": 10, "align": "center", "valign": "top",
    "text_wrap": True, "border": 1, "border_color": "#E4E7EB",
    "bold": True, "font_color": "#0B6E4F", "bg_color": "#E3F9E5",
})
tag_continua = wb.add_format({
    "font_name": FONT, "font_size": 10, "align": "center", "valign": "top",
    "text_wrap": True, "border": 1, "border_color": "#E4E7EB",
    "bold": True, "font_color": "#8A6D3B", "bg_color": "#FCF3CF",
})
tag_cambio = wb.add_format({
    "font_name": FONT, "font_size": 10, "align": "center", "valign": "top",
    "text_wrap": True, "border": 1, "border_color": "#E4E7EB",
    "bold": True, "font_color": "#1B4F72", "bg_color": "#D6EAF8",
})
tag_na = wb.add_format({
    "font_name": FONT, "font_size": 10, "align": "center", "valign": "top",
    "text_wrap": True, "border": 1, "border_color": "#E4E7EB",
    "font_color": "#7B8794", "bg_color": "#F5F7FA",
})
bullet_fmt = wb.add_format({
    "font_name": FONT, "font_size": 11, "align": "left", "valign": "top", "text_wrap": True,
})
section_header_fmt = wb.add_format({
    "font_name": FONT, "font_size": 12, "bold": True, "font_color": "white",
    "bg_color": "#334E68", "border": 1,
})

CLASS_FMT = {
    "NUEVO": tag_new, "CAMBIÓ": tag_cambio, "CONTINÚA": tag_continua,
    "DESAPARECIÓ": tag_cambio, "N-A": tag_na,
}

# =========================================================================
# TAB 1: Changes & Findings
# =========================================================================
ws1 = wb.add_worksheet("Changes & Findings")
ws1.hide_gridlines(2)
ws1.set_column("A:A", 20)   # Gallery
ws1.set_column("B:B", 7)    # Tier
ws1.set_column("C:C", 20)   # Section
ws1.set_column("D:D", 16)   # Type of Change
ws1.set_column("E:E", 55)   # Description
ws1.set_column("F:F", 42)   # Strategic Note
ws1.set_column("G:G", 30)   # URL

ws1.merge_range("A1:G1", "Galería Duque Arango — Weekly Gallery Intelligence", title_fmt)
ws1.merge_range("A2:G2", "Semana del 2026-08-17 (período de hallazgos: 10–17 ago 2026) · 16 galerías monitoreadas", subtitle_fmt)
ws1.set_row(0, 22)
ws1.set_row(1, 16)

headers = ["Gallery", "Tier", "Section", "Type of Change", "Description",
           "Strategic Note for Duque Arango", "URL"]
row0 = 3
for c, h in enumerate(headers):
    ws1.write(row0, c, h, header_fmt)
ws1.set_row(row0, 20)
ws1.freeze_panes(row0 + 1, 0)

rows = [
    # Tier 1
    ("Galería La Cometa", 1, "Exhibiciones (Bogotá)", "New Exhibition",
     "Muestra colectiva \"Apocalypse now o la fabricación del paraíso\" — 11 artistas colombianos "
     "(Miguel Ángel Rojas, Clemencia Echeverri, Maria Fernanda Cardoso, José Alejandro Restrepo, "
     "Juan Fernando Herrán, Liliana Angulo Cortés, Alberto Baraya, Fernando Arias, Miler Lagos, "
     "Nicolás Consuegra, Julieth Morales). Inaugurada 5 ago, hasta 12 sep 2026. Cierra el vacío de "
     "programación de 5+ semanas consecutivas que se venía reportando desde julio.",
     "La Cometa vuelve a tener programación activa en Bogotá, pero su hub de noticias no se actualiza "
     "desde octubre 2025 — ventana clara para que DA gane el ciclo de prensa sobre \"nueva exposición\" "
     "en el mismo momento con mejor cobertura editorial.",
     "https://galerialacometa.com/exhibiciones/bogota/apocalypse-now-o-la-fabricacion-del-paraiso-bogota-2026-es",
     "CAMBIÓ"),
    ("Galería La Cometa", 1, "Exhibiciones (Bogotá)", "New Exhibition",
     "Muestra individual \"El volumen que tocó la luz\" — Dámaxo Henao (nacido en Medellín), óleo y "
     "técnica mixta, 5 obras, motivo de puertas/umbrales pintados. Inaugurada 5 ago, hasta 12 sep 2026 "
     "(misma apertura que la muestra colectiva).",
     "Segunda apertura simultánea que refuerza el mismo punto: reactivación de programación sin "
     "respaldo editorial visible.",
     "https://galerialacometa.com/exhibiciones/bogota/el-volumen-que-toco-la-luz-damaxo-henao-es",
     "CAMBIÓ"),
    ("Galería La Cometa", 1, "Noticias", "Operational",
     "El hub de noticias/blog sigue sin publicar contenido nuevo desde octubre 2025 (10+ meses), pese "
     "a las 2 aperturas de esta semana. Sin comunicado de prensa ni post de blog acompañando las nuevas "
     "muestras.",
     "Confirma la brecha de contenido/prensa como oportunidad concreta y medible: DA puede publicar y "
     "posicionar su propio contenido de \"nueva exposición\" más rápido que La Cometa esta semana.",
     "https://galerialacometa.com/noticias/",
     "CONTINÚA"),
    ("Galería Casa Cuadrada", 1, "Exposiciones", "Operational",
     "Página de exposiciones dedicada responde ahora con error de servidor HTTP 500 (antes solo "
     "aparecía vacía/desactualizada) — falla técnica nueva esta semana. Sección de noticias sigue "
     "estancada desde noviembre 2023.",
     "Casa Cuadrada sigue sin representar ninguna amenaza real; su brecha de visibilidad SEO en "
     "Medellín/Bogotá permanece disponible para que DA la absorba.",
     "http://galeriacasacuadrada.com",
     "CONTINÚA"),
    ("Casas Riegner", 1, "Sitio completo", "Other",
     "No activity detected this week. Programación vigente (Potosí, Fruto Animal, Structures of Time) "
     "sin cambios desde marzo–junio; ausencia de Óscar Murillo en el roster público se reconfirma sin "
     "cambios (resuelto como no-amenaza desde 27 jul).",
     "Sin amenaza nueva esta semana; se puede reducir la frecuencia de auditoría completa el próximo "
     "ciclo.",
     "https://www.casasriegner.com",
     "N-A"),
    ("Galería El Museo", 1, "Sitio completo", "Operational",
     "Sitio bloqueado (HTTP 403) en homepage y subpáginas de exposición — 2ª semana consecutiva de "
     "bloqueo, iniciado el 3 de agosto. \"Pintura Inmortal\" (25 jul–29 ago) presumiblemente sigue "
     "vigente según fuentes de terceros, sin poder confirmarse directamente.",
     "Si el bloqueo persiste, El Museo también pierde visibilidad ante crawlers de búsqueda — un "
     "riesgo de ranking que DA no enfrenta. Verificar vía fuentes terceras (Artsy, ARTEINFORMADO) el "
     "próximo ciclo.",
     "https://www.galeriaelmuseo.com",
     "CONTINÚA"),
    # Tier 2
    ("Latin Art Core", 2, "Sitio completo", "Operational",
     "Sitio bloqueado (HTTP 403) — 6ª+ semana consecutiva sin verificación directa posible. Búsqueda "
     "de respaldo no encontró prensa nueva en los últimos 7 días; solo listados de directorio genéricos.",
     "Punto ciego persistente en un competidor con foco cubano/latinoamericano que solapa con el "
     "roster de maestros modernos de DA (Lam, Tamayo). Se recomienda método de verificación alterno.",
     "https://latinartcore.com",
     "CONTINÚA"),
    ("Art of the World Gallery", 2, "Homepage / Exposiciones", "Operational",
     "La exposición \"Mimetism: Echoes that Breathe\" (Karla de Lara) sigue anunciada como vigente en "
     "el homepage pese a que su rango declarado (4 abr–6 jun) terminó hace 2.5 meses; sin exposición "
     "nueva ni contenido de agosto.",
     "Su inventario incluye a Wifredo Lam y Rufino Tamayo (roster DA) — un sitio visiblemente "
     "desactualizado en esos nombres es una oportunidad SEO directa para DA en el mercado de Houston.",
     "https://www.artoftheworldgallery.com",
     "NUEVO"),
    ("Ascaso Gallery", 2, "Sitio completo", "Operational",
     "Sitio con fallo de renderizado (contenido vacío vía verificación estándar) — persiste. Sin "
     "evidencia nueva sobre Botero, Cruz-Diez o Julio Larraz; el programa \"Reciprocity\" (Larraz + "
     "nueva generación) ya cerró hace 9+ meses.",
     "Amenaza confirmada baja por 5ª+ semana consecutiva. El formato editorial \"legado + nueva "
     "generación\" de Ascaso es replicable por DA con su propio roster (Larraz, Obregón) antes de que "
     "Ascaso lo repita.",
     "https://www.ascasogallery.com",
     "CONTINÚA"),
    ("Galería Freites", 2, "Blog", "New Content",
     "Post de blog sobre el préstamo de 2 obras a la retrospectiva \"Louise Nevelson. Mrs. N's Palace\" "
     "en el Centre Pompidou-Metz (vigente hasta 31 ago 2026). Post fechado 7 ago — 3 días antes del "
     "corte estricto de 7 días; se reporta con transparencia, no se cuenta como confirmado-nuevo.",
     "Freites juega una estrategia de prestigio institucional (préstamos a museos) en vez de anuncios "
     "comerciales — un modelo replicable para DA con obras de Botero o Cruz-Diez.",
     "https://galeriafreites.com/blog",
     "CAMBIÓ"),
    ("Opera Gallery", 2, "Exposiciones (próximas)", "New Exhibition",
     "\"Fernando Botero: In Praise of an Amplified World\" — Opera Gallery Londres, 12 oct–12 nov 2026. "
     "Primera exposición individual de Botero en Londres en más de una década, coincide con Frieze "
     "London; abarca casi 6 décadas de obra. Fecha exacta del anuncio no confirmable dentro de la "
     "ventana de 7 días, pero es la observación de estado actual más relevante del ciclo.",
     "AMENAZA DIRECTA de alto perfil sobre el artista ancla de DA (Botero) en escenario internacional "
     "(Frieze). Se recomienda que DA publique/actualice contenido propio sobre Botero y considere "
     "programación de contraataque (editorial, sala de exhibición online, contacto con prensa) antes "
     "del 12 de octubre.",
     "https://www.operagallery.com/event/fernando-botero-london",
     "CAMBIÓ"),
    ("Opera Gallery", 2, "Ferias", "Fair Announcement",
     "\"Texas Contemporary\" (Houston, 17 sep–8 oct 2026) listada entre programación próxima; fecha de "
     "anuncio no confirmada dentro de la ventana.",
     "Seguimiento de baja prioridad frente al hallazgo de Botero en Londres.",
     "https://www.operagallery.com",
     "N-A"),
    # Tier 3
    ("Gagosian", 3, "Sitio completo", "Operational",
     "Sitio bloqueado (HTTP 403) — 5ª semana consecutiva. Sin hallazgo atribuible a Gagosian esta "
     "semana (la exposición \"Botero in New York\" es de Sotheby's, no de Gagosian, y es de mediados "
     "de julio — fuera de ventana).",
     "Bloqueo persistente en un competidor T3 de alto perfil; método de acceso alterno recomendado "
     "antes de concluir ausencia de actividad.",
     "https://gagosian.com",
     "CONTINÚA"),
    ("David Zwirner", 3, "Exposiciones", "New Exhibition",
     "Óscar Murillo (artista del roster DA) — muestra individual \"Unfinished thinking: Paintings and "
     "Social Maps\" programada en París, apertura 19 oct 2026; confirmado como el único artista "
     "claramente latinoamericano en el roster (~90 artistas). Fecha del anuncio no confirmada dentro "
     "de la ventana de 7 días.",
     "Comparable directo de precio/prensa para un artista del roster DA — vale la pena monitorear "
     "cobertura previa a la apertura de París.",
     "https://www.davidzwirner.com/exhibitions",
     "CAMBIÓ"),
    ("Hauser & Wirth", 3, "Sitio completo", "Operational",
     "Sitio bloqueado (HTTP 429) — 5ª semana consecutiva. La muestra del Freud Museum sobre Leonora "
     "Carrington (contenido de terceros, no de H&W) cerró el 10 de agosto — límite exacto de la "
     "ventana — sin que H&W haya publicado contenido propio de cara al centenario 2027.",
     "H&W sigue sin reclamar la narrativa de Carrington pese a representar su patrimonio — oportunidad "
     "de contenido sostenida para DA.",
     "https://www.hauserwirth.com",
     "CONTINÚA"),
    ("Galerie Lelong", 3, "Artistas / Exposiciones", "New Content",
     "Jaume Plensa (roster DA) con 4 exposiciones institucionales simultáneas en 2026 (Montpellier, "
     "Denver Botanic Gardens, Murcia, Museum Küppersmühle) — todas previas a la ventana de 7 días, sin "
     "anuncio nuevo esta semana.",
     "Lelong posiciona activamente a Plensa en circuito museístico; útil como comparable de mercado, "
     "sin amenaza nueva esta semana.",
     "https://www.galerie-lelong.com/en/",
     "CONTINÚA"),
    ("Lisson Gallery", 3, "Acceso al sitio", "Operational",
     "Cambio de estado de acceso: el sitio es accesible vía solicitud HTTP estándar/navegador (200 OK) "
     "pero la herramienta de verificación automatizada sigue recibiendo 403 en todas las URLs — sugiere "
     "filtrado por firma de solicitud (bot/WAF), no una caída real de contenido.",
     "Se recomienda ajustar el método de verificación automatizada de la rutina para no perder "
     "visibilidad de Lisson (el sitio T3 de mayor señal para DA) en próximos ciclos.",
     "https://www.lissongallery.com",
     "CAMBIÓ"),
    ("Lisson Gallery", 3, "News (Olga de Amaral)", "New Content",
     "Confirmado vía fetch directo: el último contenido sobre Olga de Amaral sigue siendo el anuncio "
     "original de la retrospectiva Pulitzer (datePublished 17 abr 2026 en los metadatos de la página). "
     "Sin contenido nuevo — entra en su 8ª+ semana consecutiva, con la apertura del 10 de sept ahora a "
     "menos de un mes.",
     "La oportunidad de contenido/SEO más grande y sostenida de todo el ciclo sigue sin reclamar por "
     "Lisson, ahora dentro de la cuenta regresiva final de un mes — máxima prioridad de acción para DA "
     "esta semana.",
     "https://www.lissongallery.com/news/artist/olga-de-amaral",
     "CONTINÚA"),
    ("Lisson Gallery", 3, "Roster de artistas", "Other",
     "Julio Le Parc confirmado ausente del roster completo (89 artistas) — sin cambios respecto a "
     "semanas anteriores.",
     "Vacío de representación póstuma persiste, sin resolver.",
     "https://www.lissongallery.com/artists",
     "CONTINÚA"),
    ("Lisson Gallery", 3, "Exposiciones", "New Exhibition",
     "Exposición de Leonilson (maestro moderno brasileño, 1957–1993) sigue en el calendario (15 "
     "sep–24 oct 2026, Nueva York, organizada con la fundación del artista); sin artículo de prensa "
     "nuevo esta semana más allá del listado vigente.",
     "Confirma la tendencia de galerías internacionales formalizando alianzas con patrimonios de "
     "artistas latinoamericanos — vigilar para el propio roster secundario de DA.",
     "https://www.lissongallery.com/exhibitions/leonilson",
     "CONTINÚA"),
    ("Lehmann Maupin", 3, "Artistas / Noticias", "Other",
     "No activity detected this week. Roster (~55 artistas) sin coincidencias con el roster DA; "
     "artistas latinoamericanos existentes sin cambios (Teresita Fernández, Cecilia Vicuña, Loriel "
     "Beltrán, Adriana Varejão).",
     "Baja prioridad de seguimiento mientras no se agregue un artista latinoamericano nuevo al roster.",
     "https://www.lehmannmaupin.com/artists",
     "N-A"),
    ("Perrotin", 3, "Roster / Página legado", "Other",
     "Roster activo re-verificado sin coincidencias con el roster DA. Página legado de Julio Le Parc "
     "sigue sin mantener desde 2021 — condición sin cambios; 2ª galería (junto con Lisson) con este "
     "vacío de representación para un artista DA fallecido.",
     "Sin cambio en el vacío de representación de Le Parc; seguimiento de bajo esfuerzo el próximo "
     "ciclo.",
     "https://www.perrotin.com/en/artists/julio-le-parc",
     "CONTINÚA"),
]

r = row0 + 1
for i, row in enumerate(rows):
    gallery, tier, section, type_of_change, desc, note, url, cls = row
    fmt = cell_fmt_alt if i % 2 else cell_fmt
    ws1.write(r, 0, gallery, fmt)
    ws1.write(r, 1, tier, fmt)
    ws1.write(r, 2, section, fmt)
    ws1.write(r, 3, type_of_change, fmt)
    ws1.write(r, 4, desc, fmt)
    ws1.write(r, 5, note, fmt)
    ws1.write_url(r, 6, url, cell_format=fmt, string=url)
    ws1.set_row(r, 60)
    r += 1

# Add a Classification column (helps internal tracking; appended after URL, not in original spec
# columns but useful for continuity — placed clearly as supplemental)
ws1.write(row0, 7, "Classification (vs. últimas 4 semanas)", header_fmt)
ws1.set_column("H:H", 16)
r = row0 + 1
for i, row in enumerate(rows):
    cls = row[7]
    ws1.write(r, 7, cls, CLASS_FMT.get(cls, cell_fmt))
    r += 1

ws1.autofilter(row0, 0, r - 1, 7)

# =========================================================================
# TAB 2: Cross-Gallery Patterns
# =========================================================================
ws2 = wb.add_worksheet("Cross-Gallery Patterns")
ws2.hide_gridlines(2)
ws2.set_column("A:A", 40)
ws2.set_column("B:B", 30)
ws2.set_column("C:C", 45)
ws2.set_column("D:D", 45)

ws2.merge_range("A1:D1", "Cross-Gallery Patterns — Semana del 2026-08-17", title_fmt)
ws2.set_row(0, 22)

headers2 = ["Pattern Observed", "Galleries Involved", "What It May Signal",
            "Recommended Action for Duque Arango"]
row0b = 2
for c, h in enumerate(headers2):
    ws2.write(row0b, c, h, header_fmt)
ws2.set_row(row0b, 20)
ws2.freeze_panes(row0b + 1, 0)

patterns = [
    ("Ola simultánea de programación internacional de alto perfil sobre artistas del roster de "
     "DA (fuera de DA): Botero en Opera Gallery Londres (oct, coincide con Frieze London) y Óscar "
     "Murillo en David Zwirner París (oct).",
     "Opera Gallery, David Zwirner",
     "El mercado internacional de arte latinoamericano/blue-chip está en un momento de mayor "
     "visibilidad comercial — no es ruido aislado, son dos galerías T2/T3 distintas moviéndose sobre "
     "dos artistas DA distintos en el mismo mes.",
     "DA debería aprovechar este momentum con contenido propio y contacto de prensa sobre Botero y "
     "Murillo, sincronizado con octubre, en vez de dejar que la narrativa la controlen galerías "
     "externas."),
    ("El bloqueo de verificación (403/429/render vacío) dejó de ser un problema puntual y ahora es "
     "estructural: 5 sitios acumulan entre 2 y 6+ semanas consecutivas bloqueados (El Museo, Latin "
     "Art Core, Ascaso, Gagosian, Hauser & Wirth). Esta semana además se confirmó que Lisson — "
     "bloqueado desde el 27 jul — sí es accesible por navegador estándar; el bloqueo era específico a "
     "la firma de la herramienta automatizada, no al sitio.",
     "El Museo, Latin Art Core, Ascaso Gallery, Gagosian, Hauser & Wirth, Lisson Gallery",
     "Posible endurecimiento generalizado de protección anti-bot en el sector (tendencia de "
     "industria), no un problema específico de DA — pero también significa que las conclusiones de "
     "\"sin actividad\" en estos sitios tienen baja confianza mientras el bloqueo persista.",
     "Actualizar el método de verificación de la rutina (p. ej. solicitud con user-agent de navegador "
     "estándar, como reveló el caso Lisson) para no perder visibilidad sobre 5–6 de 16 competidores "
     "cada semana."),
    ("Dos competidores (La Cometa, Freites) reactivaron programación física/institucional esta "
     "semana — pero sin respaldo digital: el hub de noticias de La Cometa no se actualiza desde "
     "octubre 2025 pese a 2 aperturas simultáneas; Freites promueve un préstamo museístico solo vía un "
     "post de blog aislado.",
     "Galería La Cometa, Galería Freites",
     "La programación física y el marketing de contenidos están desacoplados en estos competidores — "
     "anuncian pero no amplifican. Es un patrón repetible, no un evento aislado de una sola galería.",
     "DA puede publicar cobertura propia, mejor optimizada para SEO, de sus propias aperturas y "
     "préstamos institucionales más rápido que estos competidores logran cubrir las suyas — ganando el "
     "ciclo de búsqueda por defecto."),
    ("Los vacíos de representación póstuma para artistas fallecidos siguen sin resolverse en "
     "múltiples galerías: Julio Le Parc ausente del roster activo tanto en Lisson (6ª+ semana) como en "
     "Perrotin (3ª+ semana), y la narrativa del centenario 2027 de Leonora Carrington sigue en manos "
     "de museos (Freud Museum, Musée du Luxembourg) y no de Hauser & Wirth, su galería representante.",
     "Lisson Gallery, Perrotin, Hauser & Wirth",
     "Las galerías comerciales grandes son lentas para reclamar la narrativa SEO/editorial de "
     "artistas fallecidos que representan — incluso cuando hay un evento institucional detonante "
     "(retrospectivas, centenarios). Es un patrón repetible, no específico de un solo artista.",
     "DA puede aplicar este mismo tipo de vigilancia a su propio roster de \"maestros modernos\" "
     "fallecidos (Botero, Obregón, Grau, Rayo, Negret, etc.) para asegurarse de que sea DA — no un "
     "museo o un tercero — quien controle la narrativa digital cuando llegue un evento similar."),
]

r = row0b + 1
for i, (pat, gals, signal, action) in enumerate(patterns):
    fmt = cell_fmt_alt if i % 2 else cell_fmt
    ws2.write(r, 0, pat, fmt)
    ws2.write(r, 1, gals, fmt)
    ws2.write(r, 2, signal, fmt)
    ws2.write(r, 3, action, fmt)
    ws2.set_row(r, 110)
    r += 1

ws2.autofilter(row0b, 0, r - 1, 3)

# =========================================================================
# TAB 3: Evolución 4 semanas
# =========================================================================
ws3 = wb.add_worksheet("Evolución 4 semanas")
ws3.hide_gridlines(2)
ws3.set_column("A:A", 100)

ws3.write("A1", "Evolución y Tendencias — Últimas 4 Semanas (+ esta semana)", title_fmt)
ws3.write("A2",
          "Baseline: reportes del 13 jul, 20 jul, 27 jul y 03 ago 2026 · Este reporte: 17 ago 2026",
          subtitle_fmt)
ws3.set_row(0, 22)
ws3.set_row(1, 16)

ws3.write(3, 0, "NOTA OPERATIVA IMPORTANTE", section_header_fmt)
ws3.set_row(3, 20)
gap_note = (
    "No se envió reporte semanal la semana del 10 de agosto de 2026 — el último reporte previo a "
    "este fue el del 3 de agosto. Esta edición cubre el hallazgo estricto de los últimos 7 días "
    "(10–17 ago) según la regla operativa del programa, pero el lector debe saber que hay una "
    "semana (03–10 ago) sin cobertura registrada entre este reporte y el anterior. Los conteos de "
    "\"semana consecutiva\" para condiciones en curso (p. ej. Olga de Amaral, bloqueos de acceso) "
    "asumen continuidad durante esa semana no reportada; no fueron verificados directamente para "
    "esas fechas."
)
ws3.write(4, 0, gap_note, bullet_fmt)
ws3.set_row(4, 90)

ws3.write(6, 0, "TENDENCIAS", section_header_fmt)
ws3.set_row(6, 20)

bullets = [
    "1. La Cometa rompe su vacío de programación de 5 semanas — el hallazgo que se venía marcando "
    "como \"único competidor T1 sin programación activa\" durante 4 reportes consecutivos (13 jul → "
    "03 ago) se resuelve el 5 de agosto con 2 aperturas simultáneas en Bogotá, pero sin ningún "
    "respaldo de prensa/blog (última noticia publicada: octubre 2025). El patrón se invierte: ya no "
    "es un vacío de programación, es un vacío de comunicación — y sigue siendo explotable.",

    "2. Olga de Amaral / Lisson entra en su 8ª+ semana consecutiva sin contenido nuevo (sin "
    "publicaciones desde el 17 de abril) — la ventana de contenido más larga y sostenida detectada en "
    "todo el ciclo (4ª → 5ª → 6ª → 7ª → 8ª semana consecutiva a través de los últimos 5 reportes) "
    "ahora coincide con la cuenta regresiva final de un mes para la apertura de la retrospectiva "
    "Pulitzer Arts Foundation \"Weaving the Infinite\" (10 sept) — la oportunidad de contenido/SEO más "
    "grande de todo el ciclo llega a su momento de máxima urgencia sin que Lisson la haya reclamado.",

    "3. Los bloqueos de acceso pasan de ser un problema aislado a un patrón estructural — de 4–5 "
    "sitios bloqueados (27 jul–03 ago) a 5 sitios que ahora acumulan entre 2ª y 6ª+ semana consecutiva "
    "de bloqueo (El Museo, Latin Art Core, Ascaso Gallery, Gagosian, Hauser & Wirth). Esta semana, "
    "además, se descubrió que Lisson — bloqueado desde el 27 jul — en realidad SÍ es accesible por "
    "navegador estándar; el bloqueo era específico a la firma de la herramienta de verificación "
    "automatizada, no al sitio en sí. Esto sugiere que el método de auditoría necesita actualizarse "
    "antes de seguir reportando \"sin actividad\" sobre sitios simplemente inaccesibles a la "
    "herramienta.",

    "4. Julio Le Parc acumula su 6ª+ semana de ausencia confirmada en Lisson y su 3ª+ semana "
    "confirmada en Perrotin, sin que ninguna galería comercial reclame su legado — el vacío de "
    "representación póstuma no solo persiste sino que se confirma estable en dos galerías "
    "simultáneamente por segundo mes consecutivo de seguimiento.",

    "5. Primera señal fuerte de \"calor\" comercial internacional sobre artistas del roster DA en 4 "
    "semanas de seguimiento: tras varias semanas de actividad mayormente tranquila en Tier 2/3 sobre "
    "nombres DA, esta semana aparecen simultáneamente Botero (Opera Gallery Londres, oct, coincide "
    "con Frieze) y Óscar Murillo (David Zwirner, París, oct) — dos aperturas internacionales de alto "
    "perfil sobre dos artistas distintos de DA en el mismo mes.",

    "6. Ascaso Gallery confirma su 5ª+ semana consecutiva sosteniendo el nivel de amenaza corregido a "
    "la baja desde el 13 de julio, sin evidencia nueva sobre Botero, Cruz-Diez o Julio Larraz — la "
    "corrección de julio sigue siendo válida un mes después.",
]

r = 8
for b in bullets:
    ws3.write(r, 0, b, bullet_fmt)
    ws3.set_row(r, 78)
    r += 1

wb.close()
print("OK ->", OUT_PATH)
