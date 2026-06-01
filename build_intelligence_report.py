#!/usr/bin/env python3
"""Weekly competitive intelligence report builder for Galería Duque Arango."""

import xlsxwriter
import os

DATE = "2026-06-01"
DATE_MINUS_7 = "2026-05-25"
FILENAME = f"Gallery_Intelligence_{DATE}.xlsx"
OUTPUT_PATH = os.path.join(os.path.dirname(__file__), FILENAME)

# ---------------------------------------------------------------------------
# DATA
# ---------------------------------------------------------------------------

FINDINGS = [
    # ── TIER 1 ────────────────────────────────────────────────────────────
    {
        "gallery": "Galería La Cometa",
        "tier": "1",
        "section": "Exhibitions",
        "type": "New Exhibition",
        "description": (
            "Multiple exhibitions running simultaneously: 'Envoltorios' by Asicaz Monzón "
            "and 'Nada Es Lo Que Parece, Al Parecer Desaparece II' by Luisa Aristizábal "
            "(both Medellín, through Jul 5, 2026); 'Parar el Mundo' by Adam Goldstein and "
            "'Pensamiento mágico / El año entrante' by Alejandro Ospina (both Madrid, "
            "May 9–Jul 5, 2026); 'Fragments of a Collection' (through Jun 13, 2026). "
            "A 'Coming Soon at La Cometa Madrid' piece was also published on Correo Cultural "
            "in May 2026 announcing additional upcoming programming."
        ),
        "strategic_note": (
            "La Cometa is running 5 simultaneous shows across 3 cities—dense multi-location "
            "programming that outpaces Duque Arango's current exhibition cadence. "
            "Consider announcing upcoming Jun–Jul programming now to stay visible in market."
        ),
        "url": "https://galerialacometa.com/exhibiciones/",
    },
    {
        "gallery": "Galería Casa Cuadrada",
        "tier": "1",
        "section": "—",
        "type": "Other",
        "description": "No activity detected this week.",
        "strategic_note": "—",
        "url": "http://galeriacasacuadrada.com",
    },
    {
        "gallery": "Casas Riegner",
        "tier": "1",
        "section": "Fairs",
        "type": "Fair Announcement",
        "description": (
            "Casas Riegner is confirmed among 290 exhibitors at Art Basel 2026 "
            "(Basel, Switzerland, Jun 16–21; VIP preview Jun 16–17). "
            "Their gallery exhibition 'Potosí' closed May 21, 2026."
        ),
        "strategic_note": (
            "Direct Colombian competitor will be visible to global collectors at the world's "
            "most important fair this month. Duque Arango should monitor which artists Casas "
            "Riegner shows in Basel and use the occasion to activate its own collector "
            "communications (newsletter, social) during fair week."
        ),
        "url": "https://www.artbasel.com/catalog/gallery/1082/Casas-Riegner",
    },
    {
        "gallery": "Galería El Museo",
        "tier": "1",
        "section": "Exhibitions / Fairs",
        "type": "New Exhibition",
        "description": (
            "'Aurora Lario: Cartografías del cuerpo' on view May 9–Jun 6, 2026 "
            "(closing this week). Gallery also participated in Pinta Lima 2026 "
            "(May 9–Jun 6). No new openings confirmed after May 25."
        ),
        "strategic_note": (
            "El Museo's current show closes Jun 6—their programming gap opens immediately "
            "after. Watch for their next opening announcement; their Pinta Lima presence "
            "signals growing fair activity in the Latin American circuit."
        ),
        "url": "https://www.galeriaelmuseo.com/exposiciones/expo-actual/",
    },
    # ── TIER 2 ────────────────────────────────────────────────────────────
    {
        "gallery": "Latin Art Core",
        "tier": "2",
        "section": "Exhibitions",
        "type": "New Exhibition",
        "description": (
            "Announced solo exhibition 'Manuel Mendive: With the new day the sun shines "
            "and leads us' opening Jun 19, 2026 at Latin Art Core, Miami (Little Havana). "
            "46 works spanning 2023–2025 plus key earlier pieces. Mendive will attend "
            "the opening reception. The show explores Yoruba cosmology, spirituality, "
            "and the human relationship with nature."
        ),
        "strategic_note": (
            "Latin Art Core is promoting Cuban master Mendive with a high-profile solo show "
            "in Miami, reinforcing the Latin American masters market. Duque Arango should "
            "consider activating content around its own modern masters (Botero, Obregón, "
            "Grau) ahead of the Miami summer season to capture collector attention."
        ),
        "url": "https://latinartcore.com/exhibition/manuel-mendive-with-the-new-day-the-sun-shines-and-leads-us/",
    },
    {
        "gallery": "Art of the World Gallery",
        "tier": "2",
        "section": "Exhibitions",
        "type": "New Exhibition",
        "description": (
            "'Mimetism: Echoes that Breathe' by Mexican artist Karla de Lara "
            "(Apr 4–Jun 6, 2026, Houston) is in its final days this week. "
            "Show presents layered figurative-symbolic compositions exploring memory, "
            "identity, and place."
        ),
        "strategic_note": (
            "Gallery is actively programming Mexican contemporary artists in Houston. "
            "No direct overlap with Duque Arango's roster this week; monitor for next "
            "opening announcement."
        ),
        "url": "https://www.artoftheworldgallery.com/",
    },
    {
        "gallery": "Ascaso Gallery",
        "tier": "2",
        "section": "Exhibitions",
        "type": "New Exhibition",
        "description": (
            "Announced 'Reviver', a solo exhibition by Los Angeles artist Andrew Hem, "
            "opening Jun 30, 2026 at Ascaso Gallery, Miami (Wynwood). Gallery continues "
            "to represent Botero, Cruz-Diez, and Jesús Rafael Soto as part of its "
            "Venezuelan/Latin masters program."
        ),
        "strategic_note": (
            "ALERT: Ascaso Gallery actively represents Fernando Botero and Carlos Cruz-Diez "
            "—both Duque Arango roster artists. Their secondary-market presence for these "
            "artists in Miami is a direct competitive threat. Duque Arango should reinforce "
            "its primary/exclusive positioning for these artists in Colombia and strengthen "
            "online content around them."
        ),
        "url": "https://www.ascasogallery.com/",
    },
    {
        "gallery": "Galería Freites",
        "tier": "2",
        "section": "—",
        "type": "Other",
        "description": "No activity detected this week. Last known show: Jacobo Borges (opened Apr 19, 2026).",
        "strategic_note": "—",
        "url": "https://galeriafreites.com/",
    },
    {
        "gallery": "Opera Gallery",
        "tier": "2",
        "section": "Exhibitions",
        "type": "New Exhibition",
        "description": (
            "London location opened 'Feng Xiao-Min | Pieter Obels' on Jun 4, 2026 "
            "(through Jul 5), timed to coincide with London Gallery Weekend (Jun 5–7). "
            "Concurrent shows active in Paris (through Jun 17), Singapore (through Jul 7), "
            "and Madrid (through Jun 20). Also announced: Brazilian artist Gustavo Nazareno "
            "solo show at Paris location, Jun 26–Jul 15, 2026."
        ),
        "strategic_note": (
            "Opera Gallery is running 4 simultaneous shows across continents and is "
            "activating London Gallery Weekend for visibility. The Nazareno announcement "
            "signals ongoing commitment to Latin American artists at Tier 2 international "
            "level. Duque Arango could leverage London Gallery Weekend content for social "
            "media without attending."
        ),
        "url": "https://www.operagallery.com/exhibitions",
    },
    # ── TIER 3 ────────────────────────────────────────────────────────────
    {
        "gallery": "Gagosian",
        "tier": "3",
        "section": "—",
        "type": "Other",
        "description": "No Latin American artist or program activity detected this week.",
        "strategic_note": "—",
        "url": "https://gagosian.com/exhibitions/",
    },
    {
        "gallery": "David Zwirner",
        "tier": "3",
        "section": "Institutional",
        "type": "Institutional",
        "description": (
            "Óscar Murillo's exhibition 'Collective Osmosis' (Mar 14–Aug 9, 2026) continues "
            "at DAS MINSK Kunsthaus and Museum Barberini, Potsdam, Germany. Works placed in "
            "dialogue with Claude Monet; public 'Collective Painting' launched Apr 25 on "
            "DAS MINSK terrace. Zwirner-represented artist maintains high institutional "
            "visibility in Europe."
        ),
        "strategic_note": (
            "ROSTER OVERLAP: Óscar Murillo, represented by David Zwirner and on Duque "
            "Arango's roster, is generating significant institutional press in Germany "
            "through summer 2026. Duque Arango should amplify this visibility in Colombia "
            "by publishing content about Murillo's current institutional recognition."
        ),
        "url": "https://dasminsk.de/en/exhibitions/8017/oscar_murillo",
    },
    {
        "gallery": "Hauser & Wirth",
        "tier": "3",
        "section": "—",
        "type": "Other",
        "description": "No Latin American artist or program activity detected this week.",
        "strategic_note": "—",
        "url": "https://www.hauserwirth.com/hauser-wirth-exhibitions/",
    },
    {
        "gallery": "Galerie Lelong",
        "tier": "3",
        "section": "—",
        "type": "Other",
        "description": (
            "No new Latin American activity detected this week. Note: Galerie Lelong "
            "presented Olga de Amaral at Art Basel Qatar (Feb 5–7, 2026) and represents "
            "key Latin American figures (Cildo Meireles, Ana Mendieta, Alfredo Jaar). "
            "No post-May 25 announcement found."
        ),
        "strategic_note": "—",
        "url": "https://galerielelong.com/exhibitions/",
    },
    {
        "gallery": "Lisson Gallery",
        "tier": "3",
        "section": "Institutional",
        "type": "Institutional",
        "description": (
            "Olga de Amaral's work continues in 'Made in Tension' at The Soloviev "
            "Foundation Gallery, New York (Mar–Dec 2026)—a long-term institutional "
            "placement. Lisson also presented her at Art Basel Qatar (Feb 2026). "
            "No new Lisson show confirmed after May 25."
        ),
        "strategic_note": (
            "ROSTER OVERLAP: Olga de Amaral (Duque Arango roster) has sustained "
            "institutional presence in New York through Lisson Gallery all year. "
            "Duque Arango should ensure its Amaral content and artist page are current "
            "to capture search traffic from collectors following her U.S. momentum."
        ),
        "url": "https://www.lissongallery.com/artists/olga-de-amaral",
    },
    {
        "gallery": "Lehmann Maupin",
        "tier": "3",
        "section": "—",
        "type": "Other",
        "description": "No Latin American artist or program activity detected this week.",
        "strategic_note": "—",
        "url": "https://www.lehmannmaupin.com/exhibitions",
    },
    {
        "gallery": "Perrotin",
        "tier": "3",
        "section": "—",
        "type": "Other",
        "description": (
            "No new Latin American activity detected this week. Last Latin American "
            "show was Gabriel de la Mora's 'Repeated Original' (ended Apr 11, 2026)."
        ),
        "strategic_note": "—",
        "url": "https://www.perrotin.com/exhibitions/current",
    },
]

PATTERNS = [
    {
        "pattern": "Art Basel 2026 fair week imminent (Jun 16–21): Colombian gallery Casas Riegner confirmed among 290 exhibitors",
        "galleries": "Casas Riegner",
        "signal": (
            "The world's most important art fair opens in two weeks. Colombian galleries "
            "are represented at the highest level, signaling the market's continued "
            "recognition of Latin American contemporary art."
        ),
        "action": (
            "Activate a Basel-week content plan: collector newsletter, social posts, "
            "and a press release highlighting Duque Arango's own artists relevant to "
            "the Basel collector profile. Position as the authoritative Colombian source "
            "for collectors who attend Basel."
        ),
    },
    {
        "pattern": "Roster-artist international institutional momentum: Óscar Murillo (David Zwirner / DAS MINSK) and Olga de Amaral (Lisson / Soloviev Foundation) generating sustained European and U.S. visibility",
        "galleries": "David Zwirner, Lisson Gallery",
        "signal": (
            "Two artists on Duque Arango's roster are receiving high-profile institutional "
            "coverage by major international galleries right now—creating a halo of "
            "international credibility that the gallery can leverage."
        ),
        "action": (
            "Publish editorial content on Murillo and Amaral referencing their current "
            "institutional shows. This strengthens Duque Arango's SEO for these artist "
            "names while positioning the gallery as the authoritative Colombian voice "
            "on their work."
        ),
    },
    {
        "pattern": "Competitor roster overlap with Botero and Cruz-Diez: Ascaso Gallery (Miami) actively markets these artists in secondary market",
        "galleries": "Ascaso Gallery",
        "signal": (
            "Ascaso's open representation of Botero and Cruz-Diez in Miami means "
            "U.S. collectors searching for these artists may encounter Ascaso first. "
            "Mimetic desire and authority bias both favor whichever gallery appears "
            "first and most authoritatively."
        ),
        "action": (
            "Prioritize SEO and AI-search content for 'Fernando Botero' and 'Carlos "
            "Cruz-Diez' to ensure Duque Arango ranks before Ascaso in Colombian and "
            "Latin American search contexts. Consider artist-specific landing pages "
            "with biography, works, and exhibition history."
        ),
    },
    {
        "pattern": "Miami Latin American masters season: Latin Art Core (Mendive, Jun 19) and Ascaso (Botero/Cruz-Diez ongoing) both programming the Latin American masters category simultaneously",
        "galleries": "Latin Art Core, Ascaso Gallery",
        "signal": (
            "Miami Tier 2 galleries are doubling down on Latin American modern masters "
            "through June–July, signaling sustained collector demand for this category "
            "in the U.S. market."
        ),
        "action": (
            "Position Duque Arango as the primary Colombian source for Latin American "
            "masters. Publish editorial content timed to the Mendive opening week "
            "(Jun 19) connecting Cuban and Colombian modern art—capture search traffic "
            "from the Mendive conversation while highlighting Duque Arango's own masters."
        ),
    },
]

# ---------------------------------------------------------------------------
# BUILD EXCEL
# ---------------------------------------------------------------------------

workbook = xlsxwriter.Workbook(OUTPUT_PATH)

# ── Formats ────────────────────────────────────────────────────────────────
header_fmt = workbook.add_format({
    "bold": True,
    "font_color": "#FFFFFF",
    "bg_color": "#1A1A2E",
    "border": 1,
    "text_wrap": True,
    "valign": "vcenter",
    "align": "center",
    "font_size": 10,
})
t1_fmt = workbook.add_format({
    "bg_color": "#F2E6FF",
    "border": 1,
    "text_wrap": True,
    "valign": "top",
    "font_size": 9,
})
t2_fmt = workbook.add_format({
    "bg_color": "#E6F0FF",
    "border": 1,
    "text_wrap": True,
    "valign": "top",
    "font_size": 9,
})
t3_fmt = workbook.add_format({
    "bg_color": "#F0F0F0",
    "border": 1,
    "text_wrap": True,
    "valign": "top",
    "font_size": 9,
})
no_activity_fmt = workbook.add_format({
    "italic": True,
    "font_color": "#888888",
    "border": 1,
    "text_wrap": True,
    "valign": "top",
    "font_size": 9,
})
pattern_header_fmt = workbook.add_format({
    "bold": True,
    "font_color": "#FFFFFF",
    "bg_color": "#16213E",
    "border": 1,
    "text_wrap": True,
    "valign": "vcenter",
    "align": "center",
    "font_size": 10,
})
pattern_row_fmt = workbook.add_format({
    "bg_color": "#EAF4EA",
    "border": 1,
    "text_wrap": True,
    "valign": "top",
    "font_size": 9,
})
title_fmt = workbook.add_format({
    "bold": True,
    "font_size": 14,
    "font_color": "#1A1A2E",
})
subtitle_fmt = workbook.add_format({
    "italic": True,
    "font_size": 10,
    "font_color": "#555555",
})

# ── TAB 1: Changes & Findings ───────────────────────────────────────────────
ws1 = workbook.add_worksheet("Changes & Findings")
ws1.set_zoom(90)

# Title row
ws1.merge_range("A1:G1", f"🖼️  Gallery Competitive Intelligence — {DATE}", title_fmt)
ws1.merge_range("A2:G2", f"Galería Duque Arango  |  Period: {DATE_MINUS_7} → {DATE}", subtitle_fmt)
ws1.set_row(0, 22)
ws1.set_row(1, 16)

# Headers (row index 2 = row 3)
HEADERS_1 = ["Gallery", "Tier", "Section", "Type of Change", "Description",
             "Strategic Note for Duque Arango", "URL"]
COL_WIDTHS_1 = [22, 6, 18, 20, 60, 50, 40]
for col, (h, w) in enumerate(zip(HEADERS_1, COL_WIDTHS_1)):
    ws1.write(2, col, h, header_fmt)
    ws1.set_column(col, col, w)
ws1.set_row(2, 32)

# Data rows
TIER_FMTS = {"1": t1_fmt, "2": t2_fmt, "3": t3_fmt}
for row_idx, f in enumerate(FINDINGS):
    r = row_idx + 3  # offset by title + header rows
    tier = f["tier"]
    is_no_activity = f["description"] == "No activity detected this week."
    cell_fmt = no_activity_fmt if is_no_activity else TIER_FMTS.get(tier, t3_fmt)
    values = [f["gallery"], f["tier"], f["section"], f["type"],
              f["description"], f["strategic_note"], f["url"]]
    for col, val in enumerate(values):
        ws1.write(r, col, val, cell_fmt)
    ws1.set_row(r, 90)

ws1.freeze_panes(3, 0)
ws1.autofilter(2, 0, 2 + len(FINDINGS) - 1, len(HEADERS_1) - 1)

# ── TAB 2: Cross-Gallery Patterns ──────────────────────────────────────────
ws2 = workbook.add_worksheet("Cross-Gallery Patterns")
ws2.set_zoom(90)

ws2.merge_range("A1:D1", f"🔎  Cross-Gallery Patterns — {DATE}", title_fmt)
ws2.merge_range("A2:D2",
    "Patterns detected across multiple galleries this week | Frameworks: Marketing Psychology + Content Strategy",
    subtitle_fmt)
ws2.set_row(0, 22)
ws2.set_row(1, 16)

HEADERS_2 = ["Pattern Observed", "Galleries Involved",
             "What It May Signal", "Recommended Action for Duque Arango"]
COL_WIDTHS_2 = [50, 25, 55, 55]
for col, (h, w) in enumerate(zip(HEADERS_2, COL_WIDTHS_2)):
    ws2.write(2, col, h, pattern_header_fmt)
    ws2.set_column(col, col, w)
ws2.set_row(2, 32)

for row_idx, p in enumerate(PATTERNS):
    r = row_idx + 3
    values = [p["pattern"], p["galleries"], p["signal"], p["action"]]
    for col, val in enumerate(values):
        ws2.write(r, col, val, pattern_row_fmt)
    ws2.set_row(r, 110)

ws2.freeze_panes(3, 0)

workbook.close()

# ── Validate ───────────────────────────────────────────────────────────────
size = os.path.getsize(OUTPUT_PATH)
print(f"File: {OUTPUT_PATH}")
print(f"Size: {size} bytes")
assert size > 0, "ERROR: File is empty!"
print("✓ Validation passed — file is ready.")
