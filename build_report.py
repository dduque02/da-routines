#!/usr/bin/env python3
"""Build Gallery_Intelligence_2026-06-08.xlsx"""

import xlsxwriter
import os

FILENAME = "Gallery_Intelligence_2026-06-08.xlsx"
FILEPATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), FILENAME)

workbook = xlsxwriter.Workbook(FILEPATH)

# ── Formats ──────────────────────────────────────────────────────────────────
hdr_fmt = workbook.add_format({
    "bold": True, "bg_color": "#1a1a2e", "font_color": "#FFFFFF",
    "border": 1, "text_wrap": True, "valign": "vcenter", "align": "center"
})
title_fmt = workbook.add_format({
    "bold": True, "font_size": 14, "font_color": "#1a1a2e"
})
red_fmt = workbook.add_format({
    "bg_color": "#FFCCCC", "border": 1, "text_wrap": True, "valign": "top"
})
orange_fmt = workbook.add_format({
    "bg_color": "#FFE4B5", "border": 1, "text_wrap": True, "valign": "top"
})
normal_fmt = workbook.add_format({
    "border": 1, "text_wrap": True, "valign": "top"
})
gray_fmt = workbook.add_format({
    "bg_color": "#F5F5F5", "border": 1, "text_wrap": True,
    "valign": "top", "italic": True, "font_color": "#888888"
})
url_fmt = workbook.add_format({
    "border": 1, "text_wrap": True, "valign": "top",
    "font_color": "#0000EE", "underline": True
})

# ═══════════════════════════════════════════════════════════════════════════
# TAB 1 — Changes & Findings
# ═══════════════════════════════════════════════════════════════════════════
ws1 = workbook.add_worksheet("Changes & Findings")
ws1.set_zoom(85)
ws1.set_default_row(55)
ws1.freeze_panes(2, 0)

# Column widths
ws1.set_column("A:A", 22)   # Gallery
ws1.set_column("B:B", 8)    # Tier
ws1.set_column("C:C", 14)   # Section
ws1.set_column("D:D", 20)   # Type of Change
ws1.set_column("E:E", 50)   # Description
ws1.set_column("F:F", 42)   # Strategic Note
ws1.set_column("G:G", 45)   # URL

# Title row
ws1.set_row(0, 18)
ws1.merge_range("A1:G1",
    "Gallery Intelligence Report — Week of 2026-06-08 | Galería Duque Arango",
    title_fmt)

# Header row
ws1.set_row(1, 32)
headers = ["Gallery", "Tier", "Section", "Type of Change",
           "Description", "Strategic Note for Duque Arango", "URL"]
for col, h in enumerate(headers):
    ws1.write(1, col, h, hdr_fmt)

# ── Data rows ────────────────────────────────────────────────────────────────
rows = [
    # (gallery, tier, section, type, description, strategic_note, url, row_fmt)

    # ── CRITICAL: Le Parc death ──────────────────────────────────────────────
    (
        "Perrotin", "3", "News / Obituary", "Other",
        "Julio Le Parc (1928–2026) died in Paris on May 30, 2026 at age 97. "
        "Confirmed by family and Argentine officials. Coverage exploded across "
        "all major art outlets this week: Art News (June 1), The Art Newspaper "
        "(June 5), Hyperallergic, Ocula, Artsy, LACMA Unframed (June 4). "
        "Perrotin is his primary dealer. His Tate Modern career retrospective "
        "opens June 11 as a posthumous tribute.",
        "URGENT. Julio Le Parc is on Duque Arango's roster. His death is the "
        "dominant art-world event this week. Secondary market prices for his "
        "works will be closely watched — prepare inventory position and "
        "consider publishing a tribute statement to reinforce the gallery's "
        "role as a custodian of his legacy. The Tate retrospective (June 11) "
        "will sustain media coverage for months.",
        "https://www.theartnewspaper.com/2026/06/05/julio-le-parc-kinetic-art-pioneer-obituary",
        red_fmt
    ),
    (
        "Perrotin", "3", "Exhibitions", "Institutional",
        "Tate Modern career retrospective 'Julio Le Parc' opens June 11, 2026 "
        "(runs to May 3, 2027). More than 60 works including light "
        "installations, kinetic sculptures, and geometric paintings. Curators "
        "worked directly with the artist before his passing; the show now "
        "serves as a posthumous tribute.",
        "The Tate retrospective will dominate art-press coverage through at "
        "least July. Duque Arango should align any Le Parc content publishing "
        "with the June 11 opening. Institutional validation at this scale "
        "typically drives secondary market price appreciation significantly.",
        "https://www.tate.org.uk/whats-on/tate-modern/julio-le-parc",
        red_fmt
    ),

    # ── LISSON / Olga de Amaral ──────────────────────────────────────────────
    (
        "Lisson Gallery", "3", "Fairs / Events", "Fair Announcement",
        "Lisson Gallery confirms Olga de Amaral's works will be featured at "
        "Art Basel Basel 2026 (VIP preview June 16–17, public June 18–21) "
        "alongside Ryan Gander, Wael Shawky, and others. A major textile work "
        "'Rojo y oro' (2016) is highlighted. Additionally, Lisson presents a "
        "solo booth of de Amaral at Art Basel Qatar 2026, and her work appears "
        "at Art Basel Hong Kong 2026.",
        "PRIORITY. Olga de Amaral is on Duque Arango's roster. Her "
        "simultaneous presence at three Art Basel editions represents peak "
        "institutional visibility. This is the perfect moment for Duque Arango "
        "to publish editorial content around her work, activate collector "
        "outreach, and reinforce the gallery's association with her at the "
        "height of international attention.",
        "https://www.lissongallery.com/news/lisson-at-art-basel-2026",
        orange_fmt
    ),
    (
        "Lisson Gallery", "3", "Press / Editorial", "New Content",
        "Lisson Gallery published news coverage of the Akris Fall 2026 fashion "
        "collection collaboration with Olga de Amaral on their news page. The "
        "collection, shown at Paris Fashion Week at the Palais de Tokyo, was "
        "directly inspired by de Amaral's textile practice. Albert Kriemler "
        "visited her Bogotá studio before the collaboration.",
        "Fashion x fine art collaborations at this profile (Akris is a "
        "prestigious Swiss luxury brand) dramatically expand an artist's "
        "audience beyond traditional gallery circles. Duque Arango should "
        "leverage the fashion press coverage to introduce de Amaral's work to "
        "design-adjacent collectors in Colombia and internationally.",
        "https://www.lissongallery.com/news/fall-2026-akris-collection-featuring-olga-de-amaral",
        orange_fmt
    ),

    # ── LATIN ART CORE ───────────────────────────────────────────────────────
    (
        "Latin Art Core", "2", "Exhibitions", "New Exhibition",
        "Solo exhibition by Cuban master Manuel Mendive announced/opening "
        "June 19, 2026 at Latin Art Core, Miami. Works span 2023–2025 plus "
        "significant earlier pieces. Opening reception June 19, 7–10 pm. "
        "Theme: 'With the new day the sun shines and leads us.'",
        "Mendive is a significant Cuban master with international standing. "
        "Latin Art Core is actively programming high-caliber Cuban modernist "
        "work in Miami — the same audience Duque Arango targets. Monitor "
        "collector response to inform Duque Arango's own Cuban/Caribbean "
        "programming strategy.",
        "https://latinartcore.com/exhibition/manuel-mendive-with-the-new-day-the-sun-shines-and-leads-us/",
        normal_fmt
    ),

    # ── GALERÍA LA COMETA ────────────────────────────────────────────────────
    (
        "Galería La Cometa", "1", "News / Press", "New Content",
        "News post published announcing Alejandro Sánchez will represent "
        "Colombian art at 'Global FUTURE: Stills of Peace' (12th edition), "
        "Palazzo Acquaviva, Atri, Italy, July 5 – September 7, 2026. Sánchez "
        "presents critical work on neoliberalism in Latin America.",
        "La Cometa is actively securing international institutional placements "
        "for their artists. This Italy placement (even if a smaller venue) "
        "signals ongoing effort to build European presence for Colombian "
        "artists. Duque Arango should evaluate similar international placement "
        "opportunities for their contemporary artists.",
        "https://galerialacometa.com/noticias/alejandro-sanchez-global-future-stills-of-peace",
        normal_fmt
    ),

    # ── OPERA GALLERY ────────────────────────────────────────────────────────
    (
        "Opera Gallery", "2", "Exhibitions", "New Exhibition",
        "New group exhibition 'Pieter Obels | Feng Xiao-Min' opened June 4, "
        "2026 at Opera Gallery London (runs to July 5). Dutch sculptor Pieter "
        "Obels and French-Chinese artist Feng Xiao-Min — both new to Opera "
        "Gallery. Presented on the occasion of London Gallery Weekend. Also "
        "announced: 'The Monaco Masters Show: American 80s' (Warhol, Basquiat) "
        "at Opera Gallery Monaco, July 3 – October 3, 2026.",
        "Opera Gallery is actively programming European gallery weekend moments "
        "and pop-culture-adjacent content (Warhol/Basquiat Monaco show). This "
        "reflects an audience strategy targeting high-net-worth collectors at "
        "leisure destinations. Note Opera Gallery's absence of Latin American "
        "programming — a positioning gap Duque Arango could exploit.",
        "https://www.operagallery.com/event/pieter-obels-feng-xiao-min-contemporary-art-exhibition",
        normal_fmt
    ),

    # ── HAUSER & WIRTH (ongoing, Latin American filter) ──────────────────────
    (
        "Hauser & Wirth", "3", "Exhibitions", "New Exhibition",
        "Firelei Báez, 'feet squelching on wet grass, nourished by "
        "uncertainty' — H&W New York (22nd St), May 12 – July 31, 2026. "
        "First New York solo with the gallery. New paintings, works on paper, "
        "and bronze sculptures examining colonial histories and diasporic "
        "memory. Show generating sustained critical press through June.",
        "Báez is Dominican-American — Caribbean/Latin American diaspora. "
        "Hauser & Wirth's deep commitment to her work (full New York floor) "
        "signals growing institutional appetite for diasporic Latin American "
        "narratives. Duque Arango should track which Bogotá/Miami collectors "
        "are traveling to see this — they are Duque Arango's target audience.",
        "https://www.hauserwirth.com/hauser-wirth-exhibitions/firelei-baez/",
        normal_fmt
    ),

    # ── NO ACTIVITY ROWS ─────────────────────────────────────────────────────
    ("Galería Casa Cuadrada", "1", "—", "—",
     "No activity detected this week.", "—", "—", gray_fmt),

    ("Casas Riegner", "1", "—", "—",
     "No activity detected this week.", "—", "—", gray_fmt),

    ("Galería El Museo", "1", "—", "—",
     "No activity detected this week.", "—", "—", gray_fmt),

    ("Art of the World Gallery", "2", "—", "—",
     "No activity detected this week. Karla de Lara solo show closed June 6; "
     "no new programming announced.",
     "—", "—", gray_fmt),

    ("Ascaso Gallery", "2", "—", "—",
     "No activity detected this week.", "—", "—", gray_fmt),

    ("Galería Freites", "2", "—", "—",
     "No activity detected this week. Site returned limited indexed content.",
     "—", "—", gray_fmt),

    ("Gagosian", "3", "—", "—",
     "No Latin American artist activity detected this week.", "—", "—", gray_fmt),

    ("David Zwirner", "3", "—", "—",
     "No Latin American artist activity detected this week. "
     "Oscar Murillo (roster artist) has active shows at Mori Arts Center "
     "Tokyo (through June 8) and Kistefos Museum (through Oct), but no "
     "new Zwirner-organized activity this week.",
     "—", "—", gray_fmt),

    ("Galerie Lelong", "3", "—", "—",
     "No new Latin American artist activity detected this week. "
     "Current shows: Paula Rego, Eduardo Chillida, Kiki Smith "
     "(May 21 – July 11). Jaume Plensa solo ended March 2026.",
     "—", "—", gray_fmt),

    ("Lehmann Maupin", "3", "—", "—",
     "No Latin American artist activity detected this week. "
     "OSGEMEOS (Brazilian) show 'The Open Window' closed June 6; "
     "no new Latin American programming announced.",
     "—", "—", gray_fmt),
]

for r_idx, row in enumerate(rows):
    row_num = r_idx + 2  # 0-indexed, rows 0=title, 1=header, 2+=data
    fmt = row[7]
    # Use URL format for last column if it's a real URL
    for col, val in enumerate(row[:7]):
        if col == 6 and val.startswith("http"):
            ws1.write_url(row_num, col, val, url_fmt, val)
        else:
            ws1.write(row_num, col, val, fmt)

# ═══════════════════════════════════════════════════════════════════════════
# TAB 2 — Cross-Gallery Patterns
# ═══════════════════════════════════════════════════════════════════════════
ws2 = workbook.add_worksheet("Cross-Gallery Patterns")
ws2.set_zoom(85)
ws2.set_default_row(70)
ws2.freeze_panes(2, 0)

ws2.set_column("A:A", 38)   # Pattern
ws2.set_column("B:B", 30)   # Galleries Involved
ws2.set_column("C:C", 38)   # What It May Signal
ws2.set_column("D:D", 48)   # Recommended Action

ws2.set_row(0, 18)
ws2.merge_range("A1:D1",
    "Cross-Gallery Patterns — Week of 2026-06-08 | Galería Duque Arango",
    title_fmt)

ws2.set_row(1, 32)
hdrs2 = ["Pattern Observed", "Galleries Involved",
         "What It May Signal", "Recommended Action for Duque Arango"]
for col, h in enumerate(hdrs2):
    ws2.write(1, col, h, hdr_fmt)

patterns = [
    (
        "Death of Julio Le Parc (May 30) dominates art world discourse — "
        "all major outlets publishing tributes and market analysis this week.",
        "Perrotin (primary dealer); Tate Modern; LACMA; Artsy; Art Newspaper; "
        "Hyperallergic; Ocula",
        "Major secondary market event. When a blue-chip artist of Le Parc's "
        "stature dies, prices for existing works typically appreciate "
        "significantly in the following 6–18 months. His Tate retrospective "
        "(June 11) will sustain media attention for at least a year.",
        "URGENT: Le Parc is on Duque Arango's roster. (1) Publish a tribute "
        "statement on website and social media. (2) Audit current Le Parc "
        "inventory and assess pricing strategy. (3) Prepare collector "
        "communications for inquiries. (4) Align content calendar with the "
        "Tate opening June 11 — maximum global press moment.",
        red_fmt
    ),
    (
        "Olga de Amaral reaches peak global visibility simultaneously: "
        "Art Basel Basel + Qatar + Hong Kong (June), Akris fashion collaboration, "
        "Soloviev Foundation show (through Dec 2026).",
        "Lisson Gallery; Art Basel; Akris (fashion); "
        "The Soloviev Foundation Gallery NY",
        "Rare convergence of art market, fashion, and institutional signals "
        "around a single artist. This 'visibility peak' typically produces a "
        "step-change in collector demand and press citations for 6–12 months. "
        "Lisson is executing a coordinated cross-platform strategy for de Amaral.",
        "Duque Arango should ride this wave: publish de Amaral editorial "
        "content during Art Basel week (June 16–21), activate collector "
        "outreach with context about the Akris collaboration and Art Basel "
        "presence, and consider whether any de Amaral works can be highlighted "
        "in the gallery's programming or social media during this peak window.",
        orange_fmt
    ),
    (
        "Art Basel Basel 2026 (June 18–21) announcement week — galleries "
        "publishing booth content and fare highlights ahead of VIP preview.",
        "Lisson Gallery; Art Basel (new galleries inc. Galería Guillermo de "
        "Osma with Torres-García). 290 galleries from 43 countries attending.",
        "The week before Art Basel's VIP opening is the highest-visibility "
        "week in the international art calendar. Galleries without Art Basel "
        "presence can still capture attention by publishing relevant content "
        "that contextualizes their artists within the fair's narrative.",
        "Even without a booth at Art Basel, Duque Arango can publish content "
        "this week that positions their artists in the Art Basel conversation: "
        "e.g., 'Our artists at Art Basel this year' (Olga de Amaral via "
        "Lisson), editorial on market trends visible at the fair, or collector "
        "guides to Latin American art at Art Basel.",
        orange_fmt
    ),
    (
        "Latin American artists gaining peak institutional momentum at "
        "international Tier 1 galleries simultaneously: Olga de Amaral "
        "(Lisson/Art Basel), Oscar Murillo (David Zwirner, international "
        "museum circuit), Firelei Báez (Hauser & Wirth NY debut).",
        "Lisson Gallery (de Amaral); David Zwirner (Murillo); "
        "Hauser & Wirth (Báez); Mori Art Museum Tokyo; Kistefos Museum",
        "Latin American contemporary art is at an institutional high-water "
        "mark. Multiple Tier 1 galleries are prioritizing Latin American "
        "artists for flagship programming. This signals growing collector "
        "appetite and confirms the market's long-term bullish stance on the "
        "region's art.",
        "Duque Arango should publish thought-leadership content positioning "
        "the gallery as the premier Colombian entry point into this global "
        "trend. Frame the gallery's roster in the context of international "
        "institutional validation. Pitch editorial content to press that "
        "covers Latin American contemporary art's global rise.",
        normal_fmt
    ),
    (
        "Colombian direct competitors (La Cometa, Casa Cuadrada, Casas "
        "Riegner, El Museo) show limited digital activity this week — "
        "quiet week for Tier 1.",
        "Galería La Cometa; Galería Casa Cuadrada; "
        "Casas Riegner; Galería El Museo",
        "Either a natural quiet period before summer, or opportunity "
        "for Duque Arango to dominate the Colombian art digital conversation "
        "this week when competitors are silent.",
        "Window to publish and distribute content without local competition "
        "for attention. This week's Le Parc news and Art Basel context provide "
        "ready-made editorial hooks. Push content on social and email now.",
        normal_fmt
    ),
]

for r_idx, row in enumerate(patterns):
    row_num = r_idx + 2
    fmt = row[4]
    for col, val in enumerate(row[:4]):
        ws2.write(row_num, col, val, fmt)

workbook.close()
print(f"File written: {FILEPATH}")
size = os.path.getsize(FILEPATH)
print(f"File size: {size} bytes")
assert size > 0, "ERROR: file is empty"
print("Validation PASSED")
