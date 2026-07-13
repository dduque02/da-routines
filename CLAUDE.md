# Galería Duque Arango — Routines Context

## About
You are running automated intelligence routines for Galería Duque Arango,
a Latin American modern and contemporary art gallery with spaces in
Medellín, Bogotá, Miami, and New York.

Primary domain: galeriaduquearango.com
Main markets: Colombia and United States

## Represented Artists (full roster)

### Modern Masters (secondary market focus)
- Fernando Botero
- Alejandro Obregón
- Enrique Grau
- Ana Mercedes Hoyos
- Luis Caballero
- Omar Rayo
- Edgar Negret
- Rufino Tamayo
- Wifredo Lam
- Oswaldo Guayasamín
- Leonora Carrington
- Carlos Cruz-Diez
- Julio Le Parc
- Fernando de Szyszlo
- Darío Morales
- Manuel Hernández
- Igor Mitoraj
- Jaume Plensa
- Manolo Valdés
- Ugo Rondinone

### Contemporary Artists (primary market focus)
- Olga de Amaral
- David Manzur
- Julio Larraz
- Ariel Cabrera
- Javier Caraballo
- Gustavo Vélez
- Sair García
- Reynier Ferrer
- Óscar Murillo
- Darío Ortiz
- Alejandra Aristizábal
- Carlos Vega
- Andrés Moreno
- Manuela Echeverri
- Nadín Ospina
- Guillermo Muñoz Vera
- Tomás Ochoa
- Carlos Salas
- Pepe Toledo
- Álvaro Barrios

## Available Skills
Load and apply these skills based on the task at hand:

### For SEO analysis and recommendations:
- skills/seo-audit.md — technical SEO framework and audit criteria
- skills/ai-seo.md — AI search visibility and citation optimization
- skills/content-strategy.md — content planning from keyword and
  competitor data
- skills/competitor-alternatives.md — competitive positioning analysis

### For interpreting competitor behavior:
- skills/marketing-psychology.md — identify psychological and strategic
  patterns in competitor activity
- skills/customer-research.md — extract signal from competitor reviews,
  community presence, and digital footprint

## Skill Activation Rules
- SEO Routine (Thursdays): activate seo-audit + ai-seo +
  content-strategy + competitor-alternatives
- Competitive Intel Routine (Mondays): activate marketing-psychology +
  customer-research + content-strategy
- For recommendations in any routine: apply marketing-psychology
  quick reference table to prioritize actions

## Gallery Context for SEO
Target keywords fall into these categories:
1. Artist names — every artist in the roster above is a keyword target
2. Gallery brand (galeria duque arango, duque arango)
3. Editorial content (artistas colombianos, arte latinoamericano,
   arte moderno colombiano, escultura colombiana)
4. Commercial intent (galeria arte medellin, comprar arte colombiano,
   obras de arte colombia, coleccionar arte latinoamericano)

Content gap opportunities must be evaluated against the full artist
roster. Priority goes to artists with high search volume where the
gallery already ranks (Edgar Negret, David Manzur, Alejandro Obregón)
and to contemporary artists the gallery represents exclusively
(Javier Caraballo, Alejandra Aristizábal, Gustavo Vélez, Manuela
Echeverri, Andrés Moreno).

When a competitor is found ranking for any artist name in the roster,
flag it as a high-priority threat.

## Weekly Intelligence Routine — Operational Notes
These notes override the routine prompt's defaults wherever they
conflict. Added 2026-07-13 after the total network-block incident.

### Email delivery (draft-only, by design)
- The Gmail connector in this environment is intentionally DRAFT-ONLY.
  Do NOT treat the absence of a send tool as a failure and do not try
  to work around it: create a Gmail draft of the weekly report and the
  team sends it manually.
- Recipients for the draft: davidduque@galeriaduquearango.com,
  santiagoduque@galeriaduquearango.com, digital@galeriaduquearango.com,
  desarrollo@galeriaduquearango.com (desarrollo@ matches historical
  sends even though the routine prompt omits it).
- Keep the exact subject prefix "🖼️ Weekly Gallery Intelligence — "
  followed by the date: the sent email is next week's memory.
- Drafts do not support attachments. Deliver the Excel by (a)
  committing it to reports/ in this repo and (b) sending it to the
  user via the file-delivery tool; the draft body must state where
  the Excel lives.

### Site access and 403 handling
- If a gallery site returns HTTP 403, first check whether it is an
  egress-policy block: curl -sS "$HTTPS_PROXY/__agentproxy/status"
  and look for the domain under recentRelayFailures with
  "connect_rejected".
- A proxy policy block means the gallery is "NO VERIFICABLE (bloqueo
  de red)" — NEVER report it as "No activity detected this week."
  False "sin actividad" entries poison the following weeks' memory.
- Findings recovered only through web search (no direct page load)
  must be labeled lower-confidence in the Strategic Note.
- Before escalating any multi-week threat that has only been seen via
  search results, re-confirm it against the primary site. Lesson
  learned: Ascaso's "Reviver" show was stale 2022 indexed content that
  was reported as an active threat for several weeks.
- If the block persists, flag it prominently in the email draft and in
  the notification so the environment's network policy gets fixed.

### Network policy allowlist (for the environment settings)
The environment's network policy must allow these domains for the
routine to work. If runs report proxy 403s, re-add them (with and
without www) in the environment configuration at claude.ai/code:
galerialacometa.com, galeriacasacuadrada.com, casasriegner.com,
galeriaelmuseo.com, latinartcore.com, artoftheworldgallery.com,
ascasogallery.com, galeriafreites.com, operagallery.com,
gagosian.com, davidzwirner.com, hauserwirth.com, galerie-lelong.com,
lissongallery.com, lehmannmaupin.com, perrotin.com

### Dates
- If the routine prompt arrives with unsubstituted placeholders like
  {{DATE}} / {{DATE_MINUS_7}}, use today's date from the session
  context; the reporting window is always the last 7 days.

### Files
- Save weekly Excel reports under reports/
  (reports/Gallery_Intelligence_YYYY_MM_DD.xlsx), then commit and
  push them to the designated branch.
