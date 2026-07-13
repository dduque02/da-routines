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

## Operational Notes — SEO Weekly Routine (lessons from past runs)

### Gmail: draft-only is INTENTIONAL
The Gmail connector only creates drafts and cannot attach files.
This is by design: David reviews every report and attaches the Excel
manually before sending. Do NOT treat draft-only as an error, do NOT
add alarmist "operational failure" notes to the email body, and do NOT
retry sending. Workflow: create ONE plain-text draft (never escaped
HTML entities — they render literally), include the GitHub link to the
Excel, then notify David that the draft is ready for review.

### Repository record keeping: open a PR
Session branches (claude/*) are never merged automatically, so an
Excel committed only to a session branch is effectively lost (this
happened with the 2026-06-19 report). After committing the weekly
Excel to reports/, ALWAYS open a pull request against main so the
file reaches the permanent history. Mention the PR link in the email
draft instead of a bare branch/commit reference.

### Semrush API unit budget
This routine is expensive. Known costs: domain_organic ~10 units/line,
domain_domains (keyword gap) ~80 units/line, backlinks_overview ~40
units/request. Rules learned:
- Authority Score (0-100) is NOT available via domain_rank
  export_columns; it only comes from backlinks_overview
  (params: target, target_type=root_domain). Budget 17 calls for it.
- Never use display_limit above ~50 unless strictly required; a
  display_limit=4000 query exhausted all remaining units on 2026-07-10.
- The "keywords en top 10" email metric: if needed, query it FIRST in
  Block 1 while units are fresh, with export_columns=["Ph"] and
  display_limit=1000 max (report "1.000+" if truncated). If units are
  short, omit the metric with a one-line note — never at the expense
  of the gap analysis or benchmark blocks.
- If any call returns the "not enough API units" message, stop making
  Semrush calls, complete the report with data already collected, and
  note the gap; more units at semrush.com/mcp-access.

### Position tracking
No Semrush project/tracking campaign exists (list_projects returns
empty), and projects cannot be created via this API. Skip the
"keywords improved/declined this week" sub-section silently until a
Position Tracking project is created in the Semrush UI.

### Email recipients
The routine spec lists 3 recipients (davidduque, santiagoduque,
digital @galeriaduquearango.com), but reports sent before July 2026
also included desarrollo@galeriaduquearango.com. Use the spec's 3;
David adds others during his manual review if needed.

### Memory continuity
Weekly reports are the routine's memory. If a week is missing from
the Gmail history (e.g., no report exists for 2026-06-26), say so
explicitly in the evolution section rather than interpolating.
