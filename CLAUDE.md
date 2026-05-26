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

## Weekly Iteration & Persistence

Each routine produces two artifacts every week, stored in the `reports/`
directory and committed to the repo:

1. **Presentation Excel** — `reports/<routine>_<YYYY-MM-DD>.xlsx`
   Human-readable workbook for the team (the same multi-tab format
   currently used for `Inteligencia_SEO_*.xlsx`).
2. **Structured data sidecar** — `reports/<routine>_<YYYY-MM-DD>.json`
   Machine-readable snapshot of the raw metrics. This is what enables
   week-over-week comparison.

`<routine>` is one of: `seo`, `competitive_intel`.

### Required JSON schema — SEO routine

```json
{
  "fecha": "YYYY-MM-DD",
  "rutina": "seo",
  "propio": {
    "co": { "trafico": 0, "keywords": 0, "top10": 0, "costo_usd": 0 },
    "us": { "trafico": 0, "keywords": 0, "top10": 0, "costo_usd": 0 },
    "top_keywords_co": [ { "kw": "", "pos": 0, "vol": 0, "url": "" } ],
    "top_keywords_us": [ { "kw": "", "pos": 0, "vol": 0, "url": "" } ],
    "top_paginas_co": [ { "url": "", "trafico": 0, "keywords": 0 } ],
    "top_paginas_us": [ { "url": "", "trafico": 0, "keywords": 0 } ]
  },
  "competidores": [
    { "dominio": "", "trafico_co": 0, "keywords_co": 0,
      "trafico_us": 0, "keywords_us": 0 }
  ],
  "brechas": [
    { "kw": "", "vol": 0, "categoria": "", "competidor": "",
      "pos_competidor": 0, "nivel": "Alta|Media|Baja" }
  ],
  "paginas_top_competencia": [
    { "competidor": "", "mercado": "CO|US", "url": "",
      "trafico": 0, "keyword_principal": "", "tipo": "" }
  ],
  "insights": [
    { "titulo": "", "prioridad": "Alta|Media|Baja" }
  ]
}
```

### Required JSON schema — Competitive Intel routine

```json
{
  "fecha": "YYYY-MM-DD",
  "rutina": "competitive_intel",
  "competidores": [
    {
      "nombre": "",
      "actividad_digital": { "posts_nuevos": 0, "exposiciones_nuevas": 0 },
      "menciones_artistas_roster": [],
      "movimientos_estrategicos": [],
      "patron_psicologico": ""
    }
  ],
  "alertas": [
    { "titulo": "", "prioridad": "Alta|Media|Baja", "competidor": "" }
  ]
}
```

### Iteration Rules (mandatory from week 2 onwards)

Before generating any weekly report:

1. **Read the most recent prior snapshot** from `reports/<routine>_*.json`
   (highest date that is not today's). If none exists, it is the first run
   — skip comparison and note this in the Excel.
2. **Compute deltas vs. prior week** for every metric in the JSON.
3. **Add a "Comparativo Semanal" tab to the Excel** as the second tab
   (after the cover/summary), with these sections:
   - Δ Tráfico orgánico CO/US (absoluto y %)
   - Keywords ganadas (entraron al top 20)
   - Keywords perdidas (salieron del top 20 o cayeron >5 posiciones)
   - Brechas nuevas detectadas esta semana
   - Brechas cerradas (donde la galería ahora rankea)
   - Alertas nuevas vs. alertas resueltas
   - Movimientos de competidores (tráfico ±10% requiere mención)
4. **Highlight in the email draft** the 3 deltas más relevantes de la
   semana antes de las oportunidades de contenido.

After generating both artifacts:

5. **Write both files** to `reports/` with the correct date in the filename.
6. **Commit and push** to the current working branch so the next session
   can find them.
