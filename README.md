# DM B2 — Visualisation interactive de l'aromaticité (Diels-Alder exo)

Évolution temporelle des descripteurs d'aromaticité (HOMA + LDM, suite **AroX**) le long d'une trajectoire de dynamique moléculaire NVE partant du TS, pour la réaction Diels-Alder **exo** entre nitrobenzène et 1,3-butadiène.

## Visualisation en ligne

**[Ouvrir la figure interactive](https://ariax1973.github.io/DM-B2-aromaticity-viz/)** — page servie par GitHub Pages, ouvre directement le HTML standalone (Plotly.js + Tailwind via CDN).

## Contenu

| Fichier | Description |
|---|---|
| [`viz_DM_B2.html`](viz_DM_B2.html) | Visualisation standalone (~89 KB) — ouvrable hors ligne après premier chargement des CDN |
| [`index.html`](index.html) | Copie identique servie à la racine GitHub Pages |
| [`build_html_viz.py`](build_html_viz.py) | Script Python qui régénère le HTML à partir du tableau Excel |
| [`tableau_filled_v2.xlsx`](tableau_filled_v2.xlsx) | Données sources — 132 lignes (44 frames MD × 3 motifs) |

## Pipeline calcul (déjà fait, source des données)

- **Dynamique moléculaire** : B3LYP/6-31G(d), ensemble NVE, T = 50 K, depuis le TS
- **Pas MD** : dt = 0.2 fs, snapshot tous les 10 pas (2 fs entre frames)
- **Échantillonnage** : 44 frames (22 reverse, t < 0 ; 22 forward, t > 0), denses près du TS
- **Single-point électronique** : wB97XD/6-311++G(d,p)
- **Analyse aromaticité** : AroX v0.2.0 (HOMA + LDM)

## Trois motifs analysés

- **A (nitrobenzène)** — atomes [1,2,3,4,5,6], cycle aromatique stable
- **B (cycle en formation)** — atomes [4,5,15,16,17,18], cycle Diels-Alder qui se ferme
- **A+B (cycles fusionnés)** — atomes [1,2,3,4,5,6,15,16,17,18], macrocycle 10 atomes

## Régénérer le HTML

```bash
pip install openpyxl
python build_html_viz.py
```

## Fonctionnalités de la visualisation

- Panneaux empilés partageant l'axe X (temps fs), TS à t = 0
- 15 descripteurs disponibles (HOMA, EN, GEO, LDM Frob/Off/Diag/RMSD/CT%, Q Frob/RMSD/G, S-hom/G, etc.)
- Toggle motifs par chips, sélection libre des descripteurs (un panneau chacun)
- Zones colorées reverse / TS / forward, ligne TS pointillée
- Tooltip riche : motif, valeur, temps, step MD, direction
- Tableau filtrable + export CSV
- **Export figure publication** : PNG ×3 (300 dpi), SVG vectoriel, ou modale configurable (largeur mm/inch, format, dpi)
- Typographie serif (STIX Two), palette daltoniens Okabe-Ito, tailles ≥ 14 pt

---

*ARIA NOROOZI — Stage LCT 2026*
