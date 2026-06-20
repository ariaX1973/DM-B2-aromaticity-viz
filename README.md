# DM B2 — Visualisation interactive multi-températures (Diels-Alder exo)

Évolution temporelle des descripteurs d'aromaticité (HOMA + LDM, suite **AroX**) le long de trajectoires de dynamique moléculaire NVE partant du TS, **à plusieurs températures**, pour la réaction Diels-Alder **exo** entre nitrobenzène et 1,3-butadiène.

## Visualisation en ligne

**[Ouvrir la figure interactive](https://ariax1973.github.io/DM-B2-aromaticity-viz/)** — page servie par GitHub Pages, ouvre directement le HTML standalone (Plotly.js + Tailwind via CDN).

## Contenu

| Fichier | Description |
|---|---|
| [`viz_DM_B2.html`](viz_DM_B2.html) | Visualisation standalone (~143 KB) — ouvrable hors ligne après premier chargement des CDN |
| [`index.html`](index.html) | Copie identique servie à la racine GitHub Pages |
| [`build_html_viz.py`](build_html_viz.py) | Script Python qui régénère le HTML à partir des tableaux Excel |
| [`tableau_filled_T5K.xlsx`](tableau_filled_T5K.xlsx) | Données T = 5 K, 108 lignes (36 frames × 3 motifs) |
| [`tableau_filled_v2_T50K.xlsx`](tableau_filled_v2_T50K.xlsx) | Données T = 50 K, 132 lignes (44 frames × 3 motifs) |

## Pipeline calcul (source des données)

- **Dynamique moléculaire** : B3LYP/6-31G(d), ensemble NVE, depuis le TS
- **Températures** : T = 5 K et T = 50 K (T = 15 K et T = 35 K prévues — extensible par simple ajout d'un xlsx)
- **Pas MD** : dt = 0.2 fs, snapshot tous les 10 pas (2 fs entre frames)
- **Échantillonnage** :
  - T5K : 36 frames sur ±118 fs (denses près du TS)
  - T50K : 44 frames sur ±198 fs (denses près du TS)
- **Single-point électronique** : ωB97X-D/6-311++G(d,p)
- **Analyse aromaticité** : AroX v0.2.0 (HOMA + LDM)

## Trois motifs analysés

- **A (nitrobenzène)** — atomes [1,2,3,4,5,6], cycle aromatique stable
- **B (cycle en formation)** — atomes [4,5,15,16,17,18], cycle Diels-Alder qui se ferme
- **A+B (cycles fusionnés)** — atomes [1,2,3,4,5,6,15,16,17,18], macrocycle 10 atomes

## Encodage visuel

- **Couleur = motif** (palette Okabe-Ito daltoniens)
  - A → bleu `#0072B2` · B → vermillon `#D55E00` · A+B → vert `#009E73`
- **Style de ligne + marker = température**
  - T5K : solid + circle · T50K : dash + square · T15K : dot + diamond · T35K : dashdot + triangle
- **Courbe moyenne** (toggle) : ligne épaisse longdashdot + marker star, calculée par motif sur les T actives aux points temporels exactement communs. Bande ±σ optionnelle.

## Régénérer le HTML

```bash
pip install openpyxl
python build_html_viz.py
```

Pour ajouter une nouvelle température, ajouter une entrée à la liste `TEMPERATURES` au sommet de `build_html_viz.py` :

```python
TEMPERATURES = [
    ("T5K",  "tableau_filled_T5K.xlsx",     "5"),
    ("T50K", "tableau_filled_v2_T50K.xlsx", "50"),
    ("T15K", "tableau_filled_T15K.xlsx",   "15"),   # <-- nouveau
    ("T35K", "tableau_filled_T35K.xlsx",   "35"),   # <-- nouveau
]
```

## Fonctionnalités de la visualisation

- Panneaux empilés partageant l'axe X (temps fs), TS à t = 0
- 15 descripteurs disponibles (HOMA, EN, GEO, LDM Frob/Off/Diag/RMSD/CT%, Q Frob/RMSD/G, S-hom/G, etc.)
- Toggle motifs et températures par chips indépendants
- **Courbe moyenne** par motif sur les T actives + bande ±σ
- Zones colorées reverse / TS / forward, ligne TS pointillée, dédup automatique du TS forward
- Tooltip riche : T, motif, valeur, temps, step MD, direction
- Tableau filtrable (colonne T) + export CSV
- **Export figure publication** : PNG ×3 (300 dpi), SVG vectoriel, ou modale configurable (largeur mm/inch, format, dpi)
- Typographie serif (STIX Two), palette daltoniens Okabe-Ito, tailles ≥ 14 pt

---

*ARIA NOROOZI — Stage LCT 2026*
