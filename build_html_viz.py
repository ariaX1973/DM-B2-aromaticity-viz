"""
Construit une visualisation HTML interactive standalone — qualité publication.

Source  : DM/B2/tableau_filled_v2.xlsx (feuille `endo`, contenu = exo, 132 lignes)
Sortie  : DM/B2/viz_DM_B2.html (ouvrable hors ligne, Plotly + Tailwind via CDN)

Caractéristiques publication :
  - Typographie serif (STIX Two Text/Math via CDN, fallback Times New Roman)
  - Labels math en italique avec indices (S_hom, Q_Frob, ...) via HTML inline
  - Palette daltoniens Okabe-Ito (bleu, vermillon, vert bleuté)
  - Tailles >= 14 pt axes / 12 pt ticks / 16 pt titres panneaux
  - Lignes 2.5 pt, marqueurs 7 px, grilles claires, marges généreuses
  - Boutons d'export PNG (scale configurable 3x/4x, 300 dpi), SVG vectoriel,
    et modal de configuration figure (largeur mm/inch, format, nom de fichier auto)
  - Aucun élément d'UI (chips, boutons, sélecteurs) n'apparaît dans l'export Plotly
"""

import json
from pathlib import Path
import openpyxl

HERE = Path(__file__).parent
XLSX = HERE / "tableau_filled_v2.xlsx"
OUT = HERE / "viz_DM_B2.html"


def load_data():
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    ws = wb["endo"]
    rows = []
    for r in range(2, ws.max_row + 1):
        rows.append({
            "molecule":    ws.cell(row=r, column=1).value,
            "motif":       ws.cell(row=r, column=2).value,
            "t":           float(ws.cell(row=r, column=3).value),
            "EN":          ws.cell(row=r, column=4).value,
            "GEO":         ws.cell(row=r, column=5).value,
            "HOMA":        ws.cell(row=r, column=6).value,
            "HOMA_TOTAL":  ws.cell(row=r, column=7).value,
            "LDM_OFF":     ws.cell(row=r, column=8).value,
            "LDM_DIAG":    ws.cell(row=r, column=9).value,
            "LDM_FROB":    ws.cell(row=r, column=10).value,
            "LDM_RMSD":    ws.cell(row=r, column=11).value,
            "LDM_CT_PCT":  ws.cell(row=r, column=12).value,
            "Q_FROB":      ws.cell(row=r, column=13).value,
            "Q_RMSD":      ws.cell(row=r, column=14).value,
            "Q_G":         ws.cell(row=r, column=15).value,
            "S_HOM":       ws.cell(row=r, column=16).value,
            "S_HOM_G":     ws.cell(row=r, column=17).value,
            "S_G":         ws.cell(row=r, column=18).value,
        })
    for r in rows:
        r["step"] = int(round(abs(r["t"]) / 0.2))
        r["direction"] = "TS" if r["t"] == 0 else ("reverse" if r["t"] < 0 else "forward")

    # Dédup à t=0 : le tableau contient 2 lignes par motif à t=0
    # (frame 0 du run reverse + frame 0 du run forward, SP regénérés
    # indépendamment → valeurs légèrement différentes). On ne garde
    # que la 1re occurrence (TS issu du calcul reverse) pour rattacher
    # le TS à la branche reverse uniquement.
    seen = set()
    deduped = []
    dups = 0
    for r in rows:
        key = (r["motif"], r["t"])
        if key in seen:
            dups += 1
            continue
        seen.add(key)
        deduped.append(r)
    if dups:
        print(f"  dédoublonnage : {dups} doublons (motif, t) supprimés "
              f"(typiquement TS forward, géométrie identique au TS reverse)")
    return deduped


def build_html(rows):
    data_json = json.dumps(rows, ensure_ascii=False)

    html = r"""<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<title>DM B2 exo — Évolution temporelle des descripteurs d'aromaticité</title>
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
<script src="https://cdn.tailwindcss.com"></script>
<!-- STIX Two : police serif scientifique standard (Computer Modern-like, support unicode math) -->
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=STIX+Two+Text:ital,wght@0,400;0,600;1,400;1,600&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
  :root {
    --serif: 'STIX Two Text', 'STIX2Text-Regular', 'Latin Modern Roman', 'Times New Roman', Times, serif;
    --ui:    'Inter', system-ui, -apple-system, 'Segoe UI', sans-serif;
  }
  body { font-family: var(--ui); }
  .panel { background: #ffffff; border-radius: 12px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }
  .chip { transition: all 0.15s ease; }
  .chip-active { box-shadow: 0 0 0 2px currentColor inset; }
  .tbl-row:nth-child(odd) { background: #f8fafc; }
  .desc-card { transition: all 0.1s ease; cursor: pointer; }
  .desc-card:hover { background: #eff6ff; }
  .desc-card-active { background: #dbeafe; border-color: #2563eb; }
  /* Pour les labels de descripteurs (chips et tableau) — police math */
  .math { font-family: var(--serif); font-style: italic; }
  /* Modale */
  .modal-bg { background: rgba(15, 23, 42, 0.55); }
  /* Force la police serif dans le rendu Plotly exporté */
  #plot text { font-family: var(--serif) !important; }
</style>
</head>
<body class="bg-slate-100 min-h-screen">
<div class="max-w-[1600px] mx-auto p-6 space-y-4">

  <!-- En-tête : métadonnées -->
  <header class="panel p-6">
    <h1 class="text-2xl font-bold text-slate-800" style="font-family:var(--serif);">
      Évolution des descripteurs d'aromaticité — Diels-Alder exo (B2)
    </h1>
    <div class="text-sm text-slate-600 mt-3 grid md:grid-cols-2 gap-x-8 gap-y-1.5">
      <div><span class="font-semibold text-slate-700">Système :</span>
           Nitrobenzène + 1,3-butadiène, voie exo</div>
      <div><span class="font-semibold text-slate-700">Échantillonnage :</span>
           44 frames × 3 motifs = <span class="font-mono">132 points</span></div>
      <div><span class="font-semibold text-slate-700">Dynamique moléculaire :</span>
           B3LYP/6-31G(d), NVE, T = 50 K, depuis le TS</div>
      <div><span class="font-semibold text-slate-700">Pas MD :</span>
           dt = 0.2 fs, snapshot tous les 10 pas (2 fs)</div>
      <div><span class="font-semibold text-slate-700">Single-point électronique :</span>
           wB97XD/6-311++G(d,p)</div>
      <div><span class="font-semibold text-slate-700">Analyse aromaticité :</span>
           AroX.v0.2.0 (HOMA + LDM)</div>
    </div>
    <div class="flex flex-wrap gap-3 mt-4 text-xs">
      <span class="px-2.5 py-1 rounded bg-sky-100 text-sky-800 border border-sky-200">
        Reverse : t &lt; 0 (vers réactifs)
      </span>
      <span class="px-2.5 py-1 rounded bg-amber-100 text-amber-800 border border-amber-200">
        TS : t = 0
      </span>
      <span class="px-2.5 py-1 rounded bg-rose-100 text-rose-800 border border-rose-200">
        Forward : t &gt; 0 (vers produits)
      </span>
    </div>
  </header>

  <!-- Contrôles : motifs + descripteurs + options + export -->
  <section class="panel p-5 space-y-4">
    <div>
      <div class="text-sm font-semibold text-slate-700 mb-2">
        Motifs (cliquer pour activer/désactiver)
      </div>
      <div id="motif-chips" class="flex gap-2 flex-wrap"></div>
    </div>

    <div>
      <div class="text-sm font-semibold text-slate-700 mb-2">
        Descripteurs à afficher (un panneau empilé par descripteur)
      </div>
      <div id="descriptor-grid" class="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-5 gap-2"></div>
    </div>

    <div class="flex items-center gap-4 pt-3 border-t border-slate-200 flex-wrap">
      <label class="flex items-center gap-2 text-sm">
        <input type="checkbox" id="show-markers" checked class="rounded">
        <span>Marqueurs visibles</span>
      </label>
      <label class="flex items-center gap-2 text-sm">
        <input type="checkbox" id="show-ts-line" checked class="rounded">
        <span>Ligne TS (t=0)</span>
      </label>
      <label class="flex items-center gap-2 text-sm">
        <input type="checkbox" id="show-zones" checked class="rounded">
        <span>Zones colorées rev/TS/fwd</span>
      </label>
      <label class="flex items-center gap-2 text-sm">
        <input type="checkbox" id="show-minor-grid" class="rounded">
        <span>Grille mineure</span>
      </label>

      <div class="ml-auto flex items-center gap-2">
        <button id="btn-export-png"
                class="text-sm px-3 py-1.5 rounded bg-blue-600 text-white hover:bg-blue-700 shadow-sm">
          Export PNG (×3, 300 dpi)
        </button>
        <button id="btn-export-svg"
                class="text-sm px-3 py-1.5 rounded bg-emerald-600 text-white hover:bg-emerald-700 shadow-sm">
          Export SVG vectoriel
        </button>
        <button id="btn-export-config"
                class="text-sm px-3 py-1.5 rounded bg-slate-700 text-white hover:bg-slate-800 shadow-sm">
          Export figure publication…
        </button>
        <button id="reset-btn"
                class="text-sm px-3 py-1.5 rounded bg-slate-200 hover:bg-slate-300">
          Réinitialiser
        </button>
      </div>
    </div>
  </section>

  <!-- Graphiques -->
  <section class="panel p-3">
    <div id="plot" style="width:100%; min-height:600px;"></div>
  </section>

  <!-- Tableau filtrable -->
  <section class="panel p-5">
    <div class="flex items-center gap-4 mb-3 flex-wrap">
      <h2 class="text-lg font-semibold text-slate-700">Données brutes</h2>
      <input type="search" id="tbl-search" placeholder="Filtre texte (motif, temps...)"
             class="border rounded px-3 py-1 text-sm flex-1 min-w-[200px]">
      <select id="tbl-motif-filter" class="border rounded px-2 py-1 text-sm">
        <option value="">Tous motifs</option>
      </select>
      <label class="text-sm flex items-center gap-2">
        <span>t min</span>
        <input type="number" id="tbl-tmin" value="-200" step="10" class="border rounded px-2 py-1 w-20 text-sm">
      </label>
      <label class="text-sm flex items-center gap-2">
        <span>t max</span>
        <input type="number" id="tbl-tmax" value="200" step="10" class="border rounded px-2 py-1 w-20 text-sm">
      </label>
      <button id="tbl-export"
              class="text-sm px-3 py-1 rounded bg-emerald-100 text-emerald-800 hover:bg-emerald-200">
        Export CSV
      </button>
    </div>
    <div class="overflow-auto max-h-[400px] border rounded">
      <table id="tbl" class="text-xs w-full">
        <thead class="bg-slate-200 sticky top-0">
          <tr id="tbl-head"></tr>
        </thead>
        <tbody id="tbl-body"></tbody>
      </table>
    </div>
    <div class="text-xs text-slate-500 mt-2">
      <span id="tbl-count"></span>
    </div>
  </section>

  <footer class="text-xs text-slate-400 text-center pb-6 pt-2"
          style="letter-spacing: 0.4em;">
    ARIA NOROOZI
  </footer>
</div>

<!-- Modale Export Figure -->
<div id="export-modal" class="fixed inset-0 modal-bg hidden items-center justify-center z-50">
  <div class="bg-white rounded-xl shadow-2xl p-6 w-[460px] max-w-[95vw]">
    <h3 class="text-lg font-semibold text-slate-800 mb-4">Export figure publication</h3>
    <div class="space-y-3 text-sm">
      <div class="grid grid-cols-3 gap-3 items-center">
        <label>Format</label>
        <select id="exp-format" class="col-span-2 border rounded px-2 py-1">
          <option value="png">PNG raster</option>
          <option value="svg">SVG vectoriel</option>
          <option value="jpeg">JPEG raster</option>
          <option value="webp">WebP raster</option>
        </select>
      </div>
      <div class="grid grid-cols-3 gap-3 items-center">
        <label>Unité</label>
        <select id="exp-unit" class="col-span-2 border rounded px-2 py-1">
          <option value="mm">mm</option>
          <option value="in">inch</option>
          <option value="px">pixel</option>
        </select>
      </div>
      <div class="grid grid-cols-3 gap-3 items-center">
        <label>Largeur</label>
        <input id="exp-width" type="number" value="180" step="1" class="col-span-2 border rounded px-2 py-1">
      </div>
      <div class="grid grid-cols-3 gap-3 items-center">
        <label>Hauteur (par panneau)</label>
        <input id="exp-height" type="number" value="70" step="1" class="col-span-2 border rounded px-2 py-1">
      </div>
      <div class="grid grid-cols-3 gap-3 items-center">
        <label>Résolution (dpi)</label>
        <input id="exp-dpi" type="number" value="300" step="50" class="col-span-2 border rounded px-2 py-1">
      </div>
      <div class="grid grid-cols-3 gap-3 items-center">
        <label>Nom de fichier</label>
        <input id="exp-name" type="text" class="col-span-2 border rounded px-2 py-1">
      </div>
      <div class="text-xs text-slate-500 pt-1">
        Astuce : pour un journal ACS / RSC, largeur 1 col = 84 mm, 2 col = 174 mm, hauteur ~ 70 mm/panneau.
      </div>
    </div>
    <div class="flex justify-end gap-2 mt-5">
      <button id="exp-cancel" class="px-3 py-1.5 rounded bg-slate-200 hover:bg-slate-300 text-sm">
        Annuler
      </button>
      <button id="exp-go" class="px-3 py-1.5 rounded bg-blue-600 text-white hover:bg-blue-700 text-sm">
        Télécharger
      </button>
    </div>
  </div>
</div>

<script>
// =================== DONNÉES ===================
const DATA = __DATA_PLACEHOLDER__;

// =================== CONFIG ===================
// Palette Okabe-Ito (adaptée daltoniens, recommandée Nature Methods 2011)
const MOTIF_COLORS = {
  "A (nitrobenzène)":       "#0072B2",  // bleu — cycle aromatique stable
  "B (cycle en formation)": "#D55E00",  // vermillon — cycle Diels-Alder en formation
  "A+B (cycles fusionnés)": "#009E73",  // vert bleuté — macrocycle fusionné
};

// Labels descripteurs : `label` = texte court (chip), `axis` = HTML pour titre axe Plotly
// (Plotly accepte <i>, <sub>, <sup>, <br> dans les titres), `legend` = nom complet + unité
const DESCRIPTORS = [
  { key: "HOMA",       label: "HOMA",       axis: "<i>HOMA</i>",
    legend: "HOMA (sans unité)",                     group: "aromaticité", default: true },
  { key: "S_HOM",      label: "S-hom",      axis: "<i>S</i><sub>hom</sub>",
    legend: "S-hom (sans unité)",                    group: "aromaticité", default: true },
  { key: "S_HOM_G",    label: "S-hom(G)",   axis: "<i>S</i><sub>hom</sub>(<i>G</i>)",
    legend: "S-hom(G) (sans unité)",                 group: "aromaticité", default: false },
  { key: "EN",         label: "EN",         axis: "<i>EN</i>",
    legend: "EN — composante électronique HOMA",     group: "HOMA comp.",  default: false },
  { key: "GEO",        label: "GEO",        axis: "<i>GEO</i>",
    legend: "GEO — composante géométrique HOMA",     group: "HOMA comp.",  default: false },
  { key: "HOMA_TOTAL", label: "HOMA total", axis: "<i>HOMA</i><sub>total</sub>",
    legend: "HOMA total (sans unité)",               group: "global",      default: false },
  { key: "LDM_FROB",   label: "LDM Frob",   axis: "‖<i>LDM</i>‖<sub>F</sub>",
    legend: "LDM ‖·‖ Frobenius (full matrix)",        group: "LDM brut",    default: true },
  { key: "LDM_OFF",    label: "LDM Off",    axis: "<i>LDM</i><sub>off-diag</sub>",
    legend: "LDM off-diagonal (sans unité)",         group: "LDM brut",    default: false },
  { key: "LDM_DIAG",   label: "LDM Diag",   axis: "<i>LDM</i><sub>diag</sub>",
    legend: "LDM diagonal (sans unité)",             group: "LDM brut",    default: false },
  { key: "LDM_RMSD",   label: "LDM RMSD",   axis: "<i>LDM</i><sub>RMSD</sub>",
    legend: "LDM RMSD (sans unité)",                 group: "LDM avancé",  default: false },
  { key: "LDM_CT_PCT", label: "LDM CT %",   axis: "<i>LDM</i><sub>CT</sub> (%)",
    legend: "LDM charge-transfer (%)",               group: "LDM avancé",  default: true },
  { key: "Q_FROB",     label: "Q Frob",     axis: "<i>Q</i><sub>Frob</sub>",
    legend: "Q Frobenius (sans unité)",              group: "LDM avancé",  default: false },
  { key: "Q_RMSD",     label: "Q RMSD",     axis: "<i>Q</i><sub>RMSD</sub>",
    legend: "Q RMSD (sans unité)",                   group: "LDM avancé",  default: false },
  { key: "Q_G",        label: "Q(G)",       axis: "<i>Q</i>(<i>G</i>)",
    legend: "Q(G) (sans unité)",                     group: "LDM avancé",  default: false },
  { key: "S_G",        label: "S(G)",       axis: "<i>S</i>(<i>G</i>)",
    legend: "S(G) (sans unité)",                     group: "LDM avancé",  default: false },
];

const SERIF = "STIX Two Text, STIX2Text-Regular, Latin Modern Roman, Times New Roman, Times, serif";

// =================== ÉTAT ===================
const state = {
  activeMotifs: new Set(Object.keys(MOTIF_COLORS)),
  activeDescriptors: new Set(DESCRIPTORS.filter(d => d.default).map(d => d.key)),
  showMarkers: true,
  showTSLine: true,
  showZones: true,
  showMinorGrid: false,
};

// =================== INIT UI ===================
function initMotifChips() {
  const cont = document.getElementById("motif-chips");
  cont.innerHTML = "";
  Object.entries(MOTIF_COLORS).forEach(([motif, color]) => {
    const el = document.createElement("button");
    el.className = "chip px-3 py-1.5 rounded-full text-sm font-medium border-2";
    el.style.color = color;
    el.style.borderColor = color;
    el.style.background = "white";
    el.textContent = motif;
    el.dataset.motif = motif;
    el.classList.add("chip-active");
    el.addEventListener("click", () => {
      if (state.activeMotifs.has(motif)) {
        state.activeMotifs.delete(motif);
        el.classList.remove("chip-active");
        el.style.opacity = "0.35";
      } else {
        state.activeMotifs.add(motif);
        el.classList.add("chip-active");
        el.style.opacity = "1";
      }
      render();
    });
    cont.appendChild(el);
  });
}

function initDescriptorGrid() {
  const cont = document.getElementById("descriptor-grid");
  cont.innerHTML = "";
  DESCRIPTORS.forEach(d => {
    const el = document.createElement("div");
    el.className = "desc-card border rounded p-2 text-center select-none";
    if (d.default) el.classList.add("desc-card-active");
    el.innerHTML = `<div class="math text-base">${d.axis}</div>
                    <div class="text-[10px] text-slate-500 mt-0.5">${d.group}</div>`;
    el.addEventListener("click", () => {
      if (state.activeDescriptors.has(d.key)) {
        state.activeDescriptors.delete(d.key);
        el.classList.remove("desc-card-active");
      } else {
        state.activeDescriptors.add(d.key);
        el.classList.add("desc-card-active");
      }
      render();
    });
    cont.appendChild(el);
  });
}

function initControls() {
  document.getElementById("show-markers").addEventListener("change", e => {
    state.showMarkers = e.target.checked; render();
  });
  document.getElementById("show-ts-line").addEventListener("change", e => {
    state.showTSLine = e.target.checked; render();
  });
  document.getElementById("show-zones").addEventListener("change", e => {
    state.showZones = e.target.checked; render();
  });
  document.getElementById("show-minor-grid").addEventListener("change", e => {
    state.showMinorGrid = e.target.checked; render();
  });
  document.getElementById("reset-btn").addEventListener("click", () => {
    state.activeMotifs = new Set(Object.keys(MOTIF_COLORS));
    state.activeDescriptors = new Set(DESCRIPTORS.filter(d => d.default).map(d => d.key));
    state.showMarkers = true; state.showTSLine = true; state.showZones = true;
    state.showMinorGrid = false;
    initMotifChips(); initDescriptorGrid();
    document.getElementById("show-markers").checked = true;
    document.getElementById("show-ts-line").checked = true;
    document.getElementById("show-zones").checked = true;
    document.getElementById("show-minor-grid").checked = false;
    render();
  });

  // Export PNG rapide
  document.getElementById("btn-export-png").addEventListener("click", () => {
    const n = Math.max(1, state.activeDescriptors.size);
    Plotly.downloadImage("plot", {
      format: "png",
      filename: autoFilename("png"),
      width: 1600,
      height: 320 * n + 120,
      scale: 3,
    });
  });
  // Export SVG rapide
  document.getElementById("btn-export-svg").addEventListener("click", () => {
    const n = Math.max(1, state.activeDescriptors.size);
    Plotly.downloadImage("plot", {
      format: "svg",
      filename: autoFilename("svg"),
      width: 1600,
      height: 320 * n + 120,
    });
  });
  // Modale Export configurable
  const modal = document.getElementById("export-modal");
  document.getElementById("btn-export-config").addEventListener("click", () => {
    document.getElementById("exp-name").value = autoFilename("");
    modal.classList.remove("hidden");
    modal.classList.add("flex");
  });
  document.getElementById("exp-cancel").addEventListener("click", () => {
    modal.classList.add("hidden");
    modal.classList.remove("flex");
  });
  document.getElementById("exp-go").addEventListener("click", () => {
    const fmt   = document.getElementById("exp-format").value;
    const unit  = document.getElementById("exp-unit").value;
    const wVal  = parseFloat(document.getElementById("exp-width").value);
    const hVal  = parseFloat(document.getElementById("exp-height").value);
    const dpi   = parseFloat(document.getElementById("exp-dpi").value);
    let   name  = (document.getElementById("exp-name").value || "figure").trim();
    if (!name.toLowerCase().endsWith("." + fmt)) name += "." + fmt;
    name = name.slice(0, -(fmt.length + 1));

    const n = Math.max(1, state.activeDescriptors.size);
    const pxPerUnit = unit === "mm" ? dpi / 25.4
                    : unit === "in" ? dpi
                    : 1;  // px
    const totalW = Math.round(wVal * pxPerUnit);
    const totalH = Math.round(hVal * n * pxPerUnit);
    const scale  = unit === "px" ? Math.max(1, dpi / 96) : 1;

    Plotly.downloadImage("plot", {
      format: fmt,
      filename: name,
      width:  totalW,
      height: totalH,
      scale:  scale,
    });
    modal.classList.add("hidden");
    modal.classList.remove("flex");
  });
}

function autoFilename(ext) {
  const date = new Date().toISOString().slice(0, 10);
  const descs = [...state.activeDescriptors].slice(0, 4).join("_");
  const base = `DM_B2_exo_${date}_${descs || "vide"}`.replace(/[^A-Za-z0-9_\-]/g, "_");
  return ext ? base + "." + ext : base;
}

// =================== TRACÉ ===================
function groupByMotif(rows) {
  const out = {};
  rows.forEach(r => { (out[r.motif] = out[r.motif] || []).push(r); });
  Object.values(out).forEach(arr => arr.sort((a, b) => a.t - b.t));
  return out;
}

function render() {
  const activeDesc = DESCRIPTORS.filter(d => state.activeDescriptors.has(d.key));
  if (activeDesc.length === 0) {
    Plotly.purge("plot");
    document.getElementById("plot").innerHTML =
      '<div class="text-center py-16 text-slate-400" style="font-family:var(--serif);font-style:italic;">'
      + 'Sélectionne au moins un descripteur ci-dessus</div>';
    renderTable();
    return;
  }

  const grouped = groupByMotif(DATA);
  const tMin = Math.min(...DATA.map(r => r.t));
  const tMax = Math.max(...DATA.map(r => r.t));

  // Une trace continue par (descripteur × motif) : reverse → TS (unique, du run
  // reverse, gardé par dédup côté Python) → forward, sans rupture visuelle.
  const traces = [];
  activeDesc.forEach((desc, i) => {
    const yAxis = i === 0 ? "y" : "y" + (i + 1);
    state.activeMotifs.forEach(motif => {
      const arr = grouped[motif] || [];   // déjà trié par t croissant
      traces.push({
        x:      arr.map(r => r.t),
        y:      arr.map(r => r[desc.key]),
        type:   "scatter",
        mode:   state.showMarkers ? "lines+markers" : "lines",
        name:   motif,
        legendgroup: motif,
        showlegend: i === 0,
        xaxis:  "x",
        yaxis:  yAxis,
        line:   { color: MOTIF_COLORS[motif], width: 2.5, shape: "linear" },
        marker: { color: MOTIF_COLORS[motif], size: 7, line: { width: 0.8, color: "#ffffff" } },
        hoverlabel: { font: { family: SERIF, size: 13 } },
        hovertemplate:
          `<b>${motif}</b><br>` +
          `<i>${desc.label}</i> = %{y:.4f}<br>` +
          `t = %{x:+.1f} fs<br>` +
          `MD step = %{customdata[0]} (%{customdata[1]})` +
          `<extra></extra>`,
        customdata: arr.map(r => [r.step, r.direction]),
      });
    });
  });

  // Layout multi-panneaux empilés
  const n = activeDesc.length;
  const gap = 0.055;                 // espace inter-panneaux (domaine)
  const totalGap = gap * (n - 1);
  const h = (1 - totalGap) / n;

  // Légende horizontale, EN HAUT et hors plot → ne recouvre jamais les courbes
  const layout = {
    grid: { rows: n, columns: 1, pattern: "independent" },
    height: Math.max(360 + 280 * n, 540),
    margin: { l: 90, r: 40, t: 90, b: 80 },
    hovermode: "closest",
    font: { family: SERIF, size: 14, color: "#1e293b" },
    legend: {
      orientation: "h",
      x: 0.5, xanchor: "center",
      y: 1.04, yanchor: "bottom",
      font: { family: SERIF, size: 14 },
      bgcolor: "rgba(255,255,255,0)",
      bordercolor: "rgba(0,0,0,0)",
    },
    plot_bgcolor: "#ffffff",
    paper_bgcolor: "#ffffff",
    shapes: [],
    annotations: [],
  };

  activeDesc.forEach((desc, i) => {
    const yKey   = i === 0 ? "yaxis" : "yaxis" + (i + 1);
    const yRef   = i === 0 ? "y" : "y" + (i + 1);
    const yDomTop = 1 - i * (h + gap);
    const yDomBot = yDomTop - h;
    layout[yKey] = {
      title: {
        text: desc.axis,
        font: { family: SERIF, size: 16, color: "#0f172a" },
        standoff: 12,
      },
      domain: [yDomBot, yDomTop],
      gridcolor: "#e2e8f0",
      gridwidth: 1,
      zerolinecolor: "#94a3b8",
      zerolinewidth: 1,
      tickfont: { family: SERIF, size: 13, color: "#334155" },
      ticks: "outside", ticklen: 5, tickwidth: 1, tickcolor: "#64748b",
      showline: true, linewidth: 1, linecolor: "#334155", mirror: true,
      automargin: true,
      minor: state.showMinorGrid
        ? { showgrid: true, gridcolor: "#f1f5f9", gridwidth: 0.5, ticklen: 3, tickcolor: "#94a3b8" }
        : { showgrid: false },
    };

    // Ligne TS verticale par panneau
    if (state.showTSLine) {
      layout.shapes.push({
        type: "line", x0: 0, x1: 0, xref: "x", yref: `${yRef} domain`,
        y0: 0, y1: 1, line: { color: "#94a3b8", width: 1.5, dash: "dash" },
        layer: "above",
      });
    }

    // Zones colorées reverse / TS / forward — très discrètes
    if (state.showZones) {
      layout.shapes.push(
        { type: "rect", xref: "x", yref: `${yRef} domain`,
          x0: tMin - 5, x1: -6, y0: 0, y1: 1,
          fillcolor: "#0072B2", opacity: 0.04, line: { width: 0 }, layer: "below" },
        { type: "rect", xref: "x", yref: `${yRef} domain`,
          x0: -6, x1: 6, y0: 0, y1: 1,
          fillcolor: "#E69F00", opacity: 0.10, line: { width: 0 }, layer: "below" },
        { type: "rect", xref: "x", yref: `${yRef} domain`,
          x0: 6, x1: tMax + 5, y0: 0, y1: 1,
          fillcolor: "#CC79A7", opacity: 0.04, line: { width: 0 }, layer: "below" },
      );
    }

    // Annotation TS — seulement sur le panneau du haut, discrète (gris 60 %)
    if (i === 0 && state.showTSLine) {
      layout.annotations.push({
        x: 0, y: 1.015, xref: "x", yref: `${yRef} domain`,
        text: "TS", showarrow: false,
        font: { color: "rgba(71,85,105,0.7)", size: 12, family: SERIF, style: "italic" },
      });
    }
  });

  // Axe X commun (ancré sur dernier panneau)
  layout.xaxis = {
    title: {
      text: "Temps  <i>t</i>  (fs)  —  TS à <i>t</i> = 0  —  reverse ← | → forward",
      font: { family: SERIF, size: 15, color: "#0f172a" },
      standoff: 14,
    },
    range: [tMin - 5, tMax + 5],
    gridcolor: "#e2e8f0",
    gridwidth: 1,
    zeroline: false,
    tickfont: { family: SERIF, size: 13, color: "#334155" },
    ticks: "outside", ticklen: 5, tickwidth: 1, tickcolor: "#64748b",
    showline: true, linewidth: 1, linecolor: "#334155", mirror: true,
    anchor: "y" + n,
    automargin: true,
    minor: state.showMinorGrid
      ? { showgrid: true, gridcolor: "#f1f5f9", gridwidth: 0.5, ticklen: 3, tickcolor: "#94a3b8" }
      : { showgrid: false },
  };

  // Petits axes X masqués pour chaque panneau intermédiaire (cosmétique)
  for (let i = 1; i < n; i++) {
    layout["xaxis" + (i + 1)] = {
      matches: "x", showticklabels: false, showgrid: true, gridcolor: "#e2e8f0",
      showline: true, linewidth: 1, linecolor: "#334155", mirror: true,
      anchor: "y" + (i + 1),
    };
  }

  const config = {
    displaylogo: false,
    responsive: true,
    toImageButtonOptions: { format: "png", filename: autoFilename(""), scale: 3 },
    modeBarButtonsToRemove: ["sendDataToCloud", "lasso2d", "select2d"],
  };

  Plotly.react("plot", traces, layout, config);
  renderTable();
}

// =================== TABLEAU ===================
const TBL_COLS = [
  ["motif", "Motif"], ["t", "t (fs)"], ["step", "Step"], ["direction", "Dir"],
  ["HOMA", "HOMA"], ["EN", "EN"], ["GEO", "GEO"], ["HOMA_TOTAL", "HOMA total"],
  ["LDM_FROB", "LDM Frob"], ["LDM_OFF", "LDM Off"], ["LDM_DIAG", "LDM Diag"],
  ["LDM_RMSD", "LDM RMSD"], ["LDM_CT_PCT", "LDM CT %"],
  ["Q_FROB", "Q Frob"], ["Q_RMSD", "Q RMSD"], ["Q_G", "Q(G)"],
  ["S_HOM", "S-hom"], ["S_HOM_G", "S-hom(G)"], ["S_G", "S(G)"],
];

function initTableHead() {
  const tr = document.getElementById("tbl-head");
  tr.innerHTML = TBL_COLS.map(c => `<th class="px-2 py-1 text-left">${c[1]}</th>`).join("");
  const sel = document.getElementById("tbl-motif-filter");
  Object.keys(MOTIF_COLORS).forEach(m => {
    const opt = document.createElement("option");
    opt.value = m; opt.textContent = m;
    sel.appendChild(opt);
  });
  ["tbl-search", "tbl-motif-filter", "tbl-tmin", "tbl-tmax"].forEach(id =>
    document.getElementById(id).addEventListener("input", renderTable));
  document.getElementById("tbl-export").addEventListener("click", exportCSV);
}

function getFilteredRows() {
  const q = document.getElementById("tbl-search").value.toLowerCase();
  const motifFilter = document.getElementById("tbl-motif-filter").value;
  const tmin = parseFloat(document.getElementById("tbl-tmin").value);
  const tmax = parseFloat(document.getElementById("tbl-tmax").value);
  return DATA.filter(r => {
    if (motifFilter && r.motif !== motifFilter) return false;
    if (!isNaN(tmin) && r.t < tmin) return false;
    if (!isNaN(tmax) && r.t > tmax) return false;
    if (q) {
      const txt = `${r.motif} ${r.t} ${r.step} ${r.direction}`.toLowerCase();
      if (!txt.includes(q)) return false;
    }
    return true;
  });
}

function renderTable() {
  const rows = getFilteredRows().sort((a, b) => a.t - b.t || a.motif.localeCompare(b.motif));
  const body = document.getElementById("tbl-body");
  body.innerHTML = rows.map(r => {
    const tds = TBL_COLS.map(([k]) => {
      let v = r[k];
      if (typeof v === "number") {
        if (k === "t") v = v.toFixed(1);
        else if (k === "step") v = String(v);
        else v = v.toFixed(4);
      }
      let style = "";
      if (k === "motif") style = `style="color:${MOTIF_COLORS[r.motif]};font-weight:600"`;
      return `<td class="px-2 py-1" ${style}>${v ?? ""}</td>`;
    }).join("");
    return `<tr class="tbl-row">${tds}</tr>`;
  }).join("");
  document.getElementById("tbl-count").textContent =
    `${rows.length} lignes affichées sur ${DATA.length} total`;
}

function exportCSV() {
  const rows = getFilteredRows().sort((a, b) => a.t - b.t || a.motif.localeCompare(b.motif));
  const sep = ",";
  const head = TBL_COLS.map(c => c[1]).join(sep);
  const body = rows.map(r => TBL_COLS.map(([k]) => {
    const v = r[k];
    if (typeof v === "string" && v.includes(sep)) return `"${v.replace(/"/g, '""')}"`;
    return v ?? "";
  }).join(sep)).join("\n");
  const blob = new Blob([head + "\n" + body], { type: "text/csv;charset=utf-8" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "DM_B2_descripteurs.csv";
  a.click();
}

// =================== BOOT ===================
initMotifChips();
initDescriptorGrid();
initControls();
initTableHead();
render();
</script>
</body>
</html>
"""

    html = html.replace("__DATA_PLACEHOLDER__", data_json)
    return html


def main():
    rows = load_data()
    print(f"Chargé {len(rows)} lignes depuis {XLSX.name}")
    html = build_html(rows)
    OUT.write_text(html, encoding="utf-8")
    print(f"Visualisation HTML écrite : {OUT}")
    print(f"  taille : {OUT.stat().st_size / 1024:.1f} KB")
    print(f"  -> ouvre dans un navigateur (double-clic)")


if __name__ == "__main__":
    main()
