"""
╔══════════════════════════════════════════════════════════════════════════════╗
║         Dashboard Électrique — Relevé de Paramètres                        ║
║         Fichier : dashboard_releve.py                                       ║
╠══════════════════════════════════════════════════════════════════════════════╣
║  INSTALLATION  :  pip install streamlit pandas plotly openpyxl numpy       ║
║  LANCEMENT     :  streamlit run dashboard_releve.py                         ║
╠══════════════════════════════════════════════════════════════════════════════╣
║  INDEX DE NAVIGATION — NUMÉROS DE LIGNES                                   ║
║  ─────────────────────────────────────────────────────────────────────────  ║
║  [L  30]  CONFIG PAGE      → titre, icône, layout                          ║
║  [L  45]  CSS / THÈME      → couleurs sidebar, métriques, bandeau panel    ║
║  [L  75]  FILTRES FEUILLES → noms de feuilles FOLLOWED / HISTORIQUE        ║
║  [L  88]  PARSER FOLLOWED  → lecture structure Excel feuilles FOLLOWED     ║
║  [L 147]  PARSER HISTO     → lecture structure Excel feuilles HISTORIQUE   ║
║  [L 220]  CHARGEMENT EXCEL → cache + dispatch par type de feuille          ║
║  [L 235]  PALETTE COULEURS → liste des couleurs des graphiques             ║
║  [L 242]  RENDER FOLLOWED  → KPIs + graphiques barres + taux de charge     ║
║    [L 260]  → Couleurs barres J i-1 / J i                                  ║
║    [L 270]  → Hauteur & marge bas (dynamiques selon nb équipements)        ║
║    [L 278]  → Titre du graphique barres                                    ║
║    [L 285]  → Position & style de la légende                               ║
║    [L 296]  → Scroll horizontal si > 25 équipements                        ║
║    [L 308]  → Taux de charge (seuils couleurs vert/orange/rouge)           ║
║    [L 325]  → Nuage de points P vs Q vs S                                  ║
║    [L 338]  → Tableau récapitulatif dépliable                              ║
║  [L 355]  RENDER HISTO     → courbes temporelles P/Q/S par équipement     ║
║  [L 420]  SIDEBAR UPLOAD   → bouton import fichier Excel                   ║
║  [L 438]  GUARD FICHIER    → arrêt si aucun fichier importé               ║
║  [L 445]  SIDEBAR NAV      → radio vue + selectbox feuille                 ║
║  [L 463]  VUE GLOBALE      → synthèse tous panels (barres + taux charge)  ║
║  [L 510]  VUE FOLLOWED     → onglets par sous-panel FOLLOWED               ║
║  [L 525]  VUE HISTORIQUE   → onglets par sous-panel HISTORIQUE             ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.colors as pc
import openpyxl
from io import BytesIO
import re

# ============================================================================
# UTILITAIRES GLOBAUX
# ============================================================================

from typing import Optional

def to_float(value, default=None):
    """
    Convertit une valeur en float.

    Parameters
    ----------
    value : any
        Valeur à convertir

    default : any
        Valeur retournée en cas d'échec

    Returns
    -------
    float | default
    """
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def tc_color(tc: float) -> str:
    """
    Retourne la couleur associée au taux de charge.
    """

    if tc < 50:
        return "#e0ad21"

    if tc < 70:
        return "#bdcb26"

    if tc < 100:
        return "#096806"

    return "#6D0505"


def show_metric(col, label, value, delta=None):
    """
    Affichage uniforme des KPI Streamlit.
    """

    col.metric(
        label=label,
        value=value,
        delta=delta
    )
# ══════════════════════════════════════════════════════════════════════════════
# [L 30] CONFIG PAGE
# ══════════════════════════════════════════════════════════════════════════════
# → Pour changer : titre onglet navigateur (page_title), icône (page_icon),
#   ou forcer la sidebar fermée au démarrage (initial_sidebar_state="collapsed")

st.set_page_config(
    page_title="Relevé de Paramètres Électriques",  # ← titre onglet navigateur
    page_icon="⚡",                                   # ← icône onglet
    layout="wide",                                    # ← "wide" ou "centered"
    initial_sidebar_state="expanded",                 # ← "expanded" ou "collapsed"
)


# ══════════════════════════════════════════════════════════════════════════════
# [L 45] CSS / THÈME
# ══════════════════════════════════════════════════════════════════════════════
# → Pour changer la couleur de fond de la sidebar       : modifier #0f1923
# → Pour changer la couleur du bandeau de panel        : modifier border-left et background
# → Pour changer les couleurs des badges ok/warn/crit  : modifier les 3 dernières règles

st.markdown("""
<style>
/* Fond et texte de la barre latérale */
[data-testid="stSidebar"] { background:#0f1923; }
[data-testid="stSidebar"] * { color:#c8d6e5 !important; } # couleurs texte sidebar
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stMultiSelect label,
[data-testid="stSidebar"] .stFileUploader label { 
    color:#7f8c9a !important;
    font-size:0.78rem !important;
}

/* Cartes KPI métriques */
div[data-testid="metric-container"] {
    background:#0f1923;
    border:1px solid #1e3048;
    border-radius:10px;
    padding:10px 14px;
}
div[data-testid="metric-container"] label {
    color:#7f8c9a !important;
    font-size:0.72rem !important;
}

/* Bandeau titre de sous-panel (barre bleue à gauche) */
.panel-header {
    background: linear-gradient(90deg, #1e3048, #0f1923); /* ← dégradé fond */
    border-left: 4px solid #3498db;                        /* ← couleur barre gauche */
    padding: 8px 14px;
    border-radius: 4px;
    font-size: 0.95rem;
    font-weight: 600;
    color: #e0eaf5;
    margin: 10px 0 6px;
}

/* Badges de statut dans les alertes */
.ok-badge   { color:#27ae60; font-weight:600; }   /* ← vert  : normal  */
.warn-badge { color:#f39c12; font-weight:600; }   /* ← orange: warning */
.crit-badge { color:#e74c3c; font-weight:600; }   /* ← rouge : critique */
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# [L 75] FILTRES FEUILLES — noms reconnus comme FOLLOWED ou HISTORIQUE
# ══════════════════════════════════════════════════════════════════════════════
# → Pour ajouter un préfixe reconnu, ajoutez une condition dans ces fonctions.
#   Ex : reconnaître "SUIVI" → ajouter : or n.startswith("SUIVI")

def is_followed_sheet(name: str) -> bool:
    """Retourne True si le nom de feuille commence par FOLLOWED ou FOLLOW."""
    n = name.upper()
    return n.startswith("FOLLOWED") or n.startswith("FOLLOW ")

def is_historique_sheet(name: str) -> bool:
    """Retourne True si le nom de feuille commence par HISTORIQUE ou HISTO."""
    n = name.upper()
    return n.startswith("HISTORIQUE") or n.startswith("HISTO")


# ══════════════════════════════════════════════════════════════════════════════
# [L 88] PARSER FOLLOWED — lecture des feuilles de type FOLLOWED
# ══════════════════════════════════════════════════════════════════════════════
# Structure attendue dans Excel :
#   - Ligne titre de sous-panel : col A = texte (ex "FOLLOW BY 6,6kV"), col B vide
#   - Ligne en-tête colonnes   : col B = "Equipments" ou "EQUIPMENT"
#   - Lignes de données        : col A = numéro, col B = nom équipement,
#                                col C = Pn, D=P_i1, E=P_i, F=Q_i1, G=Q_i,
#                                H=S_i1, I=S_i, J=Time_i1, K=Time_i,
#                                L=Tc_i1, M=Tc_i

def parse_followed_sheet(ws) -> list:
    """
    Parcourt la feuille et retourne une liste de sous-panels :
      [ { 'panel': str, 'data': DataFrame }, ... ]
    """
    rows = list(ws.iter_rows(max_row=300, values_only=True))
    panels = []
    current_panel = None
    header_row    = None
    data_rows     = []

    for i, row in enumerate(rows):
        non_none = [v for v in row if v is not None]
        if not non_none:
            continue  # ligne vide → on passe

        # Détection titre de sous-panel : col A = texte, col B vide
        is_title = (
            isinstance(row[0], str) and len(row[0].strip()) > 2
            and not isinstance(row[0], (int, float))
            and row[1] is None
        )
        # Détection ligne d'en-tête : col B = "Equipments" / "EQUIPMENT" / "The Transformers"
        is_header = (
            row[1] is not None
            and str(row[1]).strip() in ['Equipments', 'EQUIPMENT', 'The Transformers']
        )
        # Détection ligne de données : col A = numéro ou vide, col B = nom, col C non nulle
        is_data = (
            (isinstance(row[0], (int, float)) or row[0] is None)
            and row[1] is not None
            and str(row[1]).strip() not in ['Equipments', 'EQUIPMENT', 'The Transformers']
            and row[2] is not None   # Pn/Sn non nul
        )

        if is_title:
            # Sauvegarder le sous-panel précédent avant d'en démarrer un nouveau
            if current_panel and data_rows:
                panels.append(_build_followed_panel(current_panel, data_rows))
            current_panel = str(row[0]).strip()
            header_row    = None
            data_rows     = []

        elif is_header:
            header_row = i   # mémoriser l'index de ligne d'en-tête

        elif is_data and header_row is not None:
            data_rows.append(row)   # accumuler les lignes de données

    # Ne pas oublier le dernier sous-panel
    if current_panel and data_rows:
        panels.append(_build_followed_panel(current_panel, data_rows))

    return panels


def _build_followed_panel(panel_name: str, raw_rows: list) -> dict:
    """
    Construction du DataFrame d'un sous-panel FOLLOWED.
    """

    records = []

    for row in raw_rows:

        records.append(
            {
                "equipement": str(row[1]).strip() if row[1] else "—",

                "Pn_kVA": to_float(row[2], 0),

                "P_i1": to_float(row[3]),
                "P_i": to_float(row[4]),

                "Q_i1": to_float(row[5]),
                "Q_i": to_float(row[6]),

                "S_i1": to_float(row[7]),
                "S_i": to_float(row[8]),

                "Time_i1": str(row[9]) if row[9] else "—",
                "Time_i": str(row[10]) if row[10] else "—",

                "Tc_i1": (
                    to_float(row[11]) * 100
                    if row[11] not in (None, "")
                    else None
                ),

                "Tc_i": (
                    to_float(row[12]) * 100
                    if row[12] not in (None, "")
                    else None
                ),
            }
        )

    return {
        "panel": panel_name,
        "data": pd.DataFrame(records)
    }


# ══════════════════════════════════════════════════════════════════════════════
# [L 147] PARSER HISTORIQUE — lecture des feuilles de type HISTORIQUE
# ══════════════════════════════════════════════════════════════════════════════
# Structure attendue dans Excel :
#   - Ligne titre de sous-panel  : col A = texte, col B vide
#   - Ligne en-tête              : col A = "EQUIPEMENTS", puis blocs répétés :
#                                  DATE | TIME | P | Q | S | DATE | TIME | ...
#   - Lignes de données          : col A = nom équipement, colonnes suivantes = valeurs

def parse_historique_sheet(ws) -> list:
    """
    Parcourt la feuille et retourne une liste de sous-panels historiques :
      [ { 'panel': str, 'data': DataFrame }, ... ]
    """
    rows = list(ws.iter_rows(max_row=500, values_only=True))
    panels      = []
    current_panel = None
    header_line   = None
    data_start    = None

    for i, row in enumerate(rows):
        non_none = [v for v in row if v is not None]
        if not non_none:
            continue

        # Titre de sous-panel : col A texte libre (pas "EQUIPEMENTS")
        is_subpanel_title = (
            isinstance(row[0], str)
            and str(row[0]).strip() not in [
                'EQUIPEMENTS',
                'Historique des Puissances releves des Equipements'
            ]
            and row[1] is None
            and len(str(row[0]).strip()) > 2
        )
        # En-tête de bloc données : col A == "EQUIPEMENTS"
        is_header = (row[0] == 'EQUIPEMENTS')

        if is_subpanel_title:
            # Sauvegarder le sous-panel précédent
            if current_panel is not None and header_line is not None and data_start is not None:
                df = _build_histo_panel(rows, header_line, data_start, i)
                if df is not None and not df.empty:
                    panels.append({"panel": current_panel, "data": df})
            current_panel = str(row[0]).strip()
            header_line   = None
            data_start    = None

        elif is_header:
            header_line = i
            data_start  = i + 1   # les données commencent à la ligne suivante

    # Dernier sous-panel
    if current_panel is not None and header_line is not None and data_start is not None:
        df = _build_histo_panel(rows, header_line, data_start, len(rows))
        if df is not None and not df.empty:
            panels.append({"panel": current_panel, "data": df})

    return panels


def _build_histo_panel(rows, header_line, data_start, data_end):
    """
    Construit le DataFrame historique en détectant les blocs DATE/TIME/P/Q/S
    répétés horizontalement, et en les empilant verticalement.
    Colonnes produites : equipement, datetime, P_kW, Q_kVar, S_kVA
    """
    header = rows[header_line]

    # Trouver toutes les positions de colonnes DATE dans la ligne d'en-tête
    # → chaque position DATE marque le début d'un bloc de 5 colonnes
    date_cols = [j for j, v in enumerate(header) if str(v).strip() == 'DATE']
    if not date_cols:
        return None   # pas de colonne DATE → feuille non reconnue

    records = []
    for i in range(data_start, data_end):
        row = rows[i]
        if not any(v is not None for v in row):
            continue   # ligne vide

        eq = str(row[0]).strip() if row[0] else None
        if not eq or eq == 'EQUIPEMENTS':
            continue   # en-tête répété → on passe

        # Pour chaque bloc DATE/TIME/P/Q/S
        for dc in date_cols:
            if dc + 4 >= len(row):
                continue
            date_val = row[dc]      # DATE
            time_val = row[dc + 1]  # TIME  (format "08H30" ou "08:30")
            p_val    = row[dc + 2]  # P (kW)
            q_val    = row[dc + 3]  # Q (kVAR)
            s_val    = row[dc + 4]  # S (kVA)

            if date_val is None and p_val is None:
                continue   # bloc vide

            # Construction du datetime combiné date + heure
            try:
                if isinstance(date_val, pd.Timestamp) or hasattr(date_val, 'date'):
                    dt = pd.Timestamp(date_val)
                    if time_val and isinstance(time_val, str):
                        # Nettoyer "08H30" → "08:30"
                        t_clean = re.sub(r'[Hh]', ':', time_val).strip()
                        parts   = t_clean.split(':')
                        if len(parts) >= 2:
                            dt = dt.replace(hour=int(parts[0]), minute=int(parts[1]))
                else:
                    dt = None
            except Exception:
                dt = None

            records.append({
                "equipement": eq,
                "datetime":   dt,
                "P_kW":   float(p_val) if p_val not in (None, "") else None,
                "Q_kVar": float(q_val) if q_val not in (None, "") else None,
                "S_kVA":  float(s_val) if s_val not in (None, "") else None,
            })

    if not records:
        return None
    df = pd.DataFrame(records)
    df = df.dropna(subset=["P_kW", "Q_kVar"], how="all")
    return df


# ══════════════════════════════════════════════════════════════════════════════
# [L 220] CHARGEMENT EXCEL — cache + dispatch par type de feuille
# ══════════════════════════════════════════════════════════════════════════════
# → @st.cache_data évite de relire le fichier à chaque interaction Streamlit.
# → Pour forcer la relecture (ex: debug), appuyer sur "Clear cache" dans le menu ⋮

@st.cache_data(show_spinner="Lecture du fichier Excel…")
def load_excel(file_bytes: bytes) -> dict:
    """
    Charge le fichier Excel et dispatche chaque feuille vers le bon parser.
    Retourne :
      {
        "followed":   { "Nom feuille": [ {panel, data}, ... ], ... },
        "historique": { "Nom feuille": [ {panel, data}, ... ], ... }
      }
    """
    wb = openpyxl.load_workbook(BytesIO(file_bytes), read_only=True, data_only=True)
    result = {"followed": {}, "historique": {}}
    for name in wb.sheetnames:
        ws = wb[name]
        if is_followed_sheet(name):
            panels = parse_followed_sheet(ws)
            if panels:
                result["followed"][name] = panels
        elif is_historique_sheet(name):
            panels = parse_historique_sheet(ws)
            if panels:
                result["historique"][name] = panels
    return result


# ══════════════════════════════════════════════════════════════════════════════
# [L 235] PALETTE COULEURS — graphiques multi-séries
# ══════════════════════════════════════════════════════════════════════════════
# → Pour changer les couleurs des courbes/barres multi-équipements,
#   modifiez cette liste (codes hex). Ordre = ordre des équipements affichés.

COLORS = [
    "#3498db",  # bleu
    "#2ecc71",  # vert
    "#e74c3c",  # rouge
    "#f39c12",  # orange
    "#9b59b6",  # violet
    "#1abc9c",  # turquoise
    "#e67e22",  # orange foncé
    "#e91e63",  # rose
    "#00bcd4",  # cyan
    "#8bc34a",  # vert clair
]


# ══════════════════════════════════════════════════════════════════════════════
# [L 242] RENDER FOLLOWED — affichage d'un sous-panel FOLLOWED
# ══════════════════════════════════════════════════════════════════════════════

def render_followed_panel(panel_info: dict):
    """
    Affiche un sous-panel FOLLOWED complet :
      - Bandeau titre
      - 4 KPIs (nb équipements, P Jour i, S Jour i, taux de charge moyen)
      - Graphique barres groupées P Jour i-1 vs Jour i
      - Graphique taux de charge horizontal (vert/orange/rouge)
      - Nuage de points P vs Q (taille = S)
      - Tableau récapitulatif dépliable
    """
    name = panel_info["panel"]
    df   = panel_info["data"].copy()

    if df.empty:
        st.info(f"Aucune donnée pour : {name}")
        return

    # Nettoyage : retirer les lignes sans nom d'équipement
    df       = df[df["equipement"] != "—"].copy()
    df_valid = df.dropna(subset=["P_i", "S_i"], how="all")

    # Bandeau titre du sous-panel
    st.markdown(f'<div class="panel-header">⚡ {name}</div>', unsafe_allow_html=True)

    # ── KPIs ─────────────────────────────────────────────────────────────────
    c1, c2, c3, c4 = st.columns(4)
    nb_eq      = len(df_valid)
    p_total_i  = df_valid["P_i"].sum()  if "P_i"  in df_valid.columns else 0
    p_total_i1 = df_valid["P_i1"].sum() if "P_i1" in df_valid.columns else 0
    q_total_i  = df_valid["Q_i"].sum()  if "Q_i"  in df_valid.columns else 0
    q_total_i1 = df_valid["Q_i1"].sum() if "Q_i1" in df_valid.columns else 0
    tc_moy     = (df_valid["Tc_i"].mean()
                  if "Tc_i" in df_valid.columns and df_valid["Tc_i"].notna().any()
                  else None)
    delta_p = (
    p_total_i - p_total_i1
    if p_total_i1 is not None
    else None
)
    delta_q = (
    q_total_i - q_total_i1
    if q_total_i1 is not None
    else None
)
    c1.metric("Équipements", f"{nb_eq}")
    c2.metric("P active Jour i", f"{p_total_i:.1f} kW",
              delta=f"{delta_p:+.1f} kW vs J i-1" if delta_p is not None else None)
    c3.metric("Q réactive Jour i", f"{q_total_i:.1f} kVAR",
              delta=f"{delta_q:+.1f} kVAR vs J i-1" if delta_q is not None else None)
    # c4.metric("Taux de charge moy.", f"{tc_moy:.1f} %" if tc_moy is not None else "—")

    if df_valid.empty:
        st.caption("Toutes les valeurs sont nulles/vides pour cette session.")
        return

    col_g, col_d = st.columns([3, 2])

    # ── Graphique barres comparatif P i-1 vs P i ─────────────────────────────
    with col_g:
        df_bar = df_valid[["equipement", "P_i1", "P_i"]].copy().fillna(0)
        fig_bar = go.Figure()

        # [L 260] COULEURS BARRES
        # → Jour i-1 : changer "rgba(180,190,200,0.6)" pour une autre couleur grise
        # → Jour i   : changer "#1a8fe3" pour une autre couleur bleue
        fig_bar.add_trace(go.Bar(
            name="P Jour i-1",
            x=df_bar["equipement"],
            y=df_bar["P_i1"],
            marker_color="rgba(180,190,200,0.6)",  # ← gris argenté Jour i-1
            marker_line_color="#a0aab5",
            marker_line_width=1,
        ))
        fig_bar.add_trace(go.Bar(
            name="P Jour i",
            x=df_bar["equipement"],
            y=df_bar["P_i"],
            marker_color="#1a8fe3",                # ← bleu vif Jour i
            marker_line_color="#0d6ebc",
            marker_line_width=1,
        ))
    # ── Graphique barres comparatif Q i-1 vs Q i ─────────────────────────────
    with col_g:
        df_barr = df_valid[["equipement", "Q_i1", "Q_i"]].copy().fillna(0)
        fig_barr = go.Figure()

        # [L 260] COULEURS BARRES
        # → Jour i-1 : changer "rgba(180,190,200,0.6)" pour une autre couleur grise
        # → Jour i   : changer "#1a8fe3" pour une autre couleur bleue
        fig_barr.add_trace(go.Bar(
            name="Q Jour i-1",
            x=df_barr["equipement"],
            y=df_barr["Q_i1"],
            marker_color="rgba(180,190,200,0.6)",  # ← gris argenté Jour i-1
            marker_line_color="#a0aab5",
            marker_line_width=1,
        ))
        fig_barr.add_trace(go.Bar(
            name="Q Jour i",
            x=df_barr["equipement"],
            y=df_barr["Q_i"],
            marker_color="#1a8fe3",                # ← bleu vif Jour i
            marker_line_color="#0d6ebc",
            marker_line_width=1,
        ))

        # [L 270] HAUTEUR & MARGE FIXES
        # → Pour changer la hauteur : modifier hauteur=300
        # → Pour changer la marge basse (espace labels) : modifier marge_bas=60
        n_eq      = len(df_bar)   # encore utilisé pour le scroll horizontal [L 296]
        marge_bas = 60
        hauteur   = 300

        fig_bar.update_layout(
            barmode="group",

            # [L 278] TITRE DU GRAPHIQUE BARRES
            # → Modifier "text" pour changer l'intitulé
            # → x=0.5 / xanchor="center" → centré ; x=0 → aligné à gauche
            title=dict(
                text="Puissance active P (kW) — Jour i-1 vs Jour i",
                x=0.5,
                xanchor="center",
                y=1.0,
                yanchor="top",
                font=dict(size=13),
            ),
            template="plotly_dark",
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="#0a1520",
            height=hauteur,
            margin=dict(t=70, b=marge_bas, l=50, r=10),

            xaxis=dict(
                tickangle=-90,         # ← labels verticaux (−90°) évite le chevauchement
                tickfont=dict(size=9), # ← taille police labels axe X
                gridcolor="#1e3048",
            ),
            yaxis=dict(title="kW", gridcolor="#1e3048"),

            # [L 285] LÉGENDE
            # → y=1.0 / yanchor="top" place la légende dans la marge haute (t=70)
            # → x=0.5 / xanchor="center" la centre horizontalement
            # → Pour déplacer à gauche : x=0, xanchor="left"
            legend=dict(
                orientation="h",
                x=0.5,
                xanchor="center",
                y=1.0,
                yanchor="top", 
                font=dict(size=11, color="#ffffff"),
                bgcolor="rgba(15,25,35,0.8)",   # ← fond semi-transparent légende
                bordercolor="#dfe0e2",
                borderwidth=1,
            ),
        )

        # [L 296] SCROLL HORIZONTAL si trop d'équipements (> 25)
        # → Pour désactiver le scroll, supprimer ce bloc if
        # → Pour changer le seuil, modifier "25"
        if n_eq > 25:
            fig_bar.update_layout(
                xaxis=dict(
                    tickangle=-90,
                    tickfont=dict(size=9),
                    gridcolor="#1e3048",
                    rangeslider=dict(visible=True, thickness=0.04),  # ← mini-carte de navigation
                )
            )
        st.plotly_chart(fig_bar, use_container_width=True)

        # [L 270] HAUTEUR & MARGE FIXES
        # → Pour changer la hauteur : modifier hauteur=300
        # → Pour changer la marge basse (espace labels) : modifier marge_bas=60
        n_eq      = len(df_barr)   # encore utilisé pour le scroll horizontal [L 296]
        marge_bas = 60
        hauteur   = 300

        fig_barr.update_layout(
            barmode="group",

            # [L 278] TITRE DU GRAPHIQUE BARRES
            # → Modifier "text" pour changer l'intitulé
            # → x=0.5 / xanchor="center" → centré ; x=0 → aligné à gauche
            title=dict(
                text="Puissance réactive Q (kVAR) — Jour i-1 vs Jour i",
                x=0.5,
                xanchor="center",
                y=1.0,
                yanchor="top",
                font=dict(size=13),
            ),
            template="plotly_dark",
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="#0a1520",
            height=hauteur,
            margin=dict(t=70, b=marge_bas, l=50, r=10),

            xaxis=dict(
                tickangle=-90,         # ← labels verticaux (−90°) évite le chevauchement
                tickfont=dict(size=9), # ← taille police labels axe X
                gridcolor="#1e3048",
            ),
            yaxis=dict(title="kVAR", gridcolor="#1e3048"),

            # [L 285] LÉGENDE
            # → y=1.0 / yanchor="top" place la légende dans la marge haute (t=70)
            # → x=0.5 / xanchor="center" la centre horizontalement
            # → Pour déplacer à gauche : x=0, xanchor="left"
            legend=dict(
                orientation="h",
                x=0.5,
                xanchor="center",
                y=1.0,
                yanchor="top", 
                font=dict(size=11, color="#ffffff"),
                bgcolor="rgba(15,25,35,0.8)",   # ← fond semi-transparent légende
                bordercolor="#ffffff",
                borderwidth=1,
            ),
        )

        # [L 296] SCROLL HORIZONTAL si trop d'équipements (> 25)
        # → Pour désactiver le scroll, supprimer ce bloc if
        # → Pour changer le seuil, modifier "25"
        if n_eq > 25:
            fig_barr.update_layout(
                xaxis=dict(
                    tickangle=-90,
                    tickfont=dict(size=9),
                    gridcolor="#1e3048",
                    rangeslider=dict(visible=True, thickness=0.04),  # ← mini-carte de navigation
                )
            )
        st.plotly_chart(fig_barr, use_container_width=True)

    # ── Taux de charge horizontal ─────────────────────────────────────────────
    with col_d:
        # [L 308] SEUILS COULEURS TAUX DE CHARGE
        # → Vert  : Tc < 70%  → modifier "70"
        # → Vert: 70% ≤ Tc < 100% → modifier "100"
        # → Rouge : Tc ≥ 100%
        df_tc = df_valid[["equipement", "Tc_i"]].dropna().sort_values("Tc_i", ascending=True)
        if not df_tc.empty:
            colors_tc = [
    tc_color(v)
    for v in df_tc["Tc_i"]
]
            fig_tc = go.Figure(go.Bar(
                x=df_tc["Tc_i"],
                y=df_tc["equipement"],
                orientation="h",
                marker_color=colors_tc,
                text=[f"{v:.1f}%" for v in df_tc["Tc_i"]],
                textposition="outside",
            ))
            fig_tc.add_vline(x=100, line_dash="dash", line_color="#e74c3c")  # ← ligne 100%
            fig_tc.update_layout(
                title="Load factor jour i (%)",
                template="plotly_dark",
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="#0a1520",
                height=300,
                margin=dict(t=40, b=10, l=10, r=40),
                xaxis=dict(range=[0, max(110, df_tc["Tc_i"].max() * 1.1)], gridcolor="#1e3048"),
                yaxis=dict(gridcolor="#1e3048"),
                showlegend=False,
            )
            st.plotly_chart(fig_tc, use_container_width=True)

    # ── [L 325] Nuage de points P vs Q (taille bulle = S) ────────────────────
    # → Pour masquer ce graphique, commenter les lignes jusqu'à st.plotly_chart(fig_pqs)
    df_pqs = df_valid[["equipement", "P_i", "Q_i", "S_i"]].dropna()
    if not df_pqs.empty:
        n_eq_pqs = len(df_pqs["equipement"].unique())
        palette  = pc.sample_colorscale("Turbo", n_eq_pqs)
        fig_pqs = px.scatter(
            df_pqs, x="P_i", y="Q_i", size="S_i", color="equipement",
            title="P vs. Q Distribution (bubble size = S) - Jour i",
            template="plotly_dark",
            color_discrete_sequence=palette,
            labels={"P_i": "P (kW)", "Q_i": "Q (kVAR)", "equipement": "Équipement"},
        )
        fig_pqs.for_each_trace(
            lambda t: t.update(showlegend=False) if t.x is None or len(t.x) == 0 else ()
        )
        fig_pqs.update_traces(
            hovertemplate="<b>%{fullData.name}</b><br>P = %{x:.1f} kW<br>Q = %{y:.1f} kVAR<extra></extra>",
        )
        fig_pqs.update_layout(
    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="#0a1520",
    height=380,
    margin=dict(t=40, b=10, l=10, r=200),  # r=200 pour la légende à droite
    xaxis_gridcolor="#1e3048", yaxis_gridcolor="#1e3048",
    showlegend=True,
    legend=dict(
        title=dict(text="Équipements", font=dict(size=11, color="#c8d6e5")),
        orientation="v",       # verticale à droite
        x=1.02,                # juste à droite du graphique
        xanchor="left",
        y=1.0,
        yanchor="top",
        font=dict(size=10, color="#c8d6e5"),
        bgcolor="rgba(15,25,35,0.85)",
        bordercolor="#1e3048",
        borderwidth=1,
    ),
    )
        st.plotly_chart(fig_pqs, use_container_width=True)

    # ── [L 338] Tableau récapitulatif dépliable ───────────────────────────────
    # → expanded=True pour l'ouvrir par défaut
    with st.expander("📋 Table of Detailed Values", expanded=False):
        cols_show = [c for c in [
            "equipement", "Pn_kVA", "P_i1", "P_i",
            "Q_i1", "Q_i", "S_i1", "S_i",
            "Time_i1", "Time_i", "Tc_i1", "Tc_i"
        ] if c in df.columns]
        df_disp = df[cols_show].rename(columns={
            "equipement": "Équipement", "Pn_kVA": "Pn/Sn",
            "P_i1": "P J i-1", "P_i": "P J i",
            "Q_i1": "Q J i-1", "Q_i": "Q J i",
            "S_i1": "S J i-1", "S_i": "S J i",
            "Time_i1": "Heure i-1", "Time_i": "Heure i",
            "Tc_i1": "Tc% J i-1", "Tc_i": "Tc% J i",
        })
        num_cols = [c for c in df_disp.columns if df_disp[c].dtype in [float, "float64"]]
        st.dataframe(
            df_disp.style.format({c: "{:.2f}" for c in num_cols}, na_rep="—"),
            use_container_width=True, height=250,
        )


# ══════════════════════════════════════════════════════════════════════════════
# [L 355] RENDER HISTORIQUE — courbes temporelles P/Q/S
# ══════════════════════════════════════════════════════════════════════════════

def render_historique_panel(panel_info: dict, color_idx: int = 0):
    """
    Affiche un sous-panel HISTORIQUE :
      - Multiselect équipements
      - Radio indicateur (P / Q / S)
      - Courbe temporelle ou barres groupées si pas de datetime
      - Statistiques (moyenne, max) par équipement / A SUPPRIMER SI NON PERTINENT
      - Statistiques (P_i, max) par équipement
      - Tableau données brutes dépliable
    """
    name = panel_info["panel"]
    df   = panel_info["data"].copy()

    if df.empty:
        st.info(f"Aucune donnée historique pour : {name}")
        return

    st.markdown(f'<div class="panel-header">📈 {name}</div>', unsafe_allow_html=True)

    # Sélecteur d'équipements à afficher
    # → Par défaut : les 5 premiers ; modifier min(5, ...) pour en afficher plus
    equipements = sorted(df["equipement"].unique())
    eq_sel = st.multiselect(
        f"Équipements — {name}",
        equipements,
        default=equipements[:min(5, len(equipements))],
        key=f"eq_sel_{name}_{color_idx}",   # clé unique pour éviter les conflits entre panels
    )
    df_f = df[df["equipement"].isin(eq_sel)] if eq_sel else df

    col_l, col_r = st.columns([3, 1])

    with col_l:
        # Sélecteur d'indicateur à tracer
        indicateur = st.radio(
            "Indicateur",
            ["P_kW", "Q_kVar", "S_kVA"],
            format_func=lambda x: {
                "P_kW":   "P active (kW)",
                "Q_kVar": "Q réactive (kVAR)",
                "S_kVA":  "S apparente (kVA)",
            }[x],
            horizontal=True,
            key=f"ind_{name}_{color_idx}",
        )

        if "datetime" in df_f.columns and df_f["datetime"].notna().any():
            # ── Courbe temporelle si datetime disponible ──────────────────────
            df_plot = df_f.dropna(subset=["datetime", indicateur])
            fig = px.line(
                df_plot.sort_values("datetime"),
                x="datetime", y=indicateur, color="equipement",
                title=f"{name} — {indicateur}",
                template="plotly_dark",
                color_discrete_sequence=COLORS,
                labels={
                    "datetime": "Date/Heure",
                    indicateur: indicateur,
                    "equipement": "Équipement",
                },
            )
            fig.update_layout(
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="#0a1520",
                height=310, margin=dict(t=40, b=10, l=10, r=10),
                xaxis_gridcolor="#1e3048", yaxis_gridcolor="#1e3048",
                legend=dict(orientation="h", y=1.12),
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            # ── Barres groupées si pas de colonne datetime ────────────────────
            df_bar = df_f.groupby("equipement")[["P_kW", "Q_kVar", "S_kVA"]].mean().reset_index()
            fig = px.bar(
                df_bar, x="equipement", y=["P_kW", "Q_kVar", "S_kVA"],
                barmode="group", template="plotly_dark",
                color_discrete_sequence=["#3498db", "#e74c3c", "#f39c12"],
                title=f"{name} — Moyennes P / Q / S",
            )
            fig.update_layout(
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="#0a1520",
                height=310, xaxis_tickangle=-30,
            )
            st.plotly_chart(fig, use_container_width=True)

    with col_r:
        # Statistiques par équipement (moyenne + max)
        st.markdown("**Statistiques**")
        stats_df = (
            df_f.groupby("equipement")
                .agg(
                    P_i=("P_kW", "first"),
                    P_max=("P_kW", "max"),
                    Q_i=("Q_kVar", "first"),
                    Q_max=("Q_kVar", "max"),
                    S_i=("S_kVA", "first"),
                    S_max=("S_kVA", "max"),
                )
                .round(1)
        )

        def fmt_val(v, unit):
            try:
                if v is None or (isinstance(v, float) and pd.isna(v)):
                    return "—"
                return f"{v:.1f} {unit}"
            except Exception:
                return str(v)

        # Afficher les métriques selon l'indicateur sélectionné
        for eq, row in stats_df.head(6).iterrows():
            if indicateur == "P_kW":
                val = row.get('P_i')
                mx  = row.get('P_max')
                delta = (f"max {mx:.1f} kW" if mx not in (None, float('nan')) else None)
                st.metric(eq[:22], fmt_val(val, 'kW'), delta=delta)
            elif indicateur == "Q_kVar":
                val = row.get('Q_i')
                mx  = row.get('Q_max')
                delta = (f"max {mx:.1f} kVar" if mx not in (None, float('nan')) else None)
                st.metric(eq[:22], fmt_val(val, 'kVar'), delta=delta)
            else:
                val = row.get('S_i')
                mx  = row.get('S_max')
                delta = (f"max {mx:.1f} kVA" if mx not in (None, float('nan')) else None)
                st.metric(eq[:22], fmt_val(val, 'kVA'), delta=delta)

    # Tableau données brutes dépliable
    with st.expander("📋 Données brutes", expanded=False):
        fmt = {"P_kW": "{:.2f}", "Q_kVar": "{:.2f}", "S_kVA": "{:.2f}"}
        st.dataframe(
            df_f.sort_values("datetime", na_position="last")
                .style.format(fmt, na_rep="—"),
            use_container_width=True, height=220,
        )


# ══════════════════════════════════════════════════════════════════════════════
# [L 420] SIDEBAR — UPLOAD FICHIER
# ══════════════════════════════════════════════════════════════════════════════
# IMPORTANT : ce bloc doit rester AVANT le guard [L 438].
# → Pour accepter d'autres extensions : ajouter dans type=["xlsx","xlsm","xls"]

with st.sidebar:
    st.markdown("## ⚡ Monitoring of Electrical Parameters")
    st.markdown("---")
    st.markdown("**📁 Excel file**")
    uploaded = st.file_uploader(
        "Import the .xlsm / .xlsx file",
        type=["xlsx", "xlsm"],
        help="Seules les feuilles FOLLOWED... et Historique... seront traitées.",
    )


# ══════════════════════════════════════════════════════════════════════════════
# [L 438] GUARD — arrêt si aucun fichier importé
# ══════════════════════════════════════════════════════════════════════════════
# → Ce bloc DOIT être hors du with st.sidebar pour que st.stop() fonctionne.
# → Ne pas déplacer ces lignes à l'intérieur d'un bloc with.

if uploaded is None:
    st.title("⚡ CES⚡ Supervision of Electrical Equipment")
    st.info("<-- Import your Excel file into the sidebar to get started.")
    # st.footer("Compagnie d'Électricité du Senegal (CES)")
    # st.image("ImageCES.jpg", width=600)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.image("ImageCES.jpg", width=1200)
        st.markdown("### Compagnie d'Électricité du Senegal (CES)", unsafe_allow_html=True)
    st.stop()   # ← tout ce qui suit n'est JAMAIS exécuté tant qu'aucun fichier n'est chargé

# Lecture sécurisée : getvalue() est idempotent (peut être appelé plusieurs fois)
file_bytes = bytes(uploaded.getvalue())
data = load_excel(file_bytes)
followed_sheets   = list(data["followed"].keys())
historique_sheets = list(data["historique"].keys())
total_panels = (
    sum(len(v) for v in data["followed"].values())
  + sum(len(v) for v in data["historique"].values())
)


# ══════════════════════════════════════════════════════════════════════════════
# [L 445] SIDEBAR — NAVIGATION (visible seulement après chargement)
# ══════════════════════════════════════════════════════════════════════════════
# → Pour ajouter une vue, ajouter une entrée dans la liste du st.radio
#   et créer le bloc elif correspondant dans la section [L 463+]

with st.sidebar:
    st.success(
        f"✅ {len(followed_sheets)} feuilles FOLLOWED\n\n"
        f"✅ {len(historique_sheets)} feuilles HISTORIQUE\n\n"
        f"📊 {total_panels} sous-panels détectés"
    )
    st.markdown("---")
    st.markdown("**📂 Navigation**")
    mode = st.radio(
        "View Type",
        ["FOLLOWED (J i-1 / J i)", "HISTORIQUE (séries temporelles)", "Vue globale"],
        index=0,   # ← 0 = FOLLOWED sélectionné par défaut
    )
    if mode == "FOLLOWED (J i-1 / J i)":
        sheet_sel = st.selectbox("Feuille", followed_sheets) if followed_sheets else None
    elif mode == "HISTORIQUE (séries temporelles)":
        sheet_sel = st.selectbox("Feuille", historique_sheets) if historique_sheets else None
    else:
        sheet_sel = None
    # Affichage de l'heure actuelle (utile pour une supervision en temps réel)
    # st.markdown("---")
    # from datetime import datetime
    # now = datetime.now()
    # st.metric("🕐 Heure actuelle", now.strftime("%H:%M:%S"))


# ══════════════════════════════════════════════════════════════════════════════
# AFFICHAGE PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════

st.title("CES⚡ Supervision of Electrical Equipment ")


# ══════════════════════════════════════════════════════════════════════════════
# [L 463] VUE GLOBALE — synthèse de tous les panels en un seul écran
# ══════════════════════════════════════════════════════════════════════════════
# → Pour modifier le titre de la vue : changer st.subheader(...)
# → Pour retirer la synthèse HISTORIQUE : supprimer le bloc "Synthèse HISTORIQUE"

if mode == "Vue globale":
    st.subheader("📊 Overall Summary — All Panels")

    # Commentaire : ce bloc agrège les données de tous les sous-panels FOLLOWED pour afficher une synthèse globale.

#     # Agrégation : une ligne par sous-panel avec totaux P, S et taux de charge moyen
#     rows_glob = []
#     for sh_name, panels in data["followed"].items():
#         for p in panels:
#             df = p["data"].dropna(subset=["P_i"])
#             if df.empty:
#                 continue
#             rows_glob.append({
#                 "feuille":   sh_name,
#                 "panel":     p["panel"],
#                 "P_i_total": df["P_i"].sum(),
#                 "S_i_total": df["S_i"].sum() if "S_i" in df.columns else 0,
#                 "Tc_moy":    df["Tc_i"].mean() if "Tc_i" in df.columns and df["Tc_i"].notna().any() else None,
#                 "nb_eq":     len(df),
#             })

#     if rows_glob:
#         df_glob = pd.DataFrame(rows_glob)

#         # KPIs globaux
#         c1, c2, c3 = st.columns(3)
#         c1.metric("Total P active (tous panels)", f"{df_glob['P_i_total'].sum():.1f} kW")
#         c2.metric("Total S apparente",            f"{df_glob['S_i_total'].sum():.1f} kVA")
#         c3.metric("Nb équipements total",         int(df_glob["nb_eq"].sum()))

#         # Barres P par sous-panel
#         fig_glob = px.bar(
#             df_glob.sort_values("P_i_total", ascending=True),
#             x="P_i_total", y="panel", orientation="h",
#             color="feuille",
#             title="Puissance active totale par sous-panel (kW)",
#             template="plotly_dark",
#             color_discrete_sequence=COLORS,
#         )
#         fig_glob.update_layout(
#             paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="#0a1520",
#             height=max(350, len(df_glob) * 28),   # ← hauteur auto selon nb de panels
#             xaxis_gridcolor="#1e3048", yaxis_gridcolor="#1e3048",
#             margin=dict(l=10, r=10, t=40, b=10),
#         )
#         st.plotly_chart(fig_glob, use_container_width=True)

#         # Taux de charge par sous-panel
#         df_tc = df_glob.dropna(subset=["Tc_moy"])
#         if not df_tc.empty:
#             fig_tc = px.bar(
#                 df_tc.sort_values("Tc_moy"),
#                 x="Tc_moy", y="panel", orientation="h",
#                 color="Tc_moy",
#                 color_continuous_scale=["#27ae60", "#f39c12", "#e74c3c"],
#                 range_color=[0, 110],
#                 title="Taux de charge moyen par sous-panel (%)",
#                 template="plotly_dark",
#             )
#             fig_tc.add_vline(x=100, line_dash="dash", line_color="#e74c3c")
#             fig_tc.update_layout(
#                 paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="#0a1520",
#                 height=max(300, len(df_tc) * 28),
#                 coloraxis_showscale=False,
#                 margin=dict(l=10, r=10, t=40, b=10),
#             )
#             st.plotly_chart(fig_tc, use_container_width=True)
#     else:
#         st.info("Aucune donnée FOLLOWED disponible pour la synthèse globale.")

#     # Synthèse HISTORIQUE : moyenne P par équipement
#     st.markdown("---")
#     st.subheader("📈 Historique — Synthèse par feuille")
#     for sh_name, panels in data["historique"].items():
#         st.markdown(f"**{sh_name}**")
#         for p in panels[:3]:   # ← afficher max 3 sous-panels par feuille en vue globale
#             df = p["data"]
#             if df.empty:
#                 continue
#             cols = st.columns(min(4, len(df["equipement"].unique())))
#             for i, eq in enumerate(df["equipement"].unique()[:4]):
#                 d_eq = df[df["equipement"] == eq]
#                 cols[i].metric(eq[:20], f"{d_eq['P_kW'].mean():.1f} kW moy")


# # ══════════════════════════════════════════════════════════════════════════════
# # [L 510] VUE FOLLOWED — onglets par sous-panel
# # ══════════════════════════════════════════════════════════════════════════════
# # → Chaque sous-panel détecté dans la feuille devient un onglet st.tabs
# # → Si la feuille n'a qu'un seul sous-panel, pas d'onglets (affichage direct)

elif mode == "FOLLOWED (J i-1 / J i)":
    if not followed_sheets:
        st.warning("Aucune feuille FOLLOWED détectée dans ce fichier.")
        st.stop()
    if not sheet_sel:
        st.info("Sélectionne une feuille dans la barre latérale.")
        st.stop()

    st.subheader(f"📋 {sheet_sel}")
    panels = data["followed"][sheet_sel]

    if not panels:
        st.info("Aucun sous-panel trouvé dans cette feuille.")
    elif len(panels) == 1:
        # Un seul sous-panel → pas d'onglets
        render_followed_panel(panels[0])
    else:
        # Plusieurs sous-panels → un onglet par sous-panel
        tab_names = [p["panel"][:40] for p in panels]   # ← tronquer à 40 chars pour l'onglet
        tabs = st.tabs(tab_names)
        for tab, panel_info in zip(tabs, panels):
            with tab:
                render_followed_panel(panel_info)


# # ══════════════════════════════════════════════════════════════════════════════
# # [L 525] VUE HISTORIQUE — onglets par sous-panel
# # ══════════════════════════════════════════════════════════════════════════════
# # → Même logique que VUE FOLLOWED mais appelle render_historique_panel

elif mode == "HISTORIQUE (séries temporelles)":
    if not historique_sheets:
        st.warning("Aucune feuille HISTORIQUE détectée dans ce fichier.")
        st.stop()
    if not sheet_sel:
        st.info("Sélectionne une feuille dans la barre latérale.")
        st.stop()

    st.subheader(f"📈 {sheet_sel}")
    panels = data["historique"][sheet_sel]

    if not panels:
        st.info("Aucun sous-panel trouvé dans cette feuille.")
    elif len(panels) == 1:
        render_historique_panel(panels[0], 0)
    else:
        tab_names = [p["panel"][:40] for p in panels]
        tabs = st.tabs(tab_names)
        for i, (tab, panel_info) in enumerate(zip(tabs, panels)):
            with tab:
                render_historique_panel(panel_info, i)

# ══════════════════════════════════════════════════════════════════════════════
# RAFRAÎCHISSEMENT AUTOMATIQUE (pour mise à jour de l'heure)
# ══════════════════════════════════════════════════════════════════════════════
# import time
# time.sleep(1)
# st.rerun()
