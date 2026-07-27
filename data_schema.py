"""Définit le contrat de colonnes de la feuille DATA."""

from __future__ import annotations

import pandas as pd


DATA_SCHEMA = (
    "ISIN",
    "Name",
    " Benchmark ICB Supersector ",
    "Exchange Country Region",
    "Score Multiffacteur Tilt",
    "Score Multifacteur",
    "Score Multifacteur_LAST",
    "Score Multifacteur_DELTA",
    "Score Value",
    "Score Value_LAST",
    "Score Value_DELTA",
    "PE FY1",
    "PE FY1_LAST",
    "PE FY1_DELTA",
    "Price to Book FY1",
    "Price to Book FY1_LAST",
    "Price to Book FY1_DELTA",
    "Price to FreeCF FY1",
    "Price to FreeCF FY1_LAST",
    "Price to FreeCF FY1_DELTA",
    "EV to Ebit FY1",
    "EV to Ebit FY1_LAST",
    "EV to Ebit FY1_DELTA",
    "EV to Sales FY1",
    "EV to Sales FY1_LAST",
    "EV to Sales FY1_DELTA",
    "EV To EBITDA FY1",
    "EV To EBITDA FY1_LAST",
    "EV To EBITDA FY1_DELTA",
    "PE LTM",
    "PE LTM_LAST",
    "PE LTM_DELTA",
    "PB LTM",
    "PB LTM_LAST",
    "PB LTM_DELTA",
    "PFCF LTM",
    "PFCF LTM_LAST",
    "PFCF LTM_DELTA",
    "EV to Sales LTM",
    "EV to Sales LTM_LAST",
    "EV to Sales LTM_DELTA",
    "EV To EBITDA LTM",
    "EV To EBITDA LTM_LAST",
    "EV To EBITDA LTM_DELTA",
    "Score Growth",
    "Score Growth_LAST",
    "Score Growth_DELTA",
    "5Y_Hist EPS TrendStab",
    "5Y_Hist EPS TrendStab_LAST",
    "5Y_Hist EPS TrendStab_DELTA",
    "5Y_Hist GrossInc TrendStab",
    "5Y_Hist GrossInc TrendStab_LAST",
    "5Y_Hist GrossInc TrendStab_DELTA",
    "5Y_Hist Sales TrendStab",
    "5Y_Hist Sales TrendStab_LAST",
    "5Y_Hist Sales TrendStab_DELTA",
    "Score Quality",
    "Score Quality_LAST",
    "Score Quality_DELTA",
    "ROE avg FY0",
    "ROE avg FY0_LAST",
    "ROE avg FY0_DELTA",
    "Oper Margin",
    "Oper Margin_LAST",
    "Oper Margin_DELTA",
    "Asset TO exFIN",
    "Asset TO exFIN_LAST",
    "Asset TO exFIN_DELTA",
    "NetDebt to EBITDA exFIN",
    "NetDebt to EBITDA exFIN_LAST",
    "NetDebt to EBITDA exFIN_DELTA",
    "TIER1 Ratio FY0",
    "TIER1 Ratio FY0_LAST",
    "TIER1 Ratio FY0_DELTA",
    "ROTE avg FY1",
    "ROTE avg FY1_LAST",
    "ROTE avg FY1_DELTA",
    "Combined Ratio FY1",
    "Combined Ratio FY1_LAST",
    "Combined Ratio FY1_DELTA",
    "Score Momentum",
    "Score Momentum_LAST",
    "Score Momentum_DELTA",
    "PMOM 12M1M",
    "PMOM 12M1M_LAST",
    "PMOM 12M1M_DELTA",
    "EPS NTM 3M Growth",
    "EPS NTM 3M Growth_LAST",
    "EPS NTM 3M Growth_DELTA",
    "EPS Revision Ratio",
    "EPS Revision Ratio_LAST",
    "EPS Revision Ratio_DELTA",
    "Score Volatility",
    "Score Volatility_LAST",
    "Score Volatility_DELTA",
    "Daily Vol 60J",
    "Daily Vol 60J_LAST",
    "Daily Vol 60J_DELTA",
    "Daily Vol 90J",
    "Daily Vol 90J_LAST",
    "Daily Vol 90J_DELTA",
    "Daily Vol 260J",
    "Daily Vol 260J_LAST",
    "Daily Vol 260J_DELTA",
    "Score Dividend",
    "Score Dividend_LAST",
    "Score Dividend_DELTA",
    "DVD Yield FY0",
    "DVD Yield FY0_LAST",
    "DVD Yield FY0_DELTA",
    "DVD Yield FY1",
    "DVD Yield FY1_LAST",
    "DVD Yield FY1_DELTA",
    "DPS FY1",
    "DPS FY1_LAST",
    "DPS FY1_DELTA",
    "Earns Yield FY0",
    "Earns Yield FY0_LAST",
    "Earns Yield FY0_DELTA",
    "Earns Yield FY1",
    "Earns Yield FY1_LAST",
    "Earns Yield FY1_DELTA",
    "Score ML rebased",
    "Perf5D",
    "Perf1M",
    "Perf3M",
    "Perf6M",
    "Contrib Alpha",
    "Contrib TE",
    "Weight PTF",
    "Weight Bench",
    "Déviation Bench",
    "Beta",
    "Company SEDOL",
    "LIBELLE",
    "ICB19 Supersector",
    "Benchmark Market Value Millions in EUR ",
    "Reco Analyst",
    "Hors indice",
    "Score ML",
    "EBITDAm FY1",
    "Gross Margin",
    "DVD Payout FY0",
    "FCF Conversion",
    "CFO Div Cov Ratio",
    "FCF Div Cov Ratio",
    "Sales Growth FY1",
    "EBITDA Growth FY1 CIQ",
    "Gross Income Growth FY1",
    "EPS Growth FY1",
    "Sales_5Y_growth",
    "EPS_5Y_growth",
    "Pct_Short_Interest",
    "SP Price Target CIQ",
    "Exchange Country Name",
    "Date",
    "Nb de fonds maison détenant la pos",
    "Statut PTF",
    "Weight PTF brut",
    "Weight Bench brut",
)

DATA_REQUIRED_COLUMNS = ("ISIN",)

DATA_ALIASES = {
    "Benchmark ICB Supersector": " Benchmark ICB Supersector ",
    "Benchmark ICB Supersector ": " Benchmark ICB Supersector ",
    " Benchmark ICB Supersector": " Benchmark ICB Supersector ",
    "Benchmark Market Value Millions in EUR": (
        "Benchmark Market Value Millions in EUR "
    ),
}


def stabiliser_schema_data(df: pd.DataFrame) -> pd.DataFrame:
    """Normalise les alias et applique l'ordre contractuel de DATA."""
    resultat = df.copy()
    for source, cible in DATA_ALIASES.items():
        if source not in resultat.columns or source == cible:
            continue
        if cible in resultat.columns:
            resultat[cible] = resultat[cible].combine_first(resultat[source])
            resultat.drop(columns=source, inplace=True)
        else:
            resultat.rename(columns={source: cible}, inplace=True)

    colonnes_manquantes = [
        colonne
        for colonne in DATA_REQUIRED_COLUMNS
        if colonne not in resultat.columns
    ]
    if colonnes_manquantes:
        raise ValueError(
            "Colonnes obligatoires absentes de DATA : "
            + ", ".join(colonnes_manquantes)
        )

    return resultat.reindex(columns=DATA_SCHEMA)
