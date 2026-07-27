"""Tests du contrat de colonnes de DATA."""

import pandas as pd
import pytest

from data_schema import DATA_SCHEMA, stabiliser_schema_data


def test_schema_data_complete_les_colonnes_optionnelles_et_fixe_l_ordre():
    """Un snapshot incomplet ne doit jamais déplacer les colonnes DATA."""
    source = pd.DataFrame({"ISIN": ["TEST"], "Beta": [1.1]})

    resultat = stabiliser_schema_data(source)

    assert tuple(resultat.columns) == DATA_SCHEMA
    assert resultat.columns.get_loc("Weight PTF") == 129
    assert resultat.columns.get_loc("Beta") == 132
    assert pd.isna(resultat.loc[0, "Reco Analyst"])


def test_schema_data_normalise_les_alias_de_colonnes():
    """Les variantes de libellé doivent alimenter la colonne canonique."""
    source = pd.DataFrame({
        "ISIN": ["TEST"],
        "Benchmark Market Value Millions in EUR": [125.0],
    })

    resultat = stabiliser_schema_data(source)

    assert resultat.loc[
        0, "Benchmark Market Value Millions in EUR "
    ] == 125.0


def test_schema_data_refuse_l_absence_d_isin():
    """ISIN reste la seule colonne obligatoire du contrat DATA."""
    with pytest.raises(ValueError, match="ISIN"):
        stabiliser_schema_data(pd.DataFrame({"Beta": [1.0]}))
