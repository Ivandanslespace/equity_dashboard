import pandas as pd
import pytest

from dashboard_xml import PortfolioDashboard


def test_screen_ts_cherche_automatiquement_la_colonne_du_benchmark(tmp_path):
    """Le nom du fonds permet de sélectionner automatiquement sa colonne de poids."""
    path = tmp_path / "screen.parquet"
    source = pd.DataFrame(
        {
            "Date": pd.to_datetime(["2025-01-31", "2025-01-31", "2025-02-28"]),
            "Name": ["A", "B", "A"],
            "Weight in MSCI WORLD": [0.25, 0.75, 1.0],
            "Weight in MSCI ACWI": [0.50, 0.50, 1.0],
        },
        index=pd.Index(["AAA", "BBB", "AAA"], name="ISIN"),
    )
    source.to_parquet(path)

    dashboard = PortfolioDashboard.__new__(PortfolioDashboard)
    result = dashboard._load_screen_bench_ts(path, "MSCI WORLD")

    assert list(result.columns) == ["Date", "ISIN", "LIBELLE", "%ACTIF"]
    assert result["ISIN"].tolist() == ["AAA", "BBB", "AAA"]
    assert result.groupby("Date")["%ACTIF"].sum().tolist() == [1.0, 1.0]


def test_screen_ts_signale_une_colonne_de_poids_absente(tmp_path):
    """Une erreur explicite est levée si le benchmark demandé n'existe pas."""
    path = tmp_path / "screen.parquet"
    source = pd.DataFrame(
        {"Date": pd.to_datetime(["2025-01-31"]), "Name": ["A"]},
        index=pd.Index(["AAA"], name="ISIN"),
    )
    source.to_parquet(path)

    dashboard = PortfolioDashboard.__new__(PortfolioDashboard)
    with pytest.raises(ValueError, match="Weight in MSCI WORLD"):
        dashboard._load_screen_bench_ts(path, "MSCI WORLD")


def test_screen_ts_charge_un_isin_stocke_comme_colonne(tmp_path):
    """Le format Parquet avec ISIN comme colonne est également accepté."""
    path = tmp_path / "screen_colonne.parquet"
    source = pd.DataFrame(
        {
            "Date": pd.to_datetime(["2025-01-31", "2025-01-31"]),
            "ISIN": ["AAA", "BBB"],
            "Name": ["A", "B"],
            "Weight in MSCI WORLD": [0.25, 0.75],
        }
    )
    source.to_parquet(path, index=False)

    dashboard = PortfolioDashboard.__new__(PortfolioDashboard)
    result = dashboard._load_screen_bench_ts(path, "MSCI WORLD")

    assert result["ISIN"].tolist() == ["AAA", "BBB"]
    assert result["%ACTIF"].sum() == 1.0
