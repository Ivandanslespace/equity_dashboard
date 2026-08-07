import pandas as pd

from dashboard_xml import PortfolioDashboard


def _dashboard_with_metadata():
    dashboard = PortfolioDashboard.__new__(PortfolioDashboard)
    dashboard.paths = {}
    dashboard.last_screen = pd.DataFrame(
        {
            "ISIN": ["AAA", "BBB", "CCC"],
            "Exchange Country Region": ["West Europe", "West Europe", "North America"],
            "ICB19 Supersector": ["Technology", "Health Care", "Technology"],
        }
    )
    return dashboard


def test_filtre_et_renormalise_les_series_du_ptf_et_du_benchmark():
    """Les filtres région/secteur sont appliqués séparément et par date."""
    dashboard = _dashboard_with_metadata()
    dashboard.fund_ts = pd.DataFrame(
        {
            "Date": pd.to_datetime(["2026-07-14", "2026-07-14", "2026-07-14"]),
            "ISIN": ["AAA", "BBB", "CCC"],
            "%ACTIF": [0.20, 0.30, 0.50],
        }
    )
    dashboard.fund = None
    dashboard.bench_ts = dashboard.fund_ts.copy()
    dashboard.bench_df = None

    assert dashboard._apply_position_filters(
        {"region": "West Europe"},
        {"sector": "Technology"},
    )

    assert dashboard.fund_ts["ISIN"].tolist() == ["AAA", "BBB"]
    assert dashboard.fund_ts["%ACTIF"].sum() == 1.0
    assert dashboard.bench_ts["ISIN"].tolist() == ["AAA", "CCC"]
    assert dashboard.bench_ts["%ACTIF"].sum() == 1.0


def test_filtre_et_renormalise_un_snapshot_excel():
    """Un snapshot statique est également renormalisé après filtrage."""
    dashboard = _dashboard_with_metadata()
    dashboard.fund_ts = None
    dashboard.fund = pd.DataFrame(
        {"ISIN": ["AAA", "BBB", "CCC"], "%ACTIF": [20.0, 30.0, 50.0]}
    )
    dashboard.bench_ts = None
    dashboard.bench_df = pd.DataFrame(
        {"ISIN": ["AAA", "BBB", "CCC"], "%ACTIF": [20.0, 30.0, 50.0]}
    )

    dashboard._apply_position_filters(
        {"sector": "Health Care"},
        {"region": "North America"},
    )

    assert dashboard.fund["ISIN"].tolist() == ["BBB"]
    assert dashboard.fund["%ACTIF"].tolist() == [1.0]
    assert dashboard.bench_df["ISIN"].tolist() == ["CCC"]
    assert dashboard.bench_df["%ACTIF"].tolist() == [1.0]
