import pandas as pd

from dashboard_xml import PortfolioDashboard


def _dashboard_with_dates(months):
    dates = pd.to_datetime(
        ["2023-07-31", "2024-01-31", "2024-07-31", "2025-07-31"]
    )
    positions = pd.DataFrame(
        {
            "Date": dates,
            "ISIN": ["AAA"] * len(dates),
            "%ACTIF": [1.0] * len(dates),
        }
    )
    dashboard = PortfolioDashboard.__new__(PortfolioDashboard)
    dashboard.bhb_months = months
    dashboard.fund_ts = positions.copy()
    dashboard.bench_ts = positions.copy()
    dashboard.fund = None
    dashboard.bench_df = None
    return dashboard


def test_la_fenetre_bhb_est_parametree_en_mois():
    """La date de début BHB suit la fenêtre configurée."""
    dashboard_12m = _dashboard_with_dates(12)
    dashboard_12m._align_snapshots()
    assert dashboard_12m.attrib_start == pd.Timestamp("2024-07-31")

    dashboard_24m = _dashboard_with_dates(24)
    dashboard_24m._align_snapshots()
    assert dashboard_24m.attrib_start == pd.Timestamp("2023-07-31")
