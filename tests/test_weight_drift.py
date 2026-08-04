import pandas as pd

from dashboard_xml import PortfolioDashboard


def test_le_drift_des_poids_projette_un_snapshot_avec_les_rendements():
    """Les poids sont recalculés à la dernière date commune disponible."""
    dashboard = PortfolioDashboard.__new__(PortfolioDashboard)
    dashboard.fund_ts = None
    dashboard.fund = pd.DataFrame({
        "Date": [pd.Timestamp("2026-07-14")] * 2,
        "ISIN": ["AAA", "BBB"],
        "%ACTIF": [0.60, 0.40],
    })
    dashboard.analysis_as_of_date = pd.Timestamp("2026-07-14")
    dashboard.bench_ts = pd.DataFrame({
        "Date": [pd.Timestamp("2026-07-14"), pd.Timestamp("2026-07-24")],
        "ISIN": ["BENCH", "BENCH"],
        "%ACTIF": [1.0, 1.0],
    })
    dashboard.bench_df = dashboard.bench_ts.iloc[[0]].copy()
    dashboard.screen_agg = pd.DataFrame({
        "ISIN": ["AAA", "BBB"],
        "Company SEDOL": ["SEDOLA", "SEDOLB"],
    })
    dashboard.list_isin_etf = []
    dashboard._df_returns = pd.DataFrame(
        {
            "SEDOLA": [0.10, 0.00],
            "SEDOLB": [0.00, 0.20],
        },
        index=pd.to_datetime(["2026-07-15", "2026-07-24"]),
    )

    dashboard._apply_weight_drift()

    assert dashboard.analysis_as_of_date == pd.Timestamp("2026-07-24")
    assert dashboard.bench_df["Date"].unique().tolist() == [pd.Timestamp("2026-07-24")]
    assert dashboard.fund["%ACTIF"].tolist() == [0.66, 0.48]

