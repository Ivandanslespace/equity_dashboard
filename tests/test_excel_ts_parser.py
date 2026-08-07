import pandas as pd
import pytest

from dashboard_xml import PortfolioDashboard as DashboardXml


@pytest.mark.parametrize("weight_header", [" Weight ", " %ACTIF "])
def test_excel_ts_reconnait_les_colonnes_par_nom_et_sans_valorisation(
    tmp_path, weight_header
):
    """Les colonnes TS peuvent être dans n'importe quel ordre et la VL est optionnelle."""
    path = tmp_path / "positions.xlsx"
    pd.DataFrame(
        {
            "Commentaire": ["A", "B"],
            weight_header: [0.60, 0.40],
            "ISIN": ["AAA", "BBB"],
            " Date ": ["2026-07-14", "2026-07-14"],
        }
    ).to_excel(path, index=False)

    dashboard = DashboardXml.__new__(DashboardXml)
    dashboard._parse_fund_data({"type": "excel_ts", "path": str(path)})

    assert list(dashboard.fund_ts.columns) == [
        "Date",
        "ISIN",
        "%ACTIF",
        "Valorisation",
    ]
    assert dashboard.fund_ts["%ACTIF"].tolist() == [0.60, 0.40]
    assert dashboard.fund_ts["Valorisation"].tolist() == [0.60, 0.40]
