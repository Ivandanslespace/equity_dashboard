from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def test_une_seule_propriete_df_returns_par_implementation():
    """Les deux implémentations doivent partager une seule entrée de rendements."""
    for filename in ("dashboard.py", "dashboard_xml.py"):
        source = (ROOT / filename).read_text(encoding="utf-8")
        assert source.count("def df_returns(") == 1


def test_les_rendements_externes_ne_sont_plus_remplis_a_zero():
    """Les ETF et le CASH conservent leurs NaN pour le contrôle de couverture."""
    for filename in ("dashboard.py", "dashboard_xml.py"):
        source = (ROOT / filename).read_text(encoding="utf-8")
        assert "etf = etf.reindex(df.index).fillna(0.0)" not in source
        assert 'cash["CASH"] = pd.to_numeric(cash["CASH"], errors="coerce").fillna(0.0)' not in source


def test_le_risque_expose_la_date_de_reference():
    """Le calcul TE accepte une date as-of explicite et conserve les diagnostics."""
    for filename in ("dashboard.py", "dashboard_xml.py"):
        source = (ROOT / filename).read_text(encoding="utf-8")
        assert "as_of_date=None" in source
        assert "self.risk_diagnostics" in source
