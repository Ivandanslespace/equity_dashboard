"""Point d'entrée minimal pour générer les rapports Ellebore et PIT."""

from __future__ import annotations

import os
import sys
from pathlib import Path

if sys.platform.startswith("linux"):
    os.environ.setdefault("MPLBACKEND", "Agg")
    try:
        from notification import sendNotification
    except ImportError:
        sendNotification = None
else:
    sendNotification = None

from dashboard_xml import PortfolioDashboard


PROJECT_DIR = Path(__file__).resolve().parent

# Configuration simple : modifier uniquement les chemins ci-dessous.
if sys.platform.startswith("win"):
    path_position = Path(r"C:\GoogleDrive\Coding\dashboard\POSITION")
    path_output = Path(r"C:\GoogleDrive\Coding\dashboard\POSITION\rapport")
    path_data = Path(r"C:\GoogleDrive\Coding\dashboard\data")
    path_template = Path(r"C:\GoogleDrive\Coding\dashboard\Analyse_MASK.xlsx")
    path_returns = Path(r"C:\GoogleDrive\Coding\dashboard\data\returns.parquet")
    path_ciq = Path(r"C:\GoogleDrive\Coding\dashboard\data\screen_aggregate.parquet")
    path_transco = Path(r"C:\GoogleDrive\Coding\dashboard\data\Transco_FactSet_ICB.xlsx")
    path_transco_isin_fonds = Path(r"C:\GoogleDrive\Coding\dashboard\data\Transco_ISIN_Fonds.xlsx")
    path_etf = Path(r"C:\GoogleDrive\Coding\dashboard\data\RETURN_SAVE_ETF.xlsx")
    path_benchmark = Path(r"C:\GoogleDrive\Coding\dashboard\data\df_merged_position_HISTO.parquet")
    path_news = Path(r"\\groupe-ufg.com\Commun\Prive\GestionAM\Ingenierie_Financiere\PROD\MAJ_news_factset_daily\0_DATA\Base_news_facset_BRUTE.parquet")
    path_news_scored = Path(r"\\groupe-ufg.com\Commun\Prive\GestionAM\Ingenierie_Financiere\PROD\MAJ_news_factset_daily\0_DATA\current_scored_news2.parquet")
    path_reco_facto = Path(r"\\groupe-ufg.com\Commun\Prive\GestionAM\Ingenierie_Financiere\PROD\_EQUITY\0_PTF_BLOOM\reco_secto_facto.xlsx")
    path_position_pickle = Path(r"\\groupe-ufg.com\Commun\Prive\GestionAM\Ingenierie_Financiere\PROD\_BASE\_BASE_PICKLE_HISTO\df_merged_position.pkl")
else:
    path_position = Path("/POSITION")
    path_output = Path("/POSITION/rapport")
    path_data = PROJECT_DIR / "data"
    path_template = PROJECT_DIR / "Analyse_MASK.xlsx"
    path_returns = Path("/usr/share/inge-fi/PROD/_EQUITY/0_RETURNS/returns.parquet")
    path_ciq = Path("/usr/share/inge-fi/PROD/_EQUITY/0_SCREEN_AGG/screen_aggregate.parquet")
    path_transco = PROJECT_DIR / "data" / "Transco_FactSet_ICB.xlsx"
    path_transco_isin_fonds = PROJECT_DIR / "data" / "Transco_ISIN_Fonds.xlsx"
    path_etf = PROJECT_DIR / "data" / "RETURN_SAVE_ETF.xlsx"
    path_benchmark = Path("/usr/share/inge-fi/PROD/_BASE/df_merged_position.parquet")
    path_news = Path("/usr/share/inge-fi/PROD/MAJ_news_factset_daily/0_DATA/Base_news_facset_BRUTE.parquet")
    path_news_scored = Path("/usr/share/inge-fi/PROD/MAJ_news_factset_daily/0_DATA/current_scored_news2.parquet")
    path_reco_facto = Path("/usr/share/inge-fi/PROD/_EQUITY/0_PTF_BLOOM/reco_secto_facto.xlsx")
    path_position_pickle = Path("/usr/share/inge-fi/PROD/_BASE/_BASE_PICKLE_HISTO/df_merged_position.pkl")

# Emplacements des deux portefeuilles et des deux fichiers produits.
path_ellebore = path_position / "Ellebore"
path_pit = path_position / "PIT"
path_report_ellebore = path_output / "Analyse_ellebore.xlsx"
path_report_pit = path_output / "Analyse_pit.xlsx"


def _latest_excel(folder: Path) -> Path:
    """Retourne le classeur Excel le plus récemment modifié dans un dossier."""
    files = [
        path
        for path in folder.iterdir()
        if path.is_file()
        and path.suffix.lower() in {".xlsx", ".xlsm", ".xls"}
        and not path.name.startswith("~$")
    ]
    if not files:
        raise FileNotFoundError(f"Aucun fichier Excel trouvé dans {folder}.")
    return max(files, key=lambda path: path.stat().st_mtime)


def _send_linux_notification(subject: str, body: str, messages: list[str]) -> None:
    """Envoie une notification si la fonction est disponible dans l'environnement Azure."""
    if sendNotification is None:
        messages.append("Notification non envoyée : module notification introuvable.")
        return

    try:
        sendNotification("jobprod", subject, body.replace("\n", "<br>"))
        messages.append("Notification envoyée.")
    except Exception as exc:  # pragma: no cover - dépend du service Azure externe
        messages.append(f"Échec de la notification : {exc}")


def _build_dashboard(
    fund_path: Path,
    bench_config: dict,
    output_path: Path,
    reference_paths: dict[str, Path],
    template_path: Path,
) -> PortfolioDashboard:
    """Construit un dashboard XML avec les chemins locaux du serveur."""
    return PortfolioDashboard(
        fund_config={
            "type": "excel_snap",
            "path": str(fund_path),
            "drift_weights": True,
        },
        bench_config=bench_config,
        path_output=str(output_path),
        wb_input=str(template_path),
        returns=str(reference_paths["returns"]),
        ciq=str(reference_paths["ciq"]),
        transco=str(reference_paths["transco"]),
        transco_ISIN_Fonds=str(reference_paths["transco_isin_fonds"]),
        list_isin_etf=str(reference_paths["list_isin_etf"]),
        news=str(reference_paths["news"]),
        news_scored=str(reference_paths["news_scored"]),
        reco_facto=str(reference_paths["reco_facto"]),
        position_pickle=str(reference_paths["position_pickle"]),
    )


def main() -> int:
    """Génère les deux rapports et retourne un code de sortie adapté au planificateur."""
    messages: list[str] = []
    errors: list[str] = []

    # Chaque répertoire ou fichier peut être remplacé indépendamment.
    output_dir = path_output
    output_dir.mkdir(parents=True, exist_ok=True)
    reference_paths = {
        "returns": path_returns,
        "ciq": path_ciq,
        "transco": path_transco,
        "transco_isin_fonds": path_transco_isin_fonds,
        "list_isin_etf": path_etf,
        "benchmark": path_benchmark,
        "news": path_news,
        "news_scored": path_news_scored,
        "reco_facto": path_reco_facto,
        "position_pickle": path_position_pickle,
    }

    reports = [
        (
            "Ellebore",
            {"type": "parquet_ts", "path": str(reference_paths["benchmark"]), "fonds_name": "EUROSTOXX50"},
            path_report_ellebore,
        ),
        (
            "PIT",
            {
                "type": "parquet_ts",
                "path": str(reference_paths["benchmark"]),
                "components": [
                    {"fonds_name": "MSCI ACWI", "weight": 90},
                    {"fonds_name": "CASH", "weight": 10},
                ],
            },
            path_report_pit,
        ),
    ]

    for name, bench_config, output_path in reports:
        try:
            folder = path_ellebore if name == "Ellebore" else path_pit
            messages.append(f"Début {name}.")
            fund_path = _latest_excel(folder)
            messages.append(f"Fichier {name} : {fund_path.name}.")
            dashboard = _build_dashboard(
                fund_path=fund_path,
                bench_config=bench_config,
                output_path=output_path,
                reference_paths=reference_paths,
                template_path=path_template,
            )
            messages.append(f"Données {name} chargées.")
            dashboard.export_to_excel()
            messages.append(f"Rapport {name} terminé : {output_path.name}.")
        except Exception as exc:
            error = f"ERREUR {name} : {exc}"
            messages.append(error)
            errors.append(error)

    summary = "\n".join(messages)
    print(summary)

    if sys.platform.startswith("linux"):
        status = "KO 2_generate_dashboard" if errors else "OK 2_generate_dashboard"
        prefix = "ERROR <br>" if errors else ""
        _send_linux_notification(status, prefix + summary, messages)

    if errors:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
