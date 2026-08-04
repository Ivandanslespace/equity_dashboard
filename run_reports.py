"""Point d'entrée minimal pour générer les rapports Ellebore et PIT."""

from __future__ import annotations

import os
import sys
from pathlib import Path

if sys.platform.startswith("linux"):
    os.environ.setdefault("MPLBACKEND", "Agg")

from dashboard_xml import PortfolioDashboard


PROJECT_DIR = Path(__file__).resolve().parent


def _path_from_env(name: str, default: Path) -> Path:
    """Retourne un chemin configuré par variable d'environnement ou sa valeur par défaut."""
    return Path(os.getenv(name, str(default))).expanduser()


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
    try:
        from notification import sendNotification
    except ImportError:
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
        fund_config={"type": "excel_snap", "path": str(fund_path)},
        bench_config=bench_config,
        path_output=str(output_path),
        wb_input=str(template_path),
        returns=str(reference_paths["returns"]),
        ciq=str(reference_paths["ciq"]),
        transco=str(reference_paths["transco"]),
        transco_ISIN_Fonds=str(reference_paths["transco_isin_fonds"]),
        list_isin_etf=str(reference_paths["list_isin_etf"]),
    )


def main() -> int:
    """Génère les deux rapports et retourne un code de sortie adapté au planificateur."""
    messages: list[str] = []
    errors: list[str] = []

    if sys.platform.startswith("win"):
        default_position = PROJECT_DIR / "POSITION"
    else:
        default_position = Path("/POSITION")

    # Chaque répertoire ou fichier peut être remplacé indépendamment.
    position_dir = _path_from_env("POSITION_DIR", default_position)
    data_dir = _path_from_env("DASHBOARD_DATA_DIR", PROJECT_DIR / "data")
    output_dir = _path_from_env("DASHBOARD_OUTPUT_DIR", position_dir / "rapport")
    template_path = _path_from_env(
        "DASHBOARD_TEMPLATE", PROJECT_DIR / "Analyse_MASK.xlsx"
    )
    reference_paths = {
        "returns": _path_from_env(
            "DASHBOARD_RETURNS_PATH", data_dir / "returns.parquet"
        ),
        "ciq": _path_from_env(
            "DASHBOARD_CIQ_PATH", data_dir / "screen_aggregate.parquet"
        ),
        "transco": _path_from_env(
            "DASHBOARD_TRANSCO_PATH", data_dir / "Transco_FactSet_ICB.xlsx"
        ),
        "transco_isin_fonds": _path_from_env(
            "DASHBOARD_TRANSCO_ISIN_FONDS_PATH",
            data_dir / "Transco_ISIN_Fonds.xlsx",
        ),
        "list_isin_etf": _path_from_env(
            "DASHBOARD_ETF_PATH", data_dir / "RETURN_SAVE_ETF.xlsx"
        ),
        "benchmark": _path_from_env(
            "DASHBOARD_BENCHMARK_PATH",
            data_dir / "df_merged_position_HISTO.parquet",
        ),
    }
    output_dir.mkdir(parents=True, exist_ok=True)

    reports = [
        (
            "Ellebore",
            {"type": "parquet_ts", "path": str(reference_paths["benchmark"]), "fonds_name": "EUROSTOXX50"},
            output_dir / "Analyse_ellebore.xlsx",
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
            output_dir / "Analyse_pit.xlsx",
        ),
    ]

    for name, bench_config, output_path in reports:
        try:
            folder = position_dir / name
            messages.append(f"Début {name}.")
            fund_path = _latest_excel(folder)
            messages.append(f"Fichier {name} : {fund_path.name}.")
            dashboard = _build_dashboard(
                fund_path=fund_path,
                bench_config=bench_config,
                output_path=output_path,
                reference_paths=reference_paths,
                template_path=template_path,
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
