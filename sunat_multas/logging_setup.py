from __future__ import annotations

import logging
from pathlib import Path


def configure_logging(log_dir: Path) -> None:
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / "reporte_multas.log"
    incident_path = log_dir / "incidentes.csv"

    if not incident_path.exists():
        incident_path.write_text(
            "item,nombre,ruc,tipo,detalle\n",
            encoding="utf-8-sig",
        )

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler(),
        ],
    )


def append_incident(log_dir: Path, item: str, name: str, ruc: str, incident_type: str, detail: str) -> None:
    safe_detail = str(detail).replace('"', "''").replace("\r", " ").replace("\n", " ")
    line = f'"{item}","{name}","{ruc}","{incident_type}","{safe_detail}"\n'
    with (log_dir / "incidentes.csv").open("a", encoding="utf-8-sig") as incident_file:
        incident_file.write(line)
