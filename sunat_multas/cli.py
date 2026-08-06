from __future__ import annotations

import argparse
import logging
from datetime import date, datetime
from pathlib import Path

from .config import load_config
from .errors import AuthenticationError, ExtractionError, NavigationError, NoRecordsFound
from .excel_io import append_records, create_company_workbook, read_companies
from .logging_setup import append_incident, configure_logging
from .models import CompanyResult, IncidentType
from .sunat_client import SunatClient


LOGGER = logging.getLogger(__name__)


def main() -> None:
    args = parse_args()
    config = load_config(args.config)
    configure_logging(config.log_dir)

    start_date = parse_date(args.desde)
    end_date = parse_date(args.hasta)
    if start_date > end_date:
        raise ValueError("--desde no puede ser posterior a --hasta")

    companies = filter_companies(
        read_companies(config.list_path),
        item=args.item,
        ruc=args.ruc,
        name=args.nombre,
    )
    LOGGER.info("Empresas por procesar: %s", len(companies))
    if not companies:
        raise ValueError("No se encontró ninguna empresa con los filtros indicados.")

    client = SunatClient(config.sunat)
    if args.solo_login or args.solo_navegar:
        if len(companies) != 1:
            raise ValueError("Para esta prueba usa --item, --ruc o --nombre hasta seleccionar una sola empresa.")
        if args.solo_login:
            client.login_only(companies[0], pause_seconds=args.pausa_login)
            LOGGER.info("Prueba de login finalizada.")
        else:
            client.navigate_only(companies[0], pause_seconds=args.pausa_login)
            LOGGER.info("Prueba de navegación finalizada.")
        return

    results: list[CompanyResult] = []
    for company in companies:
        results.append(process_company(config, client, company, start_date, end_date, args.dry_run))

    successful = sum(1 for result in results if result.incident_type is None)
    incidents = len(results) - successful
    LOGGER.info("Proceso finalizado. Correctas: %s | Incidencias: %s", successful, incidents)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Automatiza reportes de multas SUNAT por empresa.")
    parser.add_argument("--config", type=Path, default=Path("config.yaml"), help="Ruta del archivo de configuración.")
    parser.add_argument("--desde", required=True, help="Fecha inicial de consulta en formato YYYY-MM-DD.")
    parser.add_argument("--hasta", required=True, help="Fecha final de consulta en formato YYYY-MM-DD.")
    parser.add_argument("--dry-run", action="store_true", help="Crea/copias libros sin ingresar a SUNAT.")
    parser.add_argument("--solo-login", action="store_true", help="Solo abre SUNAT, llena credenciales y envía login para una empresa.")
    parser.add_argument("--solo-navegar", action="store_true", help="Ingresa y abre Consulta Manifiesto Desconsolidado para verificar la navegación.")
    parser.add_argument("--pausa-login", type=int, default=20, help="Segundos que mantiene abierta la ventana tras login en modo --solo-login.")
    parser.add_argument("--item", help="Procesa solo la empresa con este ITEM de la hoja LISTA.")
    parser.add_argument("--ruc", help="Procesa solo la empresa con este RUC de la hoja LISTA.")
    parser.add_argument("--nombre", help="Procesa solo empresas cuyo nombre contenga este texto.")
    return parser.parse_args()


def process_company(
    config,
    client: SunatClient,
    company,
    start_date: date,
    end_date: date,
    dry_run: bool,
) -> CompanyResult:
    LOGGER.info("Procesando %s | RUC %s", company.name, company.ruc)
    output_path: Path | None = None
    try:
        output_path = create_company_workbook(
            config.template_path,
            config.output_dir,
            company,
            config.keep_existing_outputs,
        )
        if dry_run:
            LOGGER.info("Dry-run: archivo preparado en %s", output_path)
            return CompanyResult(company, output_path, 0)

        records = client.fetch_records(company, start_date, end_date)
        inserted = append_records(output_path, records)
        LOGGER.info("Registros insertados para %s: %s", company.name, inserted)
        return CompanyResult(company, output_path, inserted)
    except AuthenticationError as exc:
        return record_incident(config.log_dir, company, output_path, IncidentType.AUTH, str(exc))
    except NoRecordsFound as exc:
        return record_incident(config.log_dir, company, output_path, IncidentType.NO_RECORDS, str(exc))
    except (NavigationError, ExtractionError) as exc:
        return record_incident(config.log_dir, company, output_path, IncidentType.SCRAPING, str(exc))
    except Exception as exc:
        LOGGER.exception("Error no controlado procesando %s", company.name)
        return record_incident(config.log_dir, company, output_path, IncidentType.UNKNOWN, str(exc))


def record_incident(
    log_dir: Path,
    company,
    output_path: Path | None,
    incident_type: IncidentType,
    detail: str,
) -> CompanyResult:
    append_incident(log_dir, company.item, company.name, company.ruc, incident_type.value, detail)
    LOGGER.warning("%s | %s | %s", company.name, incident_type.value, detail)
    return CompanyResult(company, output_path, 0, incident_type, detail)


def parse_date(value: str) -> date:
    return datetime.strptime(value, "%Y-%m-%d").date()


def filter_companies(companies, item: str | None, ruc: str | None, name: str | None):
    filtered = companies
    if item:
        filtered = [company for company in filtered if company.item == item.strip()]
    if ruc:
        normalized_ruc = ruc.strip()
        filtered = [company for company in filtered if company.ruc == normalized_ruc]
    if name:
        normalized_name = name.strip().casefold()
        filtered = [company for company in filtered if normalized_name in company.name.casefold()]
    return filtered



