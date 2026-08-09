from __future__ import annotations

import argparse
import logging
from calendar import monthrange
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook

from .config import load_config
from .errors import AuthenticationError, ExtractionError, NavigationError, NoRecordsFound
from .excel_io import (
    aplicar_detalle_filas,
    append_manifest_sheet,
    cargar_mapa_puertos,
    create_company_workbook,
    escribir_hoja_transmisiones,
    load_procesados,
    process_manifiestos_excel,
    read_companies,
    read_proceso_destino,
    read_proceso_destino_grupos,
    read_transmisiones,
    safe_filename,
    save_procesados,
)
from .logging_setup import append_incident, configure_logging
from .models import CompanyResult, IncidentType
from .reportes import obtener_tipo
from .sunat_client import SunatClient


LOGGER = logging.getLogger(__name__)


def main() -> None:
    args = parse_args()
    config = load_config(args.config)
    configure_logging(config.log_dir)
    tipo = obtener_tipo(args.reporte)

    companies = filter_companies(
        read_companies(config.list_path),
        item=args.item,
        ruc=args.ruc,
        name=args.nombre,
    )
    LOGGER.info("Empresas por procesar: %s", len(companies))
    if not companies:
        raise ValueError("No se encontró ninguna empresa con los filtros indicados.")

    if args.excel or args.solo_excel:
        run_excel_only(config, companies, args.archivo, args.hoja, tipo)
        return

    if args.cerrar_sesiones:
        run_logout_all(config, companies)
        return

    if args.trazabilidad:
        run_trazabilidad(config, companies, tipo)
        return

    start_date, end_date = resolve_date_range(args)

    client = SunatClient(config.sunat, tipo)
    if args.solo_login or args.solo_navegar:
        if len(companies) != 1:
            raise ValueError("Para esta prueba usa --item, --ruc o --nombre hasta seleccionar una sola empresa.")
        if args.solo_login:
            client.login_only(companies[0], pause_seconds=args.pausa_login)
            LOGGER.info("Prueba de login finalizada.")
        else:
            client.navigate_only(companies[0], start_date, end_date, pause_seconds=args.pausa_login)
            LOGGER.info("Prueba de navegación finalizada.")
        return

    if args.consulta:
        run_consulta(config, companies, start_date, end_date, tipo)
        return
    if args.detalle:
        run_detalle(config, companies, start_date, end_date, tipo)
        return
    if args.todo:
        run_todo(config, companies, start_date, end_date, tipo)
        return

    results: list[CompanyResult] = []
    for company in companies:
        results.append(process_company(config, client, company, start_date, end_date, args.dry_run))

    successful = sum(1 for result in results if result.incident_type is None)
    incidents = len(results) - successful
    LOGGER.info("Proceso finalizado. Correctas: %s | Incidencias: %s", successful, incidents)


def run_consulta(config, companies, start_date, end_date, tipo) -> set[str]:
    client = SunatClient(config.sunat, tipo)
    with_data: set[str] = set()
    for company in companies:
        LOGGER.info("Consultando SUNAT para %s | RUC %s", company.name, company.ruc)
        try:
            output_path = create_company_workbook(
                config.template_path,
                config.output_dir,
                company,
                config.keep_existing_outputs,
            )
            rows = client.fetch_records(company, start_date, end_date)
            count = escribir_hoja_transmisiones(output_path, rows, tipo)
            LOGGER.info("Hoja %s actualizada para %s: %s transmisiones.", tipo.hoja_transmisiones, company.name, count)
            with_data.add(company.ruc)
        except NoRecordsFound as exc:
            LOGGER.info("Sin datos para %s: %s", company.name, exc)
            if not tipo.es_expo:
                record_incident(config.log_dir, company, None, IncidentType.NO_RECORDS, str(exc))
        except AuthenticationError as exc:
            record_incident(config.log_dir, company, None, IncidentType.AUTH, str(exc))
        except (NavigationError, ExtractionError) as exc:
            record_incident(config.log_dir, company, None, IncidentType.SCRAPING, str(exc))
        except Exception as exc:
            LOGGER.exception("Error no controlado consultando %s", company.name)
            record_incident(config.log_dir, company, None, IncidentType.UNKNOWN, str(exc))
    return with_data


def run_todo(config, companies, start_date, end_date, tipo) -> None:
    with_data = run_consulta(config, companies, start_date, end_date, tipo)
    companies_with_data = [company for company in companies if company.ruc in with_data]
    run_excel_only(config, companies_with_data, None, None, tipo)
    for company in companies_with_data:
        run_detalle(config, [company], start_date, end_date, tipo)


def run_detalle(config, companies, start_date, end_date, tipo) -> None:
    if len(companies) != 1:
        raise ValueError("Para --detalle o --todo selecciona una sola empresa con --item, --ruc o --nombre.")
    company = companies[0]
    workbook_path = config.output_dir / f"{safe_filename(company.name)}.xlsx"
    if not workbook_path.exists():
        raise ValueError(f"No existe el archivo: {workbook_path}")

    try:
        codes = read_transmisiones(workbook_path, tipo)
    except ValueError as exc:
        LOGGER.warning("No hay transmisiones para %s: %s", company.name, exc)
        return
    groups = _build_groups(codes)
    log_path = config.output_dir / f"{safe_filename(company.name)}_procesados.json"
    processed = load_procesados(log_path)
    pending = [group for group in groups if group["code"] not in processed]
    if not pending:
        LOGGER.info("Todas las transmisiones de %s ya fueron procesadas.", company.name)
        return
    LOGGER.info("Transmisiones por procesar para %s: %s de %s.", company.name, len(pending), len(groups))

    workbook = load_workbook(workbook_path)
    client = SunatClient(config.sunat, tipo)

    mapa_puertos: dict[str, str] = {}
    if config.puertos_json and config.puertos_json.exists():
        try:
            mapa_puertos = cargar_mapa_puertos(config.puertos_json)
            LOGGER.info("Mapa de puertos cargado: %s códigos.", len(mapa_puertos))
        except (OSError, ValueError) as exc:
            LOGGER.warning("No se pudo cargar el mapa de puertos '%s': %s", config.puertos_json, exc)
    else:
        LOGGER.warning("No se encontró el JSON de puertos '%s'; la columna PUERTO quedará vacía.", config.puertos_json)

    def on_group(group, data_list) -> None:
        updates = []
        for offset, data in enumerate(data_list):
            updates.append((group["grid_start"] + 2 + offset, data))
        aplicar_detalle_filas(workbook, updates, tipo, mapa_puertos)
        workbook.save(workbook_path)
        processed.add(group["code"])
        save_procesados(log_path, processed)
        LOGGER.info("Transmisión %s procesada (%s documento(s)).", group["code"], len(data_list))

    try:
        client.procesar_detalle(company, start_date, end_date, pending, on_group)
    finally:
        workbook.close()
    LOGGER.info("Detalle finalizado para %s. Procesadas: %s.", company.name, len(processed))


def run_trazabilidad(config, companies, tipo) -> None:
    if len(companies) != 1:
        raise ValueError("Para --trazabilidad selecciona una sola empresa con --item, --ruc o --nombre.")
    company = companies[0]
    workbook_path = config.output_dir / f"{safe_filename(company.name)}.xlsx"
    if not workbook_path.exists():
        raise ValueError(f"No existe el archivo: {workbook_path}")
    codigos = read_proceso_destino(workbook_path, tipo)
    if not codigos:
        LOGGER.warning("La columna PROCESO de %s está vacía para %s.", tipo.hoja_destino, company.name)
        return
    LOGGER.info("Códigos de trazabilidad para %s: %s", company.name, len(codigos))
    grupos = read_proceso_destino_grupos(workbook_path, tipo)
    LOGGER.info("Grupos únicos de trazabilidad para %s: %s", company.name, len(grupos))
    client = SunatClient(config.sunat, tipo, workbook_path=workbook_path)
    client.consultar_trazabilidad(company, grupos)


def _build_groups(codes: list[str]) -> list[dict]:
    groups: list[dict] = []
    for index, code in enumerate(codes):
        code = code.strip()
        if not code:
            continue
        if groups and groups[-1]["code"] == code:
            groups[-1]["count"] += 1
        else:
            groups.append({"code": code, "grid_start": index, "count": 1})
    return groups


def run_logout_all(config, companies) -> None:
    client = SunatClient(config.sunat)
    for company in companies:
        LOGGER.info("Cerrando sesión de %s | RUC %s", company.name, company.ruc)
        try:
            client.logout_all(company)
            LOGGER.info("Sesión cerrada para %s | RUC %s", company.name, company.ruc)
        except AuthenticationError as exc:
            record_incident(config.log_dir, company, None, IncidentType.AUTH, str(exc))
        except (NavigationError, ExtractionError) as exc:
            record_incident(config.log_dir, company, None, IncidentType.SCRAPING, str(exc))
        except Exception as exc:
            LOGGER.exception("Error no controlado cerrando sesión de %s", company.name)
            record_incident(config.log_dir, company, None, IncidentType.UNKNOWN, str(exc))


def run_excel_only(config, companies, archivo, hoja, tipo) -> None:
    source_title = hoja or tipo.hoja_transmisiones
    if archivo:
        file_path = Path(archivo)
        if not file_path.exists():
            raise ValueError(f"No existe el archivo: {file_path}")
        copied = process_manifiestos_excel(file_path, tipo, source_title)
        LOGGER.info("Manifiestos copiados a %s desde %r en %s: %s", tipo.hoja_destino, source_title, file_path, copied)
        return
    for company in companies:
        file_path = config.output_dir / f"{safe_filename(company.name)}.xlsx"
        if not file_path.exists():
            raise ValueError(f"No existe el archivo: {file_path}")
        copied = process_manifiestos_excel(file_path, tipo, source_title)
        LOGGER.info("Manifiestos copiados a %s desde %r en %s: %s", tipo.hoja_destino, source_title, file_path, copied)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Automatiza reportes de multas SUNAT por empresa.")
    parser.add_argument("--config", type=Path, default=Path("config.yaml"), help="Ruta del archivo de configuración.")
    parser.add_argument("--reporte", default="impo118", help="Tipo de reporte: impo118, impo235, expo118, expo235 (por defecto, impo118).")
    parser.add_argument("--desde", help="Fecha inicial de consulta en formato YYYY-MM-DD.")
    parser.add_argument("--hasta", help="Fecha final de consulta en formato YYYY-MM-DD.")
    parser.add_argument("--mes", type=int, help="Mes de la consulta (1 a 12). Si se omite, se solicita al iniciar.")
    parser.add_argument("--anio", type=int, default=date.today().year, help="Año de la consulta mensual.")
    parser.add_argument("--dry-run", action="store_true", help="Crea/copias libros sin ingresar a SUNAT.")
    parser.add_argument("--solo-login", action="store_true", help="Solo abre SUNAT, llena credenciales y envía login para una empresa.")
    parser.add_argument("--solo-navegar", action="store_true", help="Ingresa y abre Consulta Manifiesto Desconsolidado para verificar la navegación.")
    parser.add_argument("--solo-excel", action="store_true", help="Alias de --excel: solo copia los manifiestos de la hoja de la empresa a IMPO118 en un Excel existente.")
    parser.add_argument("--excel", action="store_true", help="Solo procesa Excel: copia los datos base de IMPO118-Transmisiones a IMPO118.")
    parser.add_argument("--consulta", action="store_true", help="Solo consulta SUNAT y crea/actualiza la hoja IMPO118-Transmisiones.")
    parser.add_argument("--detalle", action="store_true", help="Recorre cada transmisión en SUNAT y llena el detalle de documentos en IMPO118.")
    parser.add_argument("--trazabilidad", action="store_true", help="Consulta la Trazabilidad del Manifiesto de Carga para cada código de la columna PROCESO de la hoja destino.")
    parser.add_argument("--todo", action="store_true", help="Ejecuta todo el flujo: --consulta + --excel + --detalle.")
    parser.add_argument("--cerrar-sesiones", action="store_true", help="Cierra la sesión SUNAT de todas las empresas de la LISTA (login + Salir).")
    parser.add_argument("--archivo", type=Path, help="Ruta de un Excel existente para usar con --excel/--solo-excel.")
    parser.add_argument("--hoja", help="Nombre de la hoja de datos para usar con --excel/--solo-excel (por defecto, la hoja de transmisiones del reporte).")
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

        rows = client.fetch_records(company, start_date, end_date)
        inserted = append_manifest_sheet(output_path, rows, client.tipo)
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


def resolve_date_range(args: argparse.Namespace) -> tuple[date, date]:
    if args.desde or args.hasta:
        if not (args.desde and args.hasta):
            raise ValueError("Indica ambas fechas con --desde y --hasta, o usa la consulta mensual.")
        start_date = parse_date(args.desde)
        end_date = parse_date(args.hasta)
        if start_date > end_date:
            raise ValueError("--desde no puede ser posterior a --hasta")
        return start_date, end_date

    month = args.mes if args.mes is not None else prompt_month()
    if not 1 <= month <= 12:
        raise ValueError("--mes debe estar entre 1 y 12.")
    start_date = date(args.anio, month, 1)
    end_date = date(args.anio, month, monthrange(args.anio, month)[1])
    LOGGER.info("Consulta mensual seleccionada: %s a %s", start_date, end_date)
    return start_date, end_date


def prompt_month() -> int:
    while True:
        value = input("Mes de consulta (1-12): ").strip()
        try:
            month = int(value)
        except ValueError:
            print("Ingresa un número de mes entre 1 y 12.")
            continue
        if 1 <= month <= 12:
            return month
        print("Ingresa un número de mes entre 1 y 12.")


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



