from __future__ import annotations

import re
import shutil
from pathlib import Path

from openpyxl import load_workbook

from .models import Company, ManifestRecord


REQUIRED_LIST_COLUMNS = {"ITEM", "NOMBRE", "RUC", "USUARIO", "CLAVE"}
TARGET_SHEETS = ("IMPO118", "IMPO235", "EXPO118", "EXPO235")


def safe_filename(value: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*]+', " ", value).strip()
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned[:180] or "empresa_sin_nombre"


def read_companies(list_path: Path) -> list[Company]:
    workbook = load_workbook(list_path, read_only=True, data_only=True)
    if "LISTA" not in workbook.sheetnames:
        raise ValueError(f"No existe la hoja LISTA en {list_path}")

    sheet = workbook["LISTA"]
    header_values = [str(cell.value).strip().upper() if cell.value is not None else "" for cell in sheet[1]]
    header_index = {name: index + 1 for index, name in enumerate(header_values) if name}
    missing = REQUIRED_LIST_COLUMNS - set(header_index)
    if missing:
        raise ValueError(f"Faltan columnas en LISTA: {', '.join(sorted(missing))}")

    companies: list[Company] = []
    seen_keys: set[tuple[str, str]] = set()
    for row_number in range(2, sheet.max_row + 1):
        name = _cell_text(sheet.cell(row_number, header_index["NOMBRE"]).value)
        ruc = _cell_text(sheet.cell(row_number, header_index["RUC"]).value)
        user = _cell_text(sheet.cell(row_number, header_index["USUARIO"]).value)
        password = _cell_text(sheet.cell(row_number, header_index["CLAVE"]).value)
        item = _cell_text(sheet.cell(row_number, header_index["ITEM"]).value)
        if not any([name, ruc, user, password]):
            continue
        if not all([name, ruc, user, password]):
            raise ValueError(f"Fila {row_number} incompleta en LISTA")
        key = (ruc, user)
        if key in seen_keys:
            raise ValueError(f"Empresa duplicada en LISTA: RUC {ruc}, usuario {user}")
        seen_keys.add(key)
        companies.append(Company(item=item, name=name, ruc=ruc, user=user, password=password))

    return companies


def create_company_workbook(
    template_path: Path,
    output_dir: Path,
    company: Company,
    keep_existing: bool,
) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{safe_filename(company.name)}.xlsx"
    if output_path.exists() and keep_existing:
        return output_path
    shutil.copy2(template_path, output_path)
    return output_path


def append_records(workbook_path: Path, records: list[ManifestRecord]) -> int:
    workbook = load_workbook(workbook_path)
    for sheet_name in TARGET_SHEETS:
        if sheet_name not in workbook.sheetnames:
            raise ValueError(f"La plantilla no contiene la hoja {sheet_name}")

    inserted = 0
    counters = {sheet_name: _next_sequence(workbook[sheet_name]) for sheet_name in TARGET_SHEETS}
    for record in records:
        sheet_name = classify_sheet(record)
        sheet = workbook[sheet_name]
        row = sheet.max_row + 1
        sequence = counters[sheet_name]
        counters[sheet_name] += 1

        if sheet_name.startswith("IMPO"):
            values = [
                sequence,
                record.master,
                record.process_number,
                record.manifest_number,
                record.port_code,
                record.port,
                record.cnt,
                record.numbering_type,
                record.initial_numbering_datetime,
                record.line_transmission_datetime,
                record.complementary_info_datetime,
                record.vessel_arrival_datetime,
                record.fine_status,
            ]
        else:
            values = [
                sequence,
                record.master,
                record.process_number,
                record.manifest_number,
                record.port_code,
                record.port,
                record.cargo_agent_transmission_datetime,
                record.boarding_end_datetime,
                record.fine_status,
            ]
        for column, value in enumerate(values, start=1):
            sheet.cell(row=row, column=column, value=value)
        inserted += 1

    workbook.save(workbook_path)
    return inserted


def classify_sheet(record: ManifestRecord) -> str:
    operation = _normalize(record.operation)
    manifest_type = _normalize(record.manifest_type)
    if "EXPO" in operation or "EXPORT" in operation or "SALIDA" in operation:
        prefix = "EXPO"
    else:
        prefix = "IMPO"

    if "235" in manifest_type:
        suffix = "235"
    elif "118" in manifest_type:
        suffix = "118"
    else:
        raise ValueError(f"No se pudo clasificar manifiesto: {record.raw or record}")

    return f"{prefix}{suffix}"


def _next_sequence(sheet) -> int:
    for row in range(sheet.max_row, 1, -1):
        value = sheet.cell(row, 1).value
        if isinstance(value, int):
            return value + 1
        if isinstance(value, str) and value.strip().isdigit():
            return int(value.strip()) + 1
    return 1


def _cell_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _normalize(value: str) -> str:
    return (value or "").strip().upper()
