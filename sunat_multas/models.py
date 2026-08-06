from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path


class IncidentType(str, Enum):
    AUTH = "AUTH"
    NO_RECORDS = "NO_RECORDS"
    SCRAPING = "SCRAPING"
    EXCEL = "EXCEL"
    UNKNOWN = "UNKNOWN"


@dataclass(frozen=True)
class Company:
    item: str
    name: str
    ruc: str
    user: str
    password: str


@dataclass(frozen=True)
class ManifestRecord:
    operation: str
    manifest_type: str
    master: str = ""
    process_number: str = ""
    manifest_number: str = ""
    port_code: str = ""
    port: str = ""
    cnt: str = ""
    numbering_type: str = ""
    initial_numbering_datetime: str = ""
    line_transmission_datetime: str = ""
    complementary_info_datetime: str = ""
    vessel_arrival_datetime: str = ""
    cargo_agent_transmission_datetime: str = ""
    boarding_end_datetime: str = ""
    fine_status: str = ""
    raw: dict[str, str] = field(default_factory=dict)


@dataclass(frozen=True)
class CompanyResult:
    company: Company
    output_path: Path | None
    records_count: int
    incident_type: IncidentType | None = None
    incident_detail: str = ""
