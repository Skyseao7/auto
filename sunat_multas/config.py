from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

try:
    import yaml
except ModuleNotFoundError:
    yaml = None


@dataclass(frozen=True)
class SunatSelectors:
    ruc_input: str
    user_input: str
    password_input: str
    login_button: str
    logout_link: str
    menu_search_input: str
    menu_option_text: str
    date_from_input: str
    date_to_input: str
    search_button: str
    result_table: str
    no_records_text: str
    auth_error_text: str
    ruc_mode_button: str = ""
    no_records_selector: str = ""
    session_open_text: str = ""
    session_open_confirm: str = ""


@dataclass(frozen=True)
class SunatConfig:
    login_url: str
    headless: bool
    slow_mo_ms: int
    timeout_ms: int
    selectors: SunatSelectors
    logout_manual: bool = True


@dataclass(frozen=True)
class AppConfig:
    list_path: Path
    template_path: Path
    output_dir: Path
    log_dir: Path
    keep_existing_outputs: bool
    sunat: SunatConfig


def load_config(path: Path) -> AppConfig:
    with path.open("r", encoding="utf-8") as config_file:
        text = config_file.read()
    raw = _load_yaml(text)

    sunat_raw: dict[str, Any] = raw.get("sunat", {})
    selectors_raw: dict[str, Any] = sunat_raw.get("selectors", {})
    selectors = SunatSelectors(**selectors_raw)
    sunat = SunatConfig(
        login_url=str(sunat_raw["login_url"]),
        headless=bool(sunat_raw.get("headless", False)),
        slow_mo_ms=int(sunat_raw.get("slow_mo_ms", 50)),
        timeout_ms=int(sunat_raw.get("timeout_ms", 30000)),
        logout_manual=bool(sunat_raw.get("logout_manual", True)),
        selectors=selectors,
    )
    return AppConfig(
        list_path=Path(raw.get("list_path", "lista_reporte_de_multas_sector_b.xlsx")),
        template_path=Path(raw.get("template_path", "Base.xlsx")),
        output_dir=Path(raw.get("output_dir", "output")),
        log_dir=Path(raw.get("log_dir", "logs")),
        keep_existing_outputs=bool(raw.get("keep_existing_outputs", True)),
        sunat=sunat,
    )


def _load_yaml(text: str) -> dict[str, Any]:
    if yaml is not None:
        return yaml.safe_load(text) or {}
    root: dict[str, Any] = {}
    stack: list[tuple[int, dict[str, Any]]] = [(-1, root)]
    for line in text.splitlines():
        if not line.strip() or line.lstrip().startswith("#"):
            continue
        indent = len(line) - len(line.lstrip(" "))
        key, separator, value = line.strip().partition(":")
        if not separator:
            continue
        while stack and indent <= stack[-1][0]:
            stack.pop()
        parent = stack[-1][1]
        if value.strip() == "":
            child: dict[str, Any] = {}
            parent[key] = child
            stack.append((indent, child))
        else:
            parent[key] = _parse_scalar(value.strip())
    return root


def _parse_scalar(value: str) -> Any:
    if value.lower() in {"true", "false"}:
        return value.lower() == "true"
    if value.startswith(("'", '"')) and value.endswith(("'", '"')):
        return value[1:-1]
    try:
        return int(value)
    except ValueError:
        return value
