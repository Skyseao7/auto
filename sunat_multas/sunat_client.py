from __future__ import annotations

import logging
import re
from dataclasses import fields
from datetime import date
from time import monotonic, sleep
from typing import Any

try:
    from playwright.sync_api import Error as PlaywrightError
    from playwright.sync_api import Page, TimeoutError as PlaywrightTimeoutError, sync_playwright
except ModuleNotFoundError:
    PlaywrightError = Exception
    PlaywrightTimeoutError = TimeoutError
    Page = Any
    sync_playwright = None

from .config import SunatConfig
from .errors import AuthenticationError, ExtractionError, NavigationError, NoRecordsFound
from .models import Company, ManifestRecord


LOGGER = logging.getLogger(__name__)


FIELD_ALIASES = {
    "operation": ("operacion", "operación", "tipo operacion", "tipo operación"),
    "manifest_type": ("tipo manifiesto", "tipo de manifiesto", "cod infraccion", "infracción", "multa"),
    "master": ("master", "documento master", "doc master"),
    "process_number": ("numero de proceso", "número de proceso", "proceso"),
    "manifest_number": ("numero de manifiesto", "número de manifiesto", "manifiesto"),
    "port_code": ("codigo de puerto", "código de puerto", "cod puerto"),
    "port": ("puerto", "terminal"),
    "cnt": ("cnt", "contenedor", "cantidad contenedores"),
    "numbering_type": ("tipo de numeracion", "tipo de numeración", "numeracion", "numeración"),
    "initial_numbering_datetime": ("fecha y hora de numeracion inicial", "numeración inicial"),
    "line_transmission_datetime": ("fecha y hora de transmision de la linea", "transmisión de la línea"),
    "complementary_info_datetime": ("fecha y hora de informacion complementaria", "información complementaria"),
    "vessel_arrival_datetime": ("fecha y hora de llegada de la nave", "llegada de la nave"),
    "cargo_agent_transmission_datetime": ("fecha y hora de transmision agte. de carga", "transmisión agte"),
    "boarding_end_datetime": ("fecha y hora de termino de embarque", "término de embarque"),
    "fine_status": ("estado", "estado multa", "estado (multa)", "multa"),
}


class SunatClient:
    def __init__(self, config: SunatConfig) -> None:
        self.config = config

    def login_only(self, company: Company, pause_seconds: int = 20) -> None:
        if sync_playwright is None:
            raise RuntimeError("Playwright no está instalado. Ejecuta: pip install -r requirements.txt")
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(headless=False, slow_mo=self.config.slow_mo_ms)
            context = browser.new_context()
            page = context.new_page()
            page.set_default_timeout(self.config.timeout_ms)
            try:
                self._login(page, company)
                LOGGER.info("Login enviado para %s. Manteniendo ventana %s segundos.", company.name, pause_seconds)
                sleep(pause_seconds)
            finally:
                context.close()
                browser.close()

    def navigate_only(
        self,
        company: Company,
        start_date: date,
        end_date: date,
        pause_seconds: int = 30,
    ) -> None:
        if sync_playwright is None:
            raise RuntimeError("Playwright no está instalado. Ejecuta: pip install -r requirements.txt")
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(headless=False, slow_mo=self.config.slow_mo_ms)
            context = browser.new_context()
            page = context.new_page()
            page.set_default_timeout(self.config.timeout_ms)
            try:
                self._login(page, company)
                self._navigate_to_query(page)
                self._select_fecha_numeracion_desconsolidado(page)
                self._fill_date_range(page, start_date, end_date)
                LOGGER.info("Consulta Manifiesto Desconsolidado abierta. Manteniendo ventana %s segundos.", pause_seconds)
                sleep(pause_seconds)
            finally:
                context.close()
                browser.close()
    def fetch_records(self, company: Company, start_date: date, end_date: date) -> list[ManifestRecord]:
        if sync_playwright is None:
            raise RuntimeError("Playwright no está instalado. Ejecuta: pip install -r requirements.txt")
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(
                headless=self.config.headless,
                slow_mo=self.config.slow_mo_ms,
            )
            context = browser.new_context(accept_downloads=True)
            page = context.new_page()
            page.set_default_timeout(self.config.timeout_ms)
            try:
                self._login(page, company)
                self._navigate_to_query(page)
                records = self._query_and_extract(page, start_date, end_date)
                self._logout(page)
                return records
            finally:
                context.close()
                browser.close()

    def _login(self, page: Page, company: Company) -> None:
        selectors = self.config.selectors
        LOGGER.info("Abriendo SUNAT para RUC %s", company.ruc)
        page.goto(self.config.login_url, wait_until="domcontentloaded")
        page.wait_for_load_state("networkidle", timeout=15000)

        if selectors.ruc_mode_button:
            try:
                self._click_first_available(page, selectors.ruc_mode_button, timeout_ms=5000)
            except (PlaywrightError, PlaywrightTimeoutError):
                LOGGER.info("El modo RUC ya parece activo o no se encontró el botón RUC.")

        self._fill_login_value(page, selectors.ruc_input, company.ruc, "RUC")
        self._fill_login_value(page, selectors.user_input, company.user, "USUARIO")
        self._fill_login_value(page, selectors.password_input, company.password, "CLAVE")
        LOGGER.info(
            "Credenciales cargadas desde LISTA: RUC %s | Usuario %s | Clave %s",
            company.ruc,
            company.user,
            "*" * len(company.password),
        )
        self._click_first_available(page, selectors.login_button)
        page.wait_for_load_state("domcontentloaded")

        if self._text_visible(page, selectors.auth_error_text, timeout_ms=3000):
            raise AuthenticationError("SUNAT rechazó las credenciales o solicitó validación adicional.")

    def _navigate_to_query(self, page: Page) -> None:
        try:
            self._open_foreign_trade_operator_menu(page)
            self._click_menu_option(page, ("Manifiesto de Carga de Ingreso",), timeout_ms=15000)
            sleep(2)
            self._click_menu_option(page, ("Consultas",), timeout_ms=15000)
            sleep(2)
            self._click_menu_option(
                page,
                (
                    "Consulta Manifiesto Desconsolidado",
                    "Consulta del Manifiesto Desconsolidado",
                    "Consulta de Manifiesto Desconsolidado",
                    "Consulta Manifesto Desconsolidado",
                ),
                timeout_ms=15000,
            )
            LOGGER.info("Consulta Manifiesto Desconsolidado abierta; esperando 10 segundos.")
            sleep(10)
            page.wait_for_load_state("domcontentloaded")
        except (PlaywrightError, PlaywrightTimeoutError) as exc:
            LOGGER.error("Opciones de menú visibles: %s", self._visible_menu_options(page))
            raise NavigationError(
                "No se pudo abrir Consulta Manifiesto Desconsolidado desde el menú SUNAT. "
                "Verifica que la opción exista para este usuario."
            ) from exc

    def _open_foreign_trade_operator_menu(self, page: Page) -> None:
        if self._text_visible(page, "Manifiesto de Carga de Ingreso", timeout_ms=3000):
            return
        LOGGER.info("Activando el menú Operador de Comercio Exterior.")
        self._click_menu_option(page, ("Operador de Comercio Exterior",), timeout_ms=15000)
        sleep(2)

    def _query_and_extract(self, page: Page, start_date: date, end_date: date) -> list[ManifestRecord]:
        selectors = self.config.selectors
        self._select_fecha_numeracion_desconsolidado(page)
        self._fill_date_range(page, start_date, end_date)
        self._click_first_available(page, selectors.search_button)
        page.wait_for_load_state("networkidle")

        if self._text_visible(page, selectors.no_records_text, timeout_ms=3000):
            raise NoRecordsFound("La consulta no devolvió registros.")

        rows = self._extract_table_rows(page, selectors.result_table)
        records = [self._row_to_record(row) for row in rows]
        if not records:
            raise NoRecordsFound("La tabla de resultados no contiene registros.")
        return records

    def _select_fecha_numeracion_desconsolidado(self, page: Page) -> None:
        target = "Fecha de Numeración de Manifiesto Desconsolidado"
        for scope in _page_scopes(page):
            try:
                scope.locator("#tipoBusquedaFechaNumeracionDesconsolidado").check(timeout=1000)
                LOGGER.info("Filtro de fecha seleccionado por ID.")
                return
            except (PlaywrightError, PlaywrightTimeoutError):
                pass
        for scope in _page_scopes(page):
            try:
                scope.get_by_label(re.compile(r"Fecha de Numeración.*Desconsolidado", re.IGNORECASE)).check(timeout=1000)
                LOGGER.info("Filtro de fecha seleccionado por etiqueta.")
                return
            except (PlaywrightError, PlaywrightTimeoutError):
                pass
        for scope in _page_scopes(page):
            try:
                scope.get_by_text(re.compile(r"Fecha\s+de\s+Numeraci[oó]n\s+de\s+Manifiesto\s+Desconsolidado", re.IGNORECASE)).click(timeout=1000)
                LOGGER.info("Filtro de fecha seleccionado por texto.")
                return
            except (PlaywrightError, PlaywrightTimeoutError):
                pass
        for scope in _page_scopes(page):
            try:
                if scope.evaluate(_radio_by_text_script(), target):
                    LOGGER.info("Filtro de fecha seleccionado mediante el DOM.")
                    return
            except PlaywrightError:
                continue
        raise PlaywrightTimeoutError("No se pudo seleccionar Fecha de Numeración de Manifiesto Desconsolidado")

    def _fill_date_range(self, page: Page, start_date: date, end_date: date) -> None:
        selectors = self.config.selectors
        start_value = start_date.strftime("%d/%m/%Y")
        end_value = end_date.strftime("%d/%m/%Y")
        start_input = self._fill_first_available(page, selectors.date_from_input, start_value)
        end_input = self._fill_first_available(page, selectors.date_to_input, end_value)
        if start_input.input_value() != start_value or end_input.input_value() != end_value:
            raise NavigationError("SUNAT no conservó el rango de fechas ingresado.")
        LOGGER.info("Rango mensual cargado: %s a %s", start_value, end_value)
    def _logout(self, page: Page) -> None:
        try:
            self._click_first_available(page, self.config.selectors.logout_link, timeout_ms=5000)
        except (PlaywrightError, PlaywrightTimeoutError):
            LOGGER.warning("No se pudo cerrar sesión desde el enlace configurado.")

    def _extract_table_rows(self, page: Page, table_selector: str) -> list[dict[str, str]]:
        try:
            table = self._first_available_locator(page, table_selector, timeout_ms=15000)
            headers = [_clean_text(header) for header in table.locator("thead tr th").all_inner_texts()]
            if not headers:
                headers = [_clean_text(header) for header in table.locator("tr").first.locator("th,td").all_inner_texts()]
                row_locator = table.locator("tr").nth(1)
                start_index = 1
            else:
                row_locator = table.locator("tbody tr")
                start_index = 0

            rows: list[dict[str, str]] = []
            row_count = row_locator.count() if start_index == 0 else table.locator("tr").count() - 1
            for index in range(row_count):
                cells = (
                    row_locator.nth(index).locator("td").all_inner_texts()
                    if start_index == 0
                    else table.locator("tr").nth(index + 1).locator("td").all_inner_texts()
                )
                values = [_clean_text(cell) for cell in cells]
                if any(values):
                    rows.append(dict(zip(headers, values, strict=False)))
            return rows
        except (PlaywrightError, PlaywrightTimeoutError) as exc:
            raise ExtractionError("No se pudo leer la tabla de resultados de SUNAT.") from exc

    def _row_to_record(self, row: dict[str, str]) -> ManifestRecord:
        normalized_row = {_normalize_header(key): value for key, value in row.items()}
        values: dict[str, str] = {}
        for field in fields(ManifestRecord):
            if field.name == "raw":
                continue
            values[field.name] = _find_value(normalized_row, FIELD_ALIASES.get(field.name, (field.name,)))
        return ManifestRecord(raw=row, **values)

    def _fill_login_value(self, page: Page, selector_list: str, value: str, field_label: str) -> None:
        try:
            self._fill_first_available(page, selector_list, value)
            return
        except (PlaywrightError, PlaywrightTimeoutError):
            LOGGER.info("No se encontró %s con selectores; probando búsqueda alternativa.", field_label)

        if self._fill_dom_candidate(page, value, field_label):
            return
        raise PlaywrightTimeoutError(f"No se pudo llenar el campo {field_label}")

    def _fill_dom_candidate(self, page: Page, value: str, field_label: str) -> bool:
        js = """
        ({ value, fieldLabel }) => {
            const normalize = (text) => (text || '')
                .toLowerCase()
                .normalize('NFD')
                .replace(/[\u0300-\u036f]/g, '');
            const label = normalize(fieldLabel);
            const inputs = Array.from(document.querySelectorAll('input')).filter((input) => {
                const style = window.getComputedStyle(input);
                const rect = input.getBoundingClientRect();
                return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0 && !input.disabled && !input.readOnly;
            });
            const haystack = (input) => normalize([
                input.id,
                input.name,
                input.placeholder,
                input.autocomplete,
                input.getAttribute('aria-label'),
                input.getAttribute('formcontrolname')
            ].join(' '));
            let candidates = [];
            if (label === 'ruc') {
                candidates = inputs.filter((input) => haystack(input).includes('ruc') || input.maxLength === 11);
            } else if (label === 'usuario') {
                candidates = inputs.filter((input) => haystack(input).includes('usuario') || haystack(input).includes('user'));
            } else if (label === 'clave') {
                candidates = inputs.filter((input) => input.type === 'password' || haystack(input).includes('clave') || haystack(input).includes('password'));
            }
            const input = candidates[0];
            if (!input) return false;
            input.focus();
            input.value = value;
            input.dispatchEvent(new Event('input', { bubbles: true }));
            input.dispatchEvent(new Event('change', { bubbles: true }));
            input.blur();
            return true;
        }
        """
        for scope in _page_scopes(page):
            try:
                if scope.evaluate(js, {"value": value, "fieldLabel": field_label}):
                    return True
            except PlaywrightError:
                continue
        return False

    def _fill_first_available(self, page: Page, selector_list: str, value: str):
        locator = self._first_available_locator(page, selector_list)
        locator.fill(value)
        return locator

    def _click_first_available(self, page: Page, selector_list: str, timeout_ms: int | None = None) -> None:
        locator = self._first_available_locator(page, selector_list, timeout_ms=timeout_ms)
        locator.click()

    def _first_available_locator(self, page: Page, selector_list: str, timeout_ms: int | None = None):
        selectors = [selector.strip() for selector in selector_list.split("|") if selector.strip()]
        last_error: Exception | None = None
        per_selector_timeout = min(timeout_ms or self.config.timeout_ms, 1200)
        deadline_timeout = timeout_ms or self.config.timeout_ms
        started_ms = _monotonic_ms()
        while _monotonic_ms() - started_ms < deadline_timeout:
            for scope in _page_scopes(_active_page(page)):
                for selector in selectors:
                    try:
                        locator = scope.locator(selector).first
                        locator.wait_for(state="visible", timeout=per_selector_timeout)
                        return locator
                    except (PlaywrightError, PlaywrightTimeoutError) as exc:
                        last_error = exc
        raise PlaywrightTimeoutError(f"No se encontró selector: {selector_list}") from last_error

    def _click_menu_option(self, page: Page, labels: tuple[str, ...], timeout_ms: int) -> None:
        last_error: Exception | None = None
        deadline = _monotonic_ms() + timeout_ms
        while _monotonic_ms() < deadline:
            for scope in _page_scopes(_active_page(page)):
                for label in labels:
                    try:
                        exact_label = re.compile(rf"^\s*{re.escape(label)}\s*$", re.IGNORECASE)
                        menu_option = scope.locator("span.spanNivelDescripcion:visible", has_text=exact_label).first
                        menu_option.click(timeout=750)
                        return
                    except (PlaywrightError, PlaywrightTimeoutError) as exc:
                        last_error = exc
                    try:
                        scope.get_by_text(label, exact=True).click(timeout=750)
                        return
                    except (PlaywrightError, PlaywrightTimeoutError) as exc:
                        last_error = exc
        option_list = " / ".join(labels)
        raise PlaywrightTimeoutError(f"No se encontró la opción del menú: {option_list}") from last_error

    def _visible_menu_options(self, page: Page) -> list[str]:
        options: list[str] = []
        try:
            for scope in _page_scopes(_active_page(page)):
                menu_options = scope.locator("span.spanNivelDescripcion:visible").all_inner_texts()
                options.extend(_clean_text(option) for option in menu_options if _clean_text(option))
        except PlaywrightError:
            return options
        return options

    def _text_visible(self, page: Page, text_or_regex: str, timeout_ms: int) -> bool:
        if not text_or_regex:
            return False
        for scope in _page_scopes(page):
            try:
                scope.get_by_text(re.compile(text_or_regex, re.IGNORECASE)).wait_for(state="visible", timeout=timeout_ms)
                return True
            except (PlaywrightError, PlaywrightTimeoutError):
                continue
        return False


def _monotonic_ms() -> int:
    return int(monotonic() * 1000)

def _radio_by_text_script() -> str:
    return r"""
    (targetText) => {
        const normalize = (text) => (text || '')
            .toLowerCase()
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .replace(/\s+/g, ' ')
            .trim();
        const target = normalize(targetText);
        const radios = Array.from(document.querySelectorAll('input[type="radio"]'));
        for (const radio of radios) {
            const label = radio.id ? document.querySelector(`label[for="${radio.id}"]`) : null;
            const containers = [label, radio.parentElement, radio.closest('tr'), radio.closest('td'), radio.closest('div')].filter(Boolean);
            const text = normalize(containers.map((node) => node.innerText || node.textContent || '').join(' '));
            if (text.includes(target)) {
                radio.click();
                radio.dispatchEvent(new Event('change', { bubbles: true }));
                return true;
            }
        }
        const optionTexts = Array.from(document.querySelectorAll('td, div, span, label')).filter((node) => normalize(node.innerText || node.textContent).includes(target));
        for (const node of optionTexts) {
            const container = node.closest('tr') || node.parentElement;
            const radio = container ? container.querySelector('input[type="radio"]') : null;
            if (radio) {
                radio.click();
                radio.dispatchEvent(new Event('change', { bubbles: true }));
                return true;
            }
        }
        return false;
    }
    """

def _page_scopes(page: Page):
    scopes = [page]
    try:
        scopes.extend(frame for frame in page.frames if frame is not page.main_frame)
    except Exception:
        pass
    return scopes


def _active_page(page: Page) -> Page:
    open_pages = [candidate for candidate in page.context.pages if not candidate.is_closed()]
    if not open_pages:
        raise PlaywrightError("SUNAT cerró la pestaña de la sesión.")
    return open_pages[-1]


def _find_value(normalized_row: dict[str, str], aliases: tuple[str, ...]) -> str:
    normalized_aliases = [_normalize_header(alias) for alias in aliases]
    for alias in normalized_aliases:
        if alias in normalized_row:
            return normalized_row[alias]
    for key, value in normalized_row.items():
        if any(alias in key or key in alias for alias in normalized_aliases):
            return value
    return ""


def _clean_text(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def _normalize_header(value: str) -> str:
    value = _clean_text(value).lower()
    return (
        value.replace("á", "a")
        .replace("é", "e")
        .replace("í", "i")
        .replace("ó", "o")
        .replace("ú", "u")
        .replace("ñ", "n")
    )



