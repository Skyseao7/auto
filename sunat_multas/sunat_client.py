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

from .config import SunatConfig, SunatSelectors
from .errors import AuthenticationError, ExtractionError, NavigationError, NoRecordsFound
from .excel_io import escribir_puerto_destino, leer_mapa_master_fila
from .models import Company, ManifestRecord
from .reportes import TIPOS, TipoReporte


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
    def __init__(self, config: SunatConfig, tipo: TipoReporte | None = None, workbook_path=None) -> None:
        self.config = config
        self.tipo = tipo or TIPOS["impo118"]
        self.workbook_path = workbook_path

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
                self._aplicar_pasos_formulario(page)
                self._fill_agente_carga_ruc(page, company.ruc)
                LOGGER.info("Consulta Manifiesto Desconsolidado abierta. Manteniendo ventana %s segundos.", pause_seconds)
                sleep(pause_seconds)
            finally:
                context.close()
                browser.close()
    def fetch_records(self, company: Company, start_date: date, end_date: date) -> list[list[str]]:
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
                records = self._query_and_extract(page, company.ruc, start_date, end_date)
                return records
            finally:
                self._logout(page)
                context.close()
                browser.close()

    def _login(self, page: Page, company: Company, wait_login: bool = True) -> None:
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
        if wait_login:
            self._click_first_available(page, selectors.login_button)
            page.wait_for_load_state("domcontentloaded")
            if self._text_visible(page, selectors.auth_error_text, timeout_ms=3000):
                raise AuthenticationError("SUNAT rechazó las credenciales o solicitó validación adicional.")
            self._dismiss_session_open(page)
        else:
            locator = self._first_available_locator(page, selectors.login_button)
            locator.click(no_wait_after=True)
            LOGGER.info("Login enviado sin esperar la carga completa del menú.")

    def _dismiss_session_open(self, page: Page) -> None:
        selectors = self.config.selectors
        if not selectors.session_open_text:
            return
        if not self._text_visible(page, selectors.session_open_text, timeout_ms=2000):
            return
        LOGGER.info("Aviso de sesión abierta en otro equipo detectado; confirmando el cierre de la sesión anterior.")
        confirm_selectors = selectors.session_open_confirm or "text=Continuar|text=Aceptar|text=OK|#btnAceptar"
        for _ in range(3):
            try:
                self._click_first_available(page, confirm_selectors, timeout_ms=3000)
                return
            except (PlaywrightError, PlaywrightTimeoutError):
                sleep(1)
        LOGGER.warning("No se pudo confirmar el aviso de sesión abierta en otro equipo.")

    def logout_all(self, company: Company) -> None:
        if sync_playwright is None:
            raise RuntimeError("Playwright no está instalado. Ejecuta: pip install -r requirements.txt")
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch(
                headless=self.config.headless,
                slow_mo=self.config.slow_mo_ms,
            )
            context = browser.new_context()
            page = context.new_page()
            page.set_default_timeout(self.config.timeout_ms)
            try:
                self._login(page, company, wait_login=False)
                self._logout(page)
            finally:
                context.close()
                browser.close()

    def _navigate_to_query(
        self,
        page: Page,
        ruta: tuple[tuple[str, ...], ...] | None = None,
        selector_formulario: str | None = None,
    ) -> None:
        ruta = ruta or self.tipo.ruta_menu
        selector_formulario = selector_formulario or self.tipo.selector_formulario
        try:
            self._wait_for_menu_ready(page)
            self._open_menu_tree(page, ruta)
            self._click_menu_option(page, ruta[-1], timeout_ms=15000)
            self._wait_for_query_form(page, selector_formulario=selector_formulario)
        except (PlaywrightError, PlaywrightTimeoutError) as exc:
            LOGGER.error("Opciones de menú visibles: %s", self._visible_menu_options(page))
            raise NavigationError(
                "No se pudo abrir la consulta desde el menú SUNAT. "
                "Verifica que la opción exista para este usuario."
            ) from exc

    def _navigate_to_trazabilidad(self, page: Page) -> None:
        if not self.tipo.ruta_menu_trazabilidad:
            raise NavigationError("Este reporte no tiene configurada la ruta de trazabilidad.")
        self._navigate_to_query(
            page,
            ruta=self.tipo.ruta_menu_trazabilidad,
            selector_formulario=self.tipo.selector_formulario_trazabilidad,
        )

    def _open_menu_tree(
        self,
        page: Page,
        ruta: tuple[tuple[str, ...], ...] | None = None,
    ) -> None:
        """Expande cada nivel del árbol SOL solo si su hijo aún no es visible.

        No depende de la posición de los paneles (izquierda o derecha) ni de que el
        texto exista en cualquier parte de la página: verifica los elementos del menú
        directamente (span.spanNivelDescripcion). Así funciona igual para empresas cuyo
        menú ya está expandido y para las que tienen los paneles del medio (p. ej.
        PRADIVO), donde el árbol queda desplazado a la derecha.
        """
        ruta = ruta or self.tipo.ruta_menu
        for index in range(len(ruta) - 1):
            child_labels = ruta[index + 1]
            if self._menu_option_visible(page, child_labels):
                continue
            LOGGER.info("Expandiendo el nivel del menú: %s", ruta[index][0])
            self._click_menu_option(page, ruta[index], timeout_ms=10000)
            self._wait_for_menu_option(page, child_labels, timeout_ms=8000)

    def _wait_for_menu_ready(self, page: Page, timeout_ms: int = 15000) -> None:
        """Espera a que el menú SOL esté cargado para no desperdiciar tiempo fijo."""
        deadline = _monotonic_ms() + timeout_ms
        while _monotonic_ms() < deadline:
            if self._menu_option_visible(page, self.tipo.ruta_menu[0]):
                return
            sleep(0.3)
        raise PlaywrightTimeoutError("No se cargó el menú SOL")

    def _wait_for_menu_option(self, page: Page, labels: tuple[str, ...], timeout_ms: int) -> None:
        deadline = _monotonic_ms() + timeout_ms
        while _monotonic_ms() < deadline:
            if self._menu_option_visible(page, labels):
                return
            sleep(0.2)
        LOGGER.warning("La opción del menú no apareció: %s", " / ".join(labels))

    def _wait_for_query_form(self, page: Page, timeout_ms: int = 30000, selector_formulario: str | None = None) -> None:
        """Espera el formulario de la consulta y vuelve apenas esté visible."""
        selector = selector_formulario or self.tipo.selector_formulario
        deadline = _monotonic_ms() + timeout_ms
        while _monotonic_ms() < deadline:
            for scope in _page_scopes(_active_page(page)):
                try:
                    if scope.locator(f"#{selector}:visible").count() > 0:
                        LOGGER.info("Formulario de consulta %s cargado.", self.tipo.nombre)
                        return
                except PlaywrightError:
                    continue
            sleep(0.2)
        raise PlaywrightTimeoutError(f"No se cargó el formulario de consulta {self.tipo.nombre}")

    def _menu_option_visible(self, page: Page, labels: tuple[str, ...]) -> bool:
        for scope in _page_scopes(_active_page(page)):
            for label in labels:
                try:
                    exact_label = re.compile(rf"^\s*{re.escape(label)}\s*$", re.IGNORECASE)
                    if scope.locator("span.spanNivelDescripcion:visible", has_text=exact_label).first.count() > 0:
                        return True
                except PlaywrightError:
                    continue
        return False

    def _click_consultar(self, page: Page) -> None:
        """Presiona el botón Consultar con los selectores del tipo, con fallback."""
        try:
            btn = self._first_available_locator(page, self.tipo.selector_btn_consultar, timeout_ms=3000)
            for intento in range(1, 4):
                try:
                    btn.click(timeout=1500)
                    LOGGER.info("Botón Consultar presionado (intento %s/3): %s.", intento, self.tipo.selector_btn_consultar)
                except (PlaywrightError, PlaywrightTimeoutError) as exc:
                    LOGGER.info("Intento %s/3 de Consultar falló: %s", intento, exc)
                if intento < 3:
                    sleep(2)
            sleep(3)
            return
        except (PlaywrightError, PlaywrightTimeoutError):
            pass
        deadline = _monotonic_ms() + 10000
        last_error: Exception | None = None
        while _monotonic_ms() < deadline:
            for scope in _page_scopes(_active_page(page)):
                try:
                    scope.get_by_text("Consultar", exact=True).first.click(timeout=750)
                    LOGGER.info("Botón Consultar presionado por texto.")
                    return
                except (PlaywrightError, PlaywrightTimeoutError) as exc:
                    last_error = exc
                try:
                    scope.locator("#accion5_label:visible").first.click(timeout=750)
                    LOGGER.info("Botón Consultar presionado por id accion5_label.")
                    return
                except (PlaywrightError, PlaywrightTimeoutError) as exc:
                    last_error = exc
            sleep(0.2)
        self._click_first_available(page, self.config.selectors.search_button)

    def _query_and_extract(self, page: Page, ruc: str, start_date: date, end_date: date) -> list[list[str]]:
        selectors = self.config.selectors
        self._select_fecha_numeracion_desconsolidado(page)
        self._fill_date_range(page, start_date, end_date)
        self._aplicar_pasos_formulario(page)
        self._fill_agente_carga_ruc(page, ruc)
        self._click_consultar(page)
        try:
            page.wait_for_load_state("networkidle", timeout=8000)
        except PlaywrightTimeoutError:
            LOGGER.info("La red de SUNAT no quedó inactiva; esperando los resultados directamente.")
        rows = self._wait_for_manifest_grid(page)
        if not rows:
            raise NoRecordsFound(self._no_records_message(page) or "La consulta no devolvió registros.")
        LOGGER.info("Manifiestos extraídos del grid: %s", len(rows))
        return rows

    def _wait_for_manifest_grid(self, page: Page, timeout_ms: int = 45000) -> list[list[str]]:
        selectors = self.config.selectors
        deadline = _monotonic_ms() + timeout_ms
        while _monotonic_ms() < deadline:
            if self._grid_has_no_records(page):
                return []
            if self._grid_has_rows(page):
                return self._extract_manifest_grid(page)
            if self._text_visible(page, selectors.no_records_text, timeout_ms=300):
                break
            sleep(0.4)
        return []

    def _grid_has_no_records(self, page: Page) -> bool:
        return self._no_records_message(page) is not None

    def _no_records_message(self, page: Page) -> str | None:
        pattern = re.compile(r"no existe|sin registros|no se encontraron|no existen registros", re.IGNORECASE)
        for selector in (self.tipo.selector_grid_vacio, self.config.selectors.no_records_selector):
            if not selector:
                continue
            for scope in _page_scopes(_active_page(page)):
                try:
                    node = scope.locator(selector).filter(has_text=pattern).first
                    if node.count() > 0 and node.is_visible():
                        message = _clean_text(node.inner_text())
                        if message:
                            return message
                except (PlaywrightError, PlaywrightTimeoutError):
                    continue
        return None

    def _grid_has_rows(self, page: Page) -> bool:
        scope = self._find_grid_scope(page)
        if scope is None:
            return False
        try:
            return scope.locator(f"#{self.tipo.selector_grid} {self.tipo.selector_grid_filas}").count() > 0
        except PlaywrightError:
            return False

    def _find_grid_scope(self, page: Page):
        selector = self.tipo.selector_grid
        for scope in _page_scopes(_active_page(page)):
            try:
                if scope.locator(f"#{selector}").count() > 0:
                    return scope
            except PlaywrightError:
                continue
        return None

    def _extract_manifest_grid(self, page: Page) -> list[list[str]]:
        scope = self._find_grid_scope(page)
        if scope is None:
            return []
        grid_selector = f"#{self.tipo.selector_grid}"
        max_cols = self.tipo.columnas_extraccion
        rows: list[list[str]] = []
        seen_first: set[tuple[str, ...]] = set()
        for _ in range(50):
            row_locator = scope.locator(f"{grid_selector} {self.tipo.selector_grid_filas}")
            count = row_locator.count()
            if count == 0:
                break
            page_rows: list[list[str]] = []
            for index in range(count):
                cells = [_clean_text(cell) for cell in row_locator.nth(index).locator("td").all_inner_texts()]
                if any(cells):
                    page_rows.append(cells[:max_cols])
            if not page_rows:
                break
            first_key = tuple(page_rows[0])
            if first_key in seen_first:
                break
            seen_first.add(first_key)
            rows.extend(page_rows)
            if not self._next_grid_page(scope):
                break
            sleep(1)
        return rows

    def _next_grid_page(self, scope) -> bool:
        try:
            next_button = scope.locator(self.tipo.selector_grid_siguiente).first
            if next_button.count() == 0:
                return False
            classes = (next_button.get_attribute("class") or "") + " " + (next_button.get_attribute("aria-disabled") or "")
            if "disabled" in classes.lower():
                return False
            next_button.click(timeout=3000)
            return True
        except (PlaywrightError, PlaywrightTimeoutError):
            return False

    def procesar_detalle(self, company: Company, start_date: date, end_date: date, groups, on_group) -> None:
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
                self._query_and_extract(page, company.ruc, start_date, end_date)
                for group in groups:
                    self._procesar_grupo_detalle(page, company.ruc, start_date, end_date, group, on_group)
            finally:
                self._logout(page)
                context.close()
                browser.close()

    def consultar_trazabilidad(self, company: Company, grupos: list[dict]) -> None:
        if sync_playwright is None:
            raise RuntimeError("Playwright no está instalado. Ejecuta: pip install -r requirements.txt")
        mapa_master_fila = leer_mapa_master_fila(self.workbook_path, self.tipo)
        LOGGER.info("Mapa de MASTER a fila cargado: %s entradas.", len(mapa_master_fila))
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
                self._navigate_to_trazabilidad(page)
                for index, grupo in enumerate(grupos, start=1):
                    codigo = grupo["code"]
                    self._llenar_consulta_trazabilidad(page, codigo)
                    LOGGER.info("Trazabilidad consultada (%s/%s) para código %s.", index, len(grupos), codigo)
                    self._abrir_detalle_trazabilidad(page, codigo, grupo["start_row"], grupo["count"], mapa_master_fila)
            finally:
                self._logout(page)
                context.close()
                browser.close()

    def _abrir_detalle_trazabilidad(self, page: Page, codigo: str, start_row: int, count: int, mapa_master_fila: dict) -> None:
        scope = self._abrir_tabla_detalle_consolidado(page, codigo)
        if scope is None:
            LOGGER.warning("No se abrió la tabla de detalle del documento (%s).", codigo)
            return
        try:
            total_filas = scope.locator("#tblLista tbody tr").count()
        except PlaywrightError:
            total_filas = count
        LOGGER.info("Manifiesto %s: %s documento(s) hijo(s) en la tabla de detalle.", codigo, total_filas)
        if total_filas > count:
            LOGGER.warning("La tabla de detalle tiene más filas que el grupo en Excel (%s).", codigo)
        for fila_offset in range(total_filas):
            self._procesar_hijo(page, scope, codigo, fila_offset, mapa_master_fila)
            if fila_offset < total_filas - 1:
                scope = self._abrir_tabla_detalle_consolidado(page, codigo)
                if scope is None:
                    LOGGER.warning("No se abrió la tabla de detalle para el siguiente hijo (%s).", codigo)
                    return

    def _abrir_tabla_detalle_consolidado(self, page: Page, codigo: str):
        selector = self.tipo.selector_tabla_trazabilidad
        deadline = _monotonic_ms() + 20000
        enlace = None
        while _monotonic_ms() < deadline:
            for scope in _page_scopes(_active_page(page)):
                try:
                    enlace = scope.locator(
                        f"#{selector} tbody {self.tipo.selector_enlace_detalle_trazabilidad}"
                    ).first
                    if enlace.count() > 0:
                        break
                except PlaywrightError:
                    continue
            if enlace is not None and enlace.count() > 0:
                break
            sleep(0.4)
        if enlace is None or enlace.count() == 0:
            LOGGER.warning("No se encontró el enlace de detalle en la tabla de trazabilidad (%s).", codigo)
            return None
        try:
            enlace.click(timeout=3000, no_wait_after=True)
            LOGGER.info("Enlace de detalle del Manifiesto Consolidado clicado (%s).", codigo)
        except (PlaywrightError, PlaywrightTimeoutError) as exc:
            LOGGER.warning("No se pudo clicar el enlace de detalle (%s): %s", codigo, exc)
            return None
        scope = self._esperar_tabla_en_cualquier_scope(page, "tblLista", timeout_ms=45000)
        if scope is None:
            LOGGER.warning("No se encontró la tabla de detalle del documento en la trazabilidad (%s).", codigo)
            return None
        return scope

    def _procesar_hijo(self, page: Page, scope, codigo: str, fila_offset: int, mapa_master_fila: dict) -> None:
        fila = self._fila_excel_por_master(scope, fila_offset, mapa_master_fila)
        if fila is None:
            LOGGER.warning("No se pudo determinar la fila de Excel para %s (fila %s).", codigo, fila_offset + 1)
            return
        enlace = scope.locator(f"#tblLista tbody tr:nth-child({fila_offset + 1}) a.link").first
        if enlace.count() == 0:
            LOGGER.warning("No se encontró el enlace del documento hijo en la tabla de detalle (%s).", codigo)
            return
        try:
            enlace.click(timeout=3000, no_wait_after=True)
            LOGGER.info("Enlace del documento hijo clicado (%s, fila Excel %s).", codigo, fila)
        except (PlaywrightError, PlaywrightTimeoutError) as exc:
            LOGGER.warning("No se pudo clicar el enlace del documento hijo (%s): %s", codigo, exc)
            return
        puerto = self._esperar_puerto_destino(page, timeout_ms=45000)
        if puerto is None:
            LOGGER.warning("No se encontró Puerto Destino (%s, fila %s).", codigo, fila)
            return
        LOGGER.info("Puerto Destino de %s: %s", codigo, puerto)
        try:
            escribir_puerto_destino(self.workbook_path, self.tipo, fila, puerto)
            LOGGER.info("Puerto Destino escrito en EXPO118 fila %s (E/F).", fila)
        except OSError as exc:
            LOGGER.warning("No se pudo escribir el puerto en %s: %s", self.workbook_path, exc)
            return
        self._clic_regresar(page, veces=2)

    def _fila_excel_por_master(self, scope, fila_offset: int, mapa_master_fila: dict) -> int | None:
        try:
            celda = scope.locator(
                f"#tblLista tbody tr:nth-child({fila_offset + 1}) td:nth-child(3)"
            ).first
            if celda.count() == 0:
                return None
            master_sistema = _clean_text(celda.inner_text(timeout=3000))
            if not master_sistema:
                return None
            LOGGER.info("Master del sistema (fila %s): %s", fila_offset + 1, master_sistema)
            return mapa_master_fila.get(master_sistema)
        except PlaywrightError:
            return None

    def _esperar_puerto_destino(self, page: Page, timeout_ms: int) -> str | None:
        deadline = _monotonic_ms() + timeout_ms
        while _monotonic_ms() < deadline:
            for candidate in _context_pages(page):
                for scope in _page_scopes(candidate):
                    try:
                        label = scope.locator(
                            'div.col-sm-3:has(label:has-text("Puerto Destino:"))'
                        )
                        if label.count() == 0:
                            continue
                        valor = label.first.locator("xpath=following-sibling::div[contains(@class, 'col-sm-3')][1]")
                        if valor.count() == 0:
                            continue
                        texto = _clean_text(valor.first.inner_text(timeout=3000))
                        if texto:
                            LOGGER.info("Puerto Destino localizado en %s.", scope)
                            return texto
                    except PlaywrightError:
                        continue
            sleep(0.4)
        return None

    def _clic_regresar(self, page: Page, veces: int = 1) -> None:
        for intento in range(veces):
            boton = self._esperar_boton_regresar(page, timeout_ms=15000)
            if boton is None:
                LOGGER.warning("No se encontró el botón Regresar (intento %s).", intento + 1)
                return
            try:
                boton.click(timeout=3000, no_wait_after=True)
                LOGGER.info("Botón Regresar clicado (intento %s).", intento + 1)
            except (PlaywrightError, PlaywrightTimeoutError) as exc:
                LOGGER.warning("No se pudo clicar Regresar (intento %s): %s", intento + 1, exc)
                return
            self._esperar_estabilidad_navegacion(page)
            sleep(5)

    def _esperar_estabilidad_navegacion(self, page: Page) -> None:
        try:
            activa = _active_page(page)
            activa.wait_for_load_state("domcontentloaded", timeout=10000)
        except (PlaywrightError, PlaywrightTimeoutError):
            pass

    def _esperar_boton_regresar(self, page: Page, timeout_ms: int):
        deadline = _monotonic_ms() + timeout_ms
        while _monotonic_ms() < deadline:
            for candidate in _context_pages(page):
                for scope in _page_scopes(candidate):
                    try:
                        boton = scope.locator('button:has-text("Regresar")').first
                        if boton.count() > 0:
                            return boton
                    except PlaywrightError:
                        continue
            sleep(0.4)
        return None

    def _esperar_tabla_en_cualquier_scope(self, page: Page, table_id: str, timeout_ms: int):
        deadline = _monotonic_ms() + timeout_ms
        while _monotonic_ms() < deadline:
            for candidate in _context_pages(page):
                for scope in _page_scopes(candidate):
                    try:
                        if scope.locator(f"#{table_id} tbody tr").count() > 0:
                            LOGGER.info("Tabla #%s localizada en %s.", table_id, scope)
                            return scope
                    except PlaywrightError:
                        continue
            sleep(0.4)
        return None

    def _llenar_consulta_trazabilidad(self, page: Page, codigo: str) -> None:
        self._expandir_parametros_busqueda(page)
        self._aplicar_pasos_formulario(page)
        self._seleccionar_radio_numero_manifiesto(page)
        input_numero = self._find_visible_quick(
            page,
            f"#{self.tipo.selector_input_numero_manifiesto}",
            timeout_ms=8000,
        )
        input_numero.fill(codigo)
        LOGGER.info("Número de Manifiesto Consolidado cargado: %s", codigo)
        try:
            btn = self._first_available_locator(
                page, self.tipo.selector_btn_consultar_trazabilidad, timeout_ms=3000
            )
            btn.click(timeout=3000)
            LOGGER.info("Botón Consultar trazabilidad presionado.")
            sleep(4)
        except (PlaywrightError, PlaywrightTimeoutError) as exc:
            LOGGER.warning("No se pudo presionar Consultar en la trazabilidad: %s", exc)

    def _expandir_parametros_busqueda(self, page: Page) -> None:
        if self._input_trazabilidad_visible(page):
            return
        deadline = _monotonic_ms() + 15000
        clicado = False
        while _monotonic_ms() < deadline:
            for scope in _page_scopes(_active_page(page)):
                try:
                    enlace = scope.locator('a[href="#collapseParametros"]').first
                    if enlace.count() > 0:
                        enlace.click(timeout=3000)
                        LOGGER.info("Enlace 'Parámetros Búsqueda' clicado.")
                        clicado = True
                        break
                except (PlaywrightError, PlaywrightTimeoutError):
                    continue
            if clicado:
                break
            sleep(0.4)
        if not clicado:
            LOGGER.warning("No se encontró el enlace 'Parámetros Búsqueda'.")
        deadline = _monotonic_ms() + 20000
        while _monotonic_ms() < deadline and not self._input_trazabilidad_visible(page):
            sleep(0.4)

    def _input_trazabilidad_visible(self, page: Page) -> bool:
        try:
            for scope in _page_scopes(_active_page(page)):
                locator = scope.locator(f"#{self.tipo.selector_input_numero_manifiesto}:visible").first
                if locator.count() > 0:
                    return True
            return False
        except PlaywrightError:
            return False

    def _seleccionar_radio_numero_manifiesto(self, page: Page) -> None:
        radio_id = self.tipo.selector_radio_numero_manifiesto
        for scope in _page_scopes(_active_page(page)):
            try:
                scope.locator(f"#{radio_id}").check(timeout=2000)
                LOGGER.info("Radio 'Número de Manifiesto Consolidado' marcado (%s).", radio_id)
                return
            except (PlaywrightError, PlaywrightTimeoutError):
                continue
        for scope in _page_scopes(_active_page(page)):
            try:
                scope.get_by_text("Número de Manifiesto Consolidado").first.click(timeout=2000)
                LOGGER.info("Radio 'Número de Manifiesto Consolidado' marcado por texto.")
                return
            except (PlaywrightError, PlaywrightTimeoutError):
                continue
        raise PlaywrightTimeoutError("No se pudo marcar el radio 'Número de Manifiesto Consolidado'.")

    def _procesar_grupo_detalle(self, page, ruc, start_date, end_date, group, on_group) -> None:
        LOGGER.info(
            "Procesando transmisión %s (fila grilla %s, %s documento(s)).",
            group["code"],
            group["grid_start"],
            group["count"],
        )
        self._click_manifiesto_link(page, group["grid_start"])
        if not self._esperar_detalle_transmision(page, group["code"], timeout_ms=15000):
            LOGGER.warning(
                "No se confirmó la apertura de la transmisión %s; reintentando.",
                group["code"],
            )
            self._regresar_a_grilla(page, ruc, start_date, end_date)
            self._click_manifiesto_link(page, group["grid_start"])
            self._esperar_detalle_transmision(page, group["code"], timeout_ms=15000)
        self._esperar_grid_documentos(page, timeout_ms=15000)
        sleep(1)
        listado = self._extraer_listado_documentos(page)
        if self.tipo.tiene_contenedores:
            contenedores = self._extraer_listado_contenedores(page)
        else:
            contenedores = []
        data_list: list[dict] = []
        for offset in range(group["count"]):
            if offset < len(listado):
                doc = listado[offset]
            else:
                LOGGER.warning(
                    "Transmisión %s: el listado trajo %s documento(s), se esperaban %s.",
                    group["code"],
                    len(listado),
                    group["count"],
                )
                doc = {}
            if self.tipo.tiene_contenedores:
                tipo_contenedor = contenedores[offset]["tipo_contenedor"] if offset < len(contenedores) else ""
                doc["cnt"] = "SI" if "CONTENEDOR" in tipo_contenedor.upper() else "NO"
            data_list.append(doc)
        on_group(group, data_list)
        self._regresar_a_grilla(page, ruc, start_date, end_date)

    def _click_manifiesto_link(self, page: Page, row_index: int) -> None:
        scope = self._find_grid_scope(page)
        if scope is None:
            raise NavigationError("No se encontró la grilla de manifiestos para abrir el detalle.")
        remaining = row_index
        for _ in range(100):
            rows = scope.locator("#gridManifiestoCarga .dojoxGridRow")
            count = rows.count()
            if remaining < count:
                row = rows.nth(remaining)
                link = row.locator("a.link").first
                if link.count() == 0:
                    link = row.locator("td").nth(2).locator("a").first
                if link.count() == 0:
                    link = row.locator("td").nth(2)
                link.scroll_into_view_if_needed(timeout=3000)
                link.click(timeout=5000)
                return
            if not self._next_grid_page(scope):
                break
            remaining -= count
            sleep(1)
        raise NavigationError(f"No se encontró la fila {row_index} en la grilla de manifiestos.")

    def _leer_codigo_transmision_detalle(self, page: Page) -> str:
        try:
            active = _active_page(page)
        except PlaywrightError:
            return ""
        for scope in _page_scopes(active):
            try:
                fila = scope.locator("tr", has_text=self.tipo.etiqueta_manifiesto).first
                celdas = fila.locator("td")
                for idx in range(celdas.count()):
                    texto = _clean_text(celdas.nth(idx).inner_text())
                    if re.fullmatch(self.tipo.codigo_regex, texto):
                        return texto
            except (PlaywrightError, PlaywrightTimeoutError):
                continue
        return ""

    def _esperar_detalle_transmision(self, page: Page, code: str, timeout_ms: int) -> bool:
        expected = code.strip().rsplit("-", 1)[-1].strip()
        deadline = _monotonic_ms() + timeout_ms
        while _monotonic_ms() < deadline:
            actual = self._leer_codigo_transmision_detalle(page)
            if actual and actual.rsplit("-", 1)[-1].strip() == expected:
                LOGGER.info("Transmisión confirmada en el detalle: %s.", actual)
                return True
            sleep(0.5)
        LOGGER.warning("El detalle no mostró la transmisión esperada %s.", code)
        return False

    def _esperar_grid_documentos(self, page: Page, timeout_ms: int) -> bool:
        deadline = _monotonic_ms() + timeout_ms
        while _monotonic_ms() < deadline:
            try:
                active = _active_page(page)
            except PlaywrightError:
                return False
            for scope in _page_scopes(active):
                try:
                    if scope.locator("#gridDocumentosTransporte .dojoxGridRow").count() > 0:
                        return True
                except (PlaywrightError, PlaywrightTimeoutError):
                    continue
            sleep(0.4)
        return False

    def _extraer_listado_documentos(self, page: Page) -> list[dict[str, str]]:
        rows = self._extraer_dojo_grid(page, "#gridDocumentosTransporte")
        if rows is None:
            raise ExtractionError("No se encontró LISTADO DE DOCUMENTOS DE TRANSPORTE.")
        result: list[dict[str, str]] = []
        for data in rows:
            doc = {
                "master": data.get(6, ""),
                "fecha_hijo": data.get(3, ""),
                "fecha_master": data.get(7, ""),
                "fecha_info": data.get(8, ""),
                "puerto_embarque": data.get(9, ""),
            }
            if any(doc.values()):
                result.append(doc)
        LOGGER.info("Listado de documentos extraído: %s documento(s).", len(result))
        return result

    def _extraer_listado_contenedores(self, page: Page) -> list[dict[str, str]]:
        rows = self._extraer_dojo_grid(page, "#gridEquipamientos")
        if not rows:
            LOGGER.info("LISTADO DE CONTENEDORES sin datos; se usará NO en la columna CTN.")
            return []
        result: list[dict[str, str]] = []
        for data in rows:
            contenedor = {"tipo_contenedor": data.get(3, "")}
            if any(contenedor.values()):
                result.append(contenedor)
        LOGGER.info("Listado de contenedores extraído: %s contenedor(es).", len(result))
        return result

    def _extraer_dojo_grid(self, page: Page, grid_selector: str) -> list[dict[int, str]] | None:
        scope = None
        for candidate in _page_scopes(_active_page(page)):
            try:
                if candidate.locator(grid_selector).count() > 0:
                    scope = candidate
                    break
            except PlaywrightError:
                continue
        if scope is None:
            return None
        try:
            scope.locator(grid_selector).scroll_into_view_if_needed(timeout=5000)
        except (PlaywrightError, PlaywrightTimeoutError):
            pass
        rows: list[dict[int, str]] = []
        seen_first: set[tuple] = set()
        for _ in range(50):
            grid_rows = scope.locator(f"{grid_selector} .dojoxGridRow")
            count = grid_rows.count()
            if count == 0:
                break
            page_rows: list[dict[int, str]] = []
            for index in range(count):
                data: dict[int, str] = {}
                cells = grid_rows.nth(index).locator("td.dojoxGridCell")
                cell_count = cells.count()
                for cell_index in range(cell_count):
                    idx = cells.nth(cell_index).get_attribute("idx")
                    if idx is not None:
                        text = _clean_text(cells.nth(cell_index).inner_text())
                        if text:
                            data[int(idx)] = text
                if data:
                    page_rows.append(data)
            if not page_rows:
                break
            first_key = tuple(sorted(page_rows[0].items()))
            if first_key in seen_first:
                break
            seen_first.add(first_key)
            rows.extend(page_rows)
            if not self._next_dojo_page(scope, grid_selector):
                break
            sleep(1)
        return rows

    def _next_dojo_page(self, scope, grid_selector: str) -> bool:
        try:
            next_button = scope.locator(f'{grid_selector} [title="Página siguiente"]').first
            if next_button.count() == 0:
                return False
            classes = (next_button.get_attribute("class") or "") + " " + (next_button.get_attribute("aria-disabled") or "")
            if "disable" in classes.lower():
                return False
            next_button.click(timeout=3000)
            return True
        except (PlaywrightError, PlaywrightTimeoutError):
            return False

    def _regresar_a_grilla(self, page: Page, ruc: str, start_date: date, end_date: date) -> None:
        try:
            active = _active_page(page)
            if active is not page:
                active.close()
        except PlaywrightError:
            pass

        for scope in _page_scopes(_active_page(page)):
            try:
                if scope.locator("#gridDocumentosTransporte").count() == 0:
                    continue
                regresar = scope.locator('[id="toolbar1.accion1_label"]').first
                if regresar.count() > 0:
                    LOGGER.info("Usando el botón Regresar del detalle.")
                    regresar.click(timeout=3000)
                    if self._wait_for_grid_rows(page, timeout_ms=15000):
                        return
            except (PlaywrightError, PlaywrightTimeoutError):
                continue

        LOGGER.warning("No se encontró el botón Regresar; usando el botón atrás del navegador.")
        try:
            page.go_back(wait_until="domcontentloaded")
        except (PlaywrightError, PlaywrightTimeoutError):
            LOGGER.warning("No se pudo volver con el botón atrás; re-consultando la grilla.")
        if self._wait_for_grid_rows(page, timeout_ms=15000):
            return
        self._query_and_extract(page, ruc, start_date, end_date)

    def _wait_for_grid_rows(self, page: Page, timeout_ms: int) -> bool:
        deadline = _monotonic_ms() + timeout_ms
        while _monotonic_ms() < deadline:
            if self._grid_has_rows(page):
                return True
            sleep(0.4)
        return False

    def _select_fecha_numeracion_desconsolidado(self, page: Page) -> None:
        radio_id = self.tipo.selector_radio_fecha
        target = self.tipo.radio_fecha_label
        for scope in _page_scopes(page):
            try:
                scope.locator(f"#{radio_id}").check(timeout=1000)
                LOGGER.info("Filtro de fecha seleccionado por ID (%s).", radio_id)
                return
            except (PlaywrightError, PlaywrightTimeoutError):
                pass
        for scope in _page_scopes(page):
            try:
                scope.get_by_label(re.compile(target, re.IGNORECASE)).check(timeout=1000)
                LOGGER.info("Filtro de fecha seleccionado por etiqueta.")
                return
            except (PlaywrightError, PlaywrightTimeoutError):
                pass
        for scope in _page_scopes(page):
            try:
                scope.get_by_text(re.compile(target, re.IGNORECASE)).click(timeout=1000)
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
        raise PlaywrightTimeoutError(f"No se pudo seleccionar {target}")

    def _fill_date_range(self, page: Page, start_date: date, end_date: date) -> None:
        start_value = start_date.strftime("%d/%m/%Y")
        end_value = end_date.strftime("%d/%m/%Y")
        start_input = self._fill_date_plain(page, self.tipo.date_from_input, start_value)
        end_input = self._fill_date_plain(page, self.tipo.date_to_input, end_value)
        if start_input.input_value() != start_value or end_input.input_value() != end_value:
            raise NavigationError("SUNAT no conservó el rango de fechas ingresado.")
        LOGGER.info("Rango mensual cargado: %s a %s", start_value, end_value)

    def _fill_date_plain(self, page: Page, selector_list: str, value: str) -> Any:
        locator = self._find_visible_quick(page, selector_list, timeout_ms=8000)
        locator.fill(value)
        return locator

    def _aplicar_pasos_formulario(self, page: Page) -> None:
        for paso in self.tipo.pasos_formulario:
            if paso.tipo_control == "select":
                self._seleccionar_select_html(page, paso.input_id, paso.texto)
            elif paso.tipo_control == "teclado":
                self._escribir_y_enter(page, paso.input_id, paso.texto)
            else:
                self._seleccionar_combobox(page, paso.input_id, paso.texto)

    def _escribir_y_enter(self, page: Page, input_id: str, texto: str, timeout_ms: int = 12000) -> None:
        deadline = _monotonic_ms() + timeout_ms
        last_error: Exception | None = None
        while _monotonic_ms() < deadline:
            try:
                active = _active_page(page)
            except PlaywrightError:
                return
            for scope in _page_scopes(active):
                try:
                    campo = scope.locator(f"#{input_id}").first
                    campo.click(timeout=2000)
                    campo.fill("")
                    campo.press_sequentially(texto, delay=60)
                    sleep(0.4)
                    items = scope.locator(f"#{input_id}_popup .dijitMenuItem:visible")
                    if items.count() == 0:
                        items = scope.locator(".dijitMenuItem:visible")
                    if items.count() == 0:
                        items = scope.locator("ul:visible li:visible")
                    clicked = False
                    if items.count() > 0:
                        for index in range(items.count()):
                            item_text = _clean_text(items.nth(index).inner_text())
                            if texto.casefold() in item_text.casefold():
                                items.nth(index).click(timeout=2500)
                                clicked = True
                                LOGGER.info("Opción %s clicada en %s: %s", input_id, texto, item_text)
                                break
                        if not clicked and items.count() == 1:
                            items.first.click(timeout=2500)
                            clicked = True
                    if not clicked:
                        campo.press("Enter")
                        sleep(0.3)
                    valor = campo.input_value()
                    if valor and texto.casefold() in valor.casefold():
                        LOGGER.info("Campo %s seleccionado: %s (valor %s)", input_id, texto, valor)
                        return
                    LOGGER.info("Campo %s quedó con valor: %r", input_id, valor)
                except (PlaywrightError, PlaywrightTimeoutError) as exc:
                    last_error = exc
            sleep(0.3)
        raise PlaywrightTimeoutError(f"No se pudo seleccionar el campo {input_id} = {texto}") from last_error

    def _seleccionar_select_html(self, page: Page, input_id: str, texto: str, timeout_ms: int = 12000) -> None:
        deadline = _monotonic_ms() + timeout_ms
        last_error: Exception | None = None
        while _monotonic_ms() < deadline:
            try:
                active = _active_page(page)
            except PlaywrightError:
                return
            for scope in _page_scopes(active):
                try:
                    select = scope.locator(f"#{input_id}").first
                    target_value = None
                    for idx in range(select.locator("option").count()):
                        opt = select.locator("option").nth(idx)
                        opt_text = _clean_text(opt.inner_text())
                        opt_value = opt.get_attribute("value") or ""
                        if opt_text == texto or opt_value == texto or texto in opt_text:
                            target_value = opt_value or opt_text
                            break
                    if target_value is None:
                        raise PlaywrightTimeoutError(
                            f"No se encontró la opción {texto!r} en el select #{input_id}"
                        )
                    select.select_option(value=target_value, timeout=2000)
                    selected = select.input_value()
                    if selected:
                        LOGGER.info("Select %s seleccionado: %s (%s)", input_id, texto, selected)
                        return
                except (PlaywrightError, PlaywrightTimeoutError) as exc:
                    last_error = exc
            sleep(0.3)
        raise PlaywrightTimeoutError(f"No se pudo seleccionar el select {input_id} = {texto}") from last_error

    def _seleccionar_combobox(self, page: Page, input_id: str, texto: str, timeout_ms: int = 12000) -> None:
        deadline = _monotonic_ms() + timeout_ms
        last_error: Exception | None = None
        while _monotonic_ms() < deadline:
            try:
                active = _active_page(page)
            except PlaywrightError:
                return
            for scope in _page_scopes(active):
                try:
                    combo = scope.locator(f"#{input_id}").first
                    combo.click(timeout=2000)
                    combo.fill("")
                    combo.press_sequentially(texto, delay=60)
                    sleep(0.4)
                    items = scope.locator(f"#{input_id}_popup .dijitMenuItem:visible")
                    if items.count() == 0:
                        items = scope.locator(".dijitMenuItem:visible")
                    target = None
                    for index in range(items.count()):
                        item_text = _clean_text(items.nth(index).inner_text())
                        if texto.casefold() in item_text.casefold():
                            target = items.nth(index)
                            break
                    if target is None and items.count() == 1:
                        target = items.first
                    if target is not None:
                        target.click(timeout=3000)
                        if texto.casefold() in combo.input_value().casefold():
                            LOGGER.info("Combobox %s seleccionado: %s", input_id, texto)
                            return
                except (PlaywrightError, PlaywrightTimeoutError) as exc:
                    last_error = exc
            sleep(0.3)
        raise PlaywrightTimeoutError(f"No se pudo seleccionar el combobox {input_id} = {texto}") from last_error

    def _find_visible_quick(self, page: Page, selector_list: str, timeout_ms: int) -> Any:
        """Busca el primer selector visible con sondeo rápido, sin bloquear por scope."""
        selectors = [item.strip() for item in selector_list.split("|") if item.strip()]
        deadline = _monotonic_ms() + timeout_ms
        last_error: Exception | None = None
        while _monotonic_ms() < deadline:
            for scope in _page_scopes(_active_page(page)):
                for selector in selectors:
                    try:
                        locator = scope.locator(f"{selector}:visible").first
                        if locator.count() > 0:
                            return locator
                    except PlaywrightError as exc:
                        last_error = exc
            sleep(0.1)
        raise PlaywrightTimeoutError(f"No se encontró selector: {selector_list}") from last_error

    def _fill_agente_carga_ruc(self, page: Page, ruc: str) -> None:
        """Marca el checkbox RUC del Agente de Carga y escribe el RUC de la empresa."""
        self._check_agente_carga_box(page)
        ruc_input = self._find_visible_quick(page, f"#{self.tipo.selector_input_ruc}", timeout_ms=8000)
        ruc_input.fill(ruc)
        LOGGER.info("RUC del agente de carga cargado: %s", ruc)

    def _check_agente_carga_box(self, page: Page) -> None:
        """Marca el checkbox sin desmarcarlo si ya estaba marcado (.check() es idempotente)."""
        checkbox_id = self.tipo.selector_checkbox_ruc
        for scope in _page_scopes(_active_page(page)):
            try:
                scope.locator(f"#{checkbox_id}").first.check(timeout=3000)
                LOGGER.info("Checkbox RUC del Agente de Carga marcado (%s).", checkbox_id)
                return
            except (PlaywrightError, PlaywrightTimeoutError):
                continue
        for scope in _page_scopes(_active_page(page)):
            try:
                if scope.evaluate(_set_checkbox_script(), checkbox_id):
                    LOGGER.info("Checkbox RUC del Agente de Carga marcado mediante DOM (%s).", checkbox_id)
                    return
            except PlaywrightError:
                continue
        LOGGER.warning("No se encontró el checkbox RUC del Agente de Carga (%s).", checkbox_id)
    def _logout(self, page: Page) -> None:
        if self.config.logout_manual:
            self._logout_manual(page)
        else:
            self._logout_os(page)

    def _logout_os(self, page: Page) -> None:
        inicio = _monotonic_ms()
        boton = self._wait_for_btn_salir(page, timeout_ms=30000)
        LOGGER.info("Botón Salir localizado en %s ms.", _monotonic_ms() - inicio)
        if boton is None:
            LOGGER.warning("No se encontró el botón Salir; no se pudo cerrar sesión.")
            return
        try:
            boton.dispatch_event("click")
        except (PlaywrightError, PlaywrightTimeoutError) as exc:
            LOGGER.warning("No se pudo clicar el botón Salir: %s", exc)
            return
        LOGGER.info("Botón Salir clicado (dispatch_event).")
        sleep(1)
        page.keyboard.press("Enter")
        LOGGER.info("Enter 1 enviado.")
        sleep(2)
        page.keyboard.press("Enter")
        LOGGER.info("Enter 2 enviado.")
        sleep(2)
        LOGGER.info("Sesión SUNAT cerrada con Enter.")

    def _find_btn_salir(self, page: Page, verbose: bool = True):
        for frame in page.frames:
            inicio = _monotonic_ms()
            try:
                locator = frame.locator("#btnSalir")
                if locator.count() > 0:
                    if verbose:
                        LOGGER.info(
                            "btnSalir encontrado en el frame '%s' tras %s ms.",
                            frame.name,
                            _monotonic_ms() - inicio,
                        )
                    return locator.first
            except PlaywrightError:
                pass
            if verbose:
                LOGGER.info("Frame '%s' revisado en %s ms.", frame.name, _monotonic_ms() - inicio)
        return None

    def _wait_for_btn_salir(self, page: Page, timeout_ms: int):
        deadline = _monotonic_ms() + timeout_ms
        chequeo_credenciales = _monotonic_ms() + 8000
        while _monotonic_ms() < deadline:
            boton = self._find_btn_salir(page, verbose=False)
            if boton is not None:
                return boton
            if _monotonic_ms() >= chequeo_credenciales:
                try:
                    if self._at_login_page(page):
                        raise AuthenticationError(
                            "SUNAT rechazó las credenciales o solicitó validación adicional."
                        )
                except (PlaywrightError, PlaywrightTimeoutError):
                    pass
            sleep(0.5)
        return None

    def _logout_manual(self, page: Page) -> None:
        selectors = self.config.selectors

        def mantener_dialogo(dialog) -> None:
            LOGGER.info(
                "Diálogo SUNAT abierto (%s). Acepta manualmente: 'Volver a cargar' y luego 'Salir'.",
                dialog.type,
            )

        page.on("dialog", mantener_dialogo)
        try:
            locator = self._first_available_locator(page, selectors.logout_link, timeout_ms=5000)
            locator.click(no_wait_after=True)
        except (PlaywrightError, PlaywrightTimeoutError):
            LOGGER.warning("No se encontró el botón Salir; no se pudo cerrar sesión.")
            page.remove_listener("dialog", mantener_dialogo)
            return

        LOGGER.info("Clickea 'Volver a cargar' y luego 'Salir' en el navegador para cerrar la sesión SUNAT.")
        deadline = _monotonic_ms() + 300000
        salio = False
        while _monotonic_ms() < deadline:
            try:
                if self._at_login_page(page):
                    salio = True
                    break
            except (PlaywrightError, PlaywrightTimeoutError):
                pass
            sleep(0.5)
        page.remove_listener("dialog", mantener_dialogo)
        if salio:
            LOGGER.info("Sesión SUNAT cerrada: se detectó la pantalla de inicio de sesión.")
        else:
            LOGGER.warning("No se detectó la pantalla de inicio de sesión dentro de los 5 minutos.")

    def _at_login_page(self, page: Page) -> bool:
        selectors = self.config.selectors
        ruc_selectors = [s.strip() for s in selectors.ruc_input.split("|") if s.strip()]
        user_selectors = [s.strip() for s in selectors.user_input.split("|") if s.strip()]
        try:
            scopes = _page_scopes(_active_page(page))
        except (PlaywrightError, PlaywrightTimeoutError):
            return False
        for scope in scopes:
            try:
                has_ruc = any(scope.locator(s).count() > 0 for s in ruc_selectors)
                has_user = any(scope.locator(s).count() > 0 for s in user_selectors)
                if has_ruc and has_user:
                    return True
            except PlaywrightError:
                continue
        return False

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

    def _fill_first_available(self, page: Page, selector_list: str, value: str, timeout_ms: int | None = None):
        locator = self._first_available_locator(page, selector_list, timeout_ms=timeout_ms)
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
                        scope.get_by_text(label, exact=True).first.click(timeout=750)
                        return
                    except (PlaywrightError, PlaywrightTimeoutError) as exc:
                        last_error = exc
            for scope in _page_scopes(_active_page(page)):
                for label in labels:
                    try:
                        if scope.evaluate(_click_menu_item_script(), {"label": label}):
                            LOGGER.info("Opción del menú activada mediante DOM: %s", label)
                            return
                    except PlaywrightError as exc:
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

def _set_checkbox_script() -> str:
    return r"""
    (id) => {
        const el = document.getElementById(id);
        if (!el) return false;
        if (!el.checked) {
            el.checked = true;
            el.dispatchEvent(new Event('change', { bubbles: true }));
        }
        return true;
    }
    """

def _click_menu_item_script() -> str:
    return r"""
    ({ label }) => {
        const normalize = (text) => (text || '')
            .toLowerCase()
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .replace(/\s+/g, ' ')
            .trim();
        const target = normalize(label);
        const isVisible = (node) => {
            const rect = node.getBoundingClientRect();
            return node.offsetParent !== null && rect.width > 0 && rect.height > 0;
        };
        const items = Array.from(document.querySelectorAll('span.spanNivelDescripcion'))
            .filter((item) => normalize(item.innerText || item.textContent) === target)
            .sort((a, b) => (isVisible(b) ? 1 : 0) - (isVisible(a) ? 1 : 0));
        const item = items[0];
        if (!item) return false;
        item.scrollIntoView({ block: 'nearest', inline: 'nearest' });
        item.click();
        return true;
    }
    """

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


def _context_pages(page: Page):
    try:
        return [candidate for candidate in page.context.pages if not candidate.is_closed()]
    except Exception:
        return [page]


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



