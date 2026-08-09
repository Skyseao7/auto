from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class PasoFormulario:
    input_id: str
    texto: str
    tipo_control: str = "combobox"  # "combobox" (dijit) o "select" (HTML)


@dataclass(frozen=True)
class TipoReporte:
    nombre: str
    prefijo: str
    sufijo: str
    hoja_destino: str
    hoja_transmisiones: str
    etiqueta_manifiesto: str = "Manifiesto Desconsolidado"
    codigo_regex: str = r"\d{2}-\d{3}-\d+-\d{4}-\d+"
    tiene_contenedores: bool = True
    pasos_formulario: tuple[PasoFormulario, ...] = ()
    selector_formulario: str = "tipoBusquedaFechaNumeracionDesconsolidado"
    selector_radio_fecha: str = "tipoBusquedaFechaNumeracionDesconsolidado"
    radio_fecha_label: str = "Fecha de Numeración de Manifiesto Desconsolidado"
    selector_checkbox_ruc: str = "tipoBusquedaAgenteCarga"
    selector_input_ruc: str = "numeroRucAgenteCarga"
    date_from_input: str = "#fechaInicial|input[name*='fecIni']|input[placeholder*='Desde']|input[aria-label*='Desde']"
    date_to_input: str = "#fechaFinal|input[name*='fecFin']|input[placeholder*='Hasta']|input[aria-label*='Hasta']"
    selector_grid: str = "gridManifiestoCarga"
    selector_btn_consultar: str = "#accion5_label|text=Consultar|button:has-text('Consultar')"
    selector_grid_filas: str = ".dojoxGridRow"
    selector_grid_siguiente: str = '[title="Página siguiente"]'
    selector_grid_vacio: str = "div.dojoxGridMasterMessages"
    columnas_extraccion: int = 9

    @property
    def es_expo(self) -> bool:
        return self.prefijo == "EXPO"
    cabeceras_transmisiones: tuple[str, ...] = (
        "Manifiesto de Carga",
        "Fecha del Manifiesto de Carga",
        "Manifiesto Desconsolidado",
        "Fecha del Manifiesto Desconsolidado",
        "Agente de Carga RUC",
        "Número de Ticket",
        "Fecha de Llegada",
        "Fecha de Término de la Descarga",
        "Estado del Manifiesto Desconsolidado",
    )
    mapeo_copia: tuple[tuple[int, int, bool], ...] = (
        (3, 3, True),
        (1, 4, False),
        (7, 12, False),
    )
    columna_tipo_numeracion: int = 8
    ruta_menu: tuple[tuple[str, ...], ...] = (
        ("Operaciones de Comercio Exterior", "Operador de Comercio Exterior"),
        ("Manifiesto de Carga de Ingreso",),
        ("Consultas",),
        (
            "Consulta Manifiesto Desconsolidado",
            "Consulta del Manifiesto Desconsolidado",
            "Consulta de Manifiesto Desconsolidado",
            "Consulta Manifesto Desconsolidado",
        ),
    )
    columnas_detalle: tuple[tuple[str, int], ...] = (
        ("master", 2),
        ("puerto_embarque", 5),
        ("cnt", 7),
        ("fecha_hijo", 9),
        ("fecha_master", 10),
        ("fecha_info", 11),
    )


IMPO118 = TipoReporte(
    nombre="IMPO118",
    prefijo="IMPO",
    sufijo="118",
    hoja_destino="IMPO118",
    hoja_transmisiones="IMPO118-Transmisiones",
)

IMPO235 = TipoReporte(
    nombre="IMPO235",
    prefijo="IMPO",
    sufijo="235",
    hoja_destino="IMPO235",
    hoja_transmisiones="IMPO235-Transmisiones",
    tiene_contenedores=False,
    pasos_formulario=(
        PasoFormulario(input_id="codigoAduana", texto="235-AEREA DEL CALLAO"),
        PasoFormulario(input_id="viaTransporte", texto="AEREO"),
    ),
    mapeo_copia=(
        (3, 3, True),
        (1, 4, False),
        (7, 11, False),
    ),
    columnas_detalle=(
        ("master", 2),
        ("puerto_embarque", 5),
        ("fecha_hijo", 8),
        ("fecha_master", 9),
        ("fecha_info", 10),
    ),
    columna_tipo_numeracion=7,
)

EXPO118 = TipoReporte(
    nombre="EXPO118",
    prefijo="EXPO",
    sufijo="118",
    hoja_destino="EXPO118",
    hoja_transmisiones="EXPO118-Transmisiones",
    ruta_menu=(
        ("Operaciones de Comercio Exterior", "Operador de Comercio Exterior"),
        ("Manifiesto de Carga de Salida",),
        ("Consultas",),
        (
            "Control del Cumplimiento del Manifiesto Consolidado",
            "Control de Cumplimiento del Manifiesto Consolidado",
        ),
    ),
    pasos_formulario=(
        PasoFormulario(input_id="selCodigoAduana", texto="118", tipo_control="select"),
    ),
    selector_formulario="tipoBusquedaFechaTransCons",
    selector_radio_fecha="tipoBusquedaFechaTransCons",
    radio_fecha_label="Fecha de Transmisión del Manifiesto Consolidado",
    selector_checkbox_ruc="chbRucAgente",
    selector_input_ruc="txtNroRucAgente",
    date_from_input="#txtFechaInicial|input[name='txtFechaInicial']",
    date_to_input="#txtFechaFinal|input[name='txtFechaFinal']",
    selector_grid="tblLista",
    selector_btn_consultar="#btnBuscar|text=Consultar|button:has-text('Consultar')",
    selector_grid_filas="tbody tr",
    selector_grid_siguiente="#tblLista_next",
    selector_grid_vacio="td.dataTables_empty",
    columnas_extraccion=13,
    cabeceras_transmisiones=(
        "Manifiesto de Carga",
        "Manifiesto Consolidado",
        "Agente de Carga Internacional",
        "Detalle Master",
        "Documento de Transporte Master",
        "Detalle Hijo",
        "Documento de Transporte Hijo",
        "Fecha y Hora de Transmisión de Documento de Transporte Hijo (1)",
        "Fecha y Hora del Término del Embarque",
        "Fecha Limite Para Transmitir (2)",
        "Plazo Excedido (2-1)",
        "Condición",
        "Estado del Documento de Transporte Hijo",
    ),
)

EXPO235 = TipoReporte(
    nombre="EXPO235",
    prefijo="EXPO",
    sufijo="235",
    hoja_destino="EXPO235",
    hoja_transmisiones="EXPO235-Transmisiones",
    tiene_contenedores=False,
)

TIPOS = {
    "impo118": IMPO118,
    "impo235": IMPO235,
    "expo118": EXPO118,
    "expo235": EXPO235,
}


def obtener_tipo(nombre: str) -> TipoReporte:
    key = (nombre or "").strip().lower().replace("-", "").replace("_", "")
    if key not in TIPOS:
        raise ValueError(f"Reporte no reconocido: {nombre}. Válidos: {', '.join(TIPOS)}")
    return TIPOS[key]