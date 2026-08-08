from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class PasoFormulario:
    input_id: str
    texto: str


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
