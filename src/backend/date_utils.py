"""
Utilidades de fecha para resolución de placeholders dinámicos en asunto/cuerpo.

Placeholders soportados:
    [Mes anterior en letras]  →  Nombre del mes anterior en español (ej: "Mayo")
    [año en numero]           →  Año correspondiente al mes anterior (ej: "2026")
"""
import datetime
import logging

logger = logging.getLogger(__name__)

# Diccionario de meses: número → nombre en español (minúsculas)
MESES_ES: dict[int, str] = {
    1: "enero",   2: "febrero", 3: "marzo",    4: "abril",
    5: "mayo",    6: "junio",   7: "julio",     8: "agosto",
    9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre",
}


def _get_previous_month_info() -> tuple[str, str]:
    """
    Calcula el nombre del mes anterior y su año en base a la fecha actual.

    El "mes anterior" se usa porque el reminder se envía el mes siguiente
    al período que se está pagando (ej: en junio se paga el agua de mayo).

    Returns:
        Tupla (nombre_mes, año_str), ej: ("Mayo", "2026").
    """
    now = datetime.datetime.now()

    if now.month == 1:
        # Enero → mes anterior es diciembre del año pasado
        prev_month = 12
        prev_year = now.year - 1
    else:
        prev_month = now.month - 1
        prev_year = now.year

    month_name = MESES_ES[prev_month].capitalize()
    return month_name, str(prev_year)


def resolve_placeholders(text: str) -> str:
    """
    Reemplaza los placeholders dinámicos en un texto con los valores actuales.

    Args:
        text: Cadena con posibles placeholders.

    Returns:
        Cadena con placeholders reemplazados. Si no hay placeholders, retorna
        el texto original sin cambios.
    """
    month_name, year_str = _get_previous_month_info()

    result = text.replace("[Mes anterior en letras]", month_name)
    result = result.replace("[año en numero]", year_str)

    if result != text:
        logger.debug("Placeholders resueltos: mes=%s, año=%s", month_name, year_str)

    return result
