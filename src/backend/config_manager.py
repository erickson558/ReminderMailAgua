"""
Gestor de configuración de la aplicación.
Lee y escribe config.json con manejo de errores y valores por defecto.
Mantiene compatibilidad total con el formato v1 de config.json.
"""
import json
import os
import logging

logger = logging.getLogger(__name__)

# Valores por defecto: se aplican solo a claves no presentes en config.json
DEFAULT_CONFIG: dict = {
    "destinatarios": [],
    "asunto": "Reminder de Pagar el Agua del mes de [Mes anterior en letras]",
    "cuerpo": "Recordatorio de pagar el agua del [Mes anterior en letras] de [año en numero]",
    "auto_close": True,
    "auto_close_delay": 60,
    "auto_send_on_open": False,
    "cuenta_seleccionada": "",
    "language": "es",
}

CONFIG_FILENAME = "config.json"


def _resolve_config_path(base_path: str) -> str:
    """
    Resuelve la ubicación de config.json.

    Prioriza la carpeta base recibida. Como fallback para desarrollo, si el
    ejecutable vive dentro de dist/ y allí todavía no existe config.json,
    reutiliza el config.json de la raíz del proyecto para evitar trabajar con
    dos archivos de configuración distintos.
    """
    primary_path = os.path.join(base_path, CONFIG_FILENAME)
    if os.path.exists(primary_path):
        return primary_path

    parent_path = os.path.dirname(base_path)
    parent_config_path = os.path.join(parent_path, CONFIG_FILENAME)
    parent_has_project_markers = all(
        os.path.exists(os.path.join(parent_path, marker))
        for marker in ("main.py", "src")
    )

    if (
        os.path.basename(os.path.normpath(base_path)).lower() == "dist"
        and os.path.exists(parent_config_path)
        and parent_has_project_markers
    ):
        logger.info(
            "Usando config.json de la raíz del proyecto: %s",
            parent_config_path,
        )
        return parent_config_path

    return primary_path


class ConfigManager:
    """
    Maneja la lectura y escritura de config.json.

    El archivo se ubica junto al ejecutable/script. En desarrollo, si el exe
    corre desde dist/ dentro del repo, reutiliza el config.json de la raíz del
    proyecto para evitar configuraciones duplicadas.
    """

    def __init__(self, base_path: str):
        """
        Args:
            base_path: Ruta raíz del proyecto (donde vive config.json).
        """
        self.config_path = _resolve_config_path(base_path)
        self.data: dict = self._load()

    # ── Lectura ───────────────────────────────────────────────────────────────

    def _load(self) -> dict:
        """
        Carga config.json fusionando con defaults.
        Si el archivo no existe lo crea con los valores por defecto.
        """
        config = DEFAULT_CONFIG.copy()

        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, "r", encoding="utf-8") as f:
                    on_disk = json.load(f)
                # update() preserva claves nuevas y sobreescribe defaults solo con
                # valores existentes en disco → compatibilidad hacia adelante y atrás
                config.update(on_disk)
                logger.info("Configuración cargada desde %s", self.config_path)
            except (json.JSONDecodeError, OSError) as exc:
                logger.error("Error al leer config.json: %s — usando valores por defecto", exc)
        else:
            logger.info("config.json no encontrado. Creando con valores por defecto.")
            self._write(config)

        return config

    def get(self, key: str, default=None):
        """Retorna el valor de una clave de configuración."""
        return self.data.get(key, default)

    # ── Escritura ─────────────────────────────────────────────────────────────

    def save(self, new_data: dict) -> None:
        """
        Actualiza self.data con new_data y persiste en disco.

        Args:
            new_data: Diccionario con los valores a guardar.

        Raises:
            OSError: Si no se puede escribir el archivo.
        """
        self.data.update(new_data)
        self._write(self.data)
        logger.info("Configuración guardada en %s", self.config_path)

    def _write(self, data: dict) -> None:
        """Escribe el diccionario en config.json con formato indentado y UTF-8."""
        with open(self.config_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
