"""
Módulo de internacionalización (i18n) para soporte multi-idioma.

Carga cadenas de texto traducidas desde archivos JSON en locales/{lang}.json.
Las cadenas soportan formateo con .format(**kwargs) para variables dinámicas.

Idiomas soportados: "es" (español, default), "en" (inglés).
"""
import json
import os
import sys
import logging

logger = logging.getLogger(__name__)

SUPPORTED_LANGUAGES = ("es", "en")
DEFAULT_LANGUAGE = "es"


def _get_locales_dir(base_path: str) -> str:
    """
    Retorna la ruta a la carpeta locales/.

    Cuando el ejecutable está congelado por PyInstaller en modo one-file,
    los datos bundled se extraen a sys._MEIPASS. En modo script, usamos
    la raíz del proyecto recibida como base_path.
    """
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        # PyInstaller extrae los datas de la spec a sys._MEIPASS
        return os.path.join(sys._MEIPASS, "locales")
    return os.path.join(base_path, "locales")


class I18n:
    """
    Cargador de cadenas de texto traducidas.

    Uso:
        i18n = I18n(base_path, language="es")
        label = i18n.t("send")                  # → "Enviar Correo"
        msg   = i18n.t("msg_closing", n=5)       # → "Cerrando en 5 segundos..."
    """

    def __init__(self, base_path: str, language: str = DEFAULT_LANGUAGE):
        """
        Args:
            base_path: Ruta raíz del proyecto (para localizar locales/).
            language:  Código de idioma ("es" o "en").
        """
        if language not in SUPPORTED_LANGUAGES:
            logger.warning("Idioma '%s' no soportado. Usando '%s'.", language, DEFAULT_LANGUAGE)
            language = DEFAULT_LANGUAGE

        self.language = language
        self._base_path = base_path
        self._strings: dict = self._load(language)

    def _load(self, language: str) -> dict:
        """
        Carga el archivo JSON del idioma indicado.

        Si el archivo no existe, registra el error y retorna diccionario vacío
        (las claves faltantes retornarán la propia clave, lo que es aceptable
        como fallback degradado).
        """
        locales_dir = _get_locales_dir(self._base_path)
        file_path = os.path.join(locales_dir, f"{language}.json")

        if not os.path.exists(file_path):
            logger.error("Archivo de idioma no encontrado: %s", file_path)
            return {}

        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        logger.info("Idioma cargado: %s (%d cadenas)", language, len(data))
        return data

    def t(self, key: str, **kwargs) -> str:
        """
        Retorna la cadena traducida para la clave dada.

        Args:
            key:    Clave en el archivo de idioma.
            **kwargs: Variables para formatear la cadena
                      (ej: t("msg_closing", n=5)).

        Returns:
            Cadena traducida y formateada.
            Si la clave no existe, retorna la propia clave como fallback.
        """
        text = self._strings.get(key, key)
        if kwargs:
            try:
                text = text.format(**kwargs)
            except KeyError as exc:
                logger.warning("Placeholder %s no encontrado en clave '%s'", exc, key)
        return text
