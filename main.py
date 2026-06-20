"""
ReminderMailAgua - Aplicación de recordatorio de pago de agua.
Punto de entrada principal: configura logging, determina rutas y lanza la GUI.
"""
import sys
import os
import logging

# ── Rutas ─────────────────────────────────────────────────────────────────────
# Cuando el app está compilada con PyInstaller, sys.frozen es True y
# sys.executable apunta al .exe. En modo script, usamos __file__.
if getattr(sys, 'frozen', False):
    BASE_PATH = os.path.dirname(sys.executable)
else:
    BASE_PATH = os.path.dirname(os.path.abspath(__file__))

# Agrega la raíz del proyecto al path para que "from src.X import Y" funcione
sys.path.insert(0, BASE_PATH)

# ── Logging ───────────────────────────────────────────────────────────────────
# Escribe en consola y en archivo de log junto al ejecutable/script.
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(name)-25s  %(levelname)-8s  %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(
            os.path.join(BASE_PATH, "reminderagua.log"),
            encoding="utf-8",
            mode="a"
        ),
    ],
)

logger = logging.getLogger(__name__)

# ── Lanzar aplicación ─────────────────────────────────────────────────────────
from src.frontend.app import ReminderApp  # noqa: E402 (import after path setup)

if __name__ == "__main__":
    logger.info("Iniciando ReminderMailAgua v2.0")
    app = ReminderApp(BASE_PATH)
    app.run()
    logger.info("Aplicación cerrada.")
