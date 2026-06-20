"""
GUI principal de ReminderMailAgua.

Implementada como clase (ReminderApp) para separar responsabilidades:
  - La GUI solo gestiona widgets y eventos de usuario.
  - La lógica de negocio (envío, fechas, config) vive en src/backend/.

El envío de correo se ejecuta en un hilo de fondo (threading.Thread) para
que la interfaz nunca se congele durante la operación COM de Outlook.
"""
import tkinter as tk
from tkinter import ttk, simpledialog
import threading
import webbrowser
import logging
import os

from src.backend.config_manager import ConfigManager
from src.backend.email_service import get_outlook_accounts, send_email
from src.backend.date_utils import resolve_placeholders
from src.frontend.i18n import I18n

logger = logging.getLogger(__name__)

# URL del botón donación "Cómprame una cerveza"
BEER_URL = "https://www.paypal.com/donate/?hosted_button_id=ZABFRXC2P3JQN"

# Colores de la paleta principal
COLOR_SEND_BG = "#4CAF50"   # Verde: botón de envío
COLOR_SEND_FG = "white"
COLOR_BEER_FG = "#FF8C00"   # Naranja: botón donación
COLOR_STATUS_OK = "green"
COLOR_STATUS_ERR = "red"
COLOR_STATUS_INFO = "blue"


class ReminderApp:
    """
    Aplicación principal de recordatorio de agua.

    Gestiona el ciclo de vida de la ventana tkinter, los widgets,
    y coordina las llamadas al backend para envío y configuración.
    """

    def __init__(self, base_path: str):
        """
        Args:
            base_path: Ruta raíz del proyecto. Se usa para localizar
                       config.json y la carpeta locales/.
        """
        self._base_path = base_path

        # ── Servicios de soporte ──────────────────────────────────────────────
        self._config = ConfigManager(base_path)
        self._i18n = I18n(base_path, self._config.get("language", "es"))

        # ── Ventana raíz ──────────────────────────────────────────────────────
        self.root = tk.Tk()
        self.root.title(self._t("title"))
        self.root.resizable(False, False)

        # ── Construir UI y cargar datos ───────────────────────────────────────
        self._build_ui()
        self._load_config_into_ui()
        self._update_status(self._t("msg_ready"), COLOR_STATUS_INFO)
        self._schedule_auto_send_on_open_if_enabled()

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _t(self, key: str, **kwargs) -> str:
        """Atajo para traducción de cadenas."""
        return self._i18n.t(key, **kwargs)

    def _update_status(self, message: str, color: str = "black") -> None:
        """
        Actualiza la barra de estado de forma thread-safe.
        Usa root.after(0) para evitar llamadas directas a tkinter desde hilos.
        """
        self.root.after(
            0,
            lambda: self._status_label.config(text=message, fg=color),
        )

    def _get_log_path(self) -> str:
        """Retorna la ruta del archivo de log junto al script o ejecutable."""
        return os.path.join(self._base_path, "reminderagua.log")

    def _get_recipients_from_ui(self) -> list[str]:
        """Retorna la lista actual de destinatarios visibles en la GUI."""
        return self._normalize_recipients(self._listbox_recipients.get(0, tk.END))

    def _normalize_recipients(self, recipients) -> list[str]:
        """Limpia espacios, elimina vacíos y evita duplicados case-insensitive."""
        normalized = []
        seen = set()

        for raw_value in recipients:
            recipient = raw_value.strip()
            if not recipient:
                continue

            recipient_key = recipient.lower()
            if recipient_key in seen:
                continue

            seen.add(recipient_key)
            normalized.append(recipient)

        return normalized

    def _replace_recipients_in_ui(self, recipients: list[str]) -> None:
        """Normaliza y reemplaza el contenido del listbox de destinatarios."""
        self._listbox_recipients.delete(0, tk.END)
        for recipient in self._normalize_recipients(recipients):
            self._listbox_recipients.insert(tk.END, recipient)

    # ── Construcción de la UI ─────────────────────────────────────────────────

    def _build_ui(self) -> None:
        """Orquesta la creación de todos los componentes de la interfaz."""
        self._build_menu()
        self._build_recipients_frame()
        self._build_subject_field()
        self._build_body_field()
        self._build_account_frame()
        self._build_auto_close_frame()
        self._build_status_bar()
        self._build_action_buttons()
        self._build_beer_button()

    def _build_menu(self) -> None:
        """Crea la barra de menú con selector de idioma."""
        menubar = tk.Menu(self.root)

        # Menú de idioma: cada opción recarga la app con el idioma seleccionado
        lang_menu = tk.Menu(menubar, tearoff=0)
        lang_menu.add_command(label="Español", command=lambda: self._change_language("es"))
        lang_menu.add_command(label="English", command=lambda: self._change_language("en"))
        menubar.add_cascade(label=self._t("language"), menu=lang_menu)

        self.root.config(menu=menubar)

    def _build_recipients_frame(self) -> None:
        """Frame con listbox de destinatarios y botones de gestión."""
        frame = tk.LabelFrame(self.root, text=self._t("recipients"), padx=5, pady=5)
        frame.pack(pady=(10, 0), fill=tk.X, padx=10)

        # Listbox con selección múltiple para poder eliminar varios a la vez
        self._listbox_recipients = tk.Listbox(
            frame, width=55, height=5, selectmode=tk.EXTENDED
        )
        self._listbox_recipients.pack(pady=(0, 5))

        btn_frame = tk.Frame(frame)
        btn_frame.pack()
        tk.Button(btn_frame, text=self._t("add"), width=15, command=self._add_recipient).pack(
            side=tk.LEFT, padx=5
        )
        tk.Button(btn_frame, text=self._t("remove"), width=15, command=self._remove_recipient).pack(
            side=tk.LEFT, padx=5
        )

    def _build_subject_field(self) -> None:
        """Etiqueta y campo de entrada para el asunto del correo."""
        tk.Label(self.root, text=self._t("subject")).pack(pady=(10, 0))
        self._entry_subject = tk.Entry(self.root, width=55)
        self._entry_subject.pack(pady=5)

    def _build_body_field(self) -> None:
        """Etiqueta y área de texto para el cuerpo del correo."""
        tk.Label(self.root, text=self._t("body")).pack(pady=(10, 0))
        self._text_body = tk.Text(self.root, width=55, height=10)
        self._text_body.pack(pady=5)

    def _build_account_frame(self) -> None:
        """Frame con combobox para seleccionar la cuenta de Outlook desde donde enviar."""
        frame = tk.LabelFrame(self.root, text=self._t("send_account"), padx=10, pady=5)
        frame.pack(padx=10, pady=5, fill="both")

        tk.Label(frame, text=self._t("select_account")).pack(anchor="w", pady=(0, 3))

        # Obtiene las cuentas disponibles en Outlook (puede ser [] si Outlook no está instalado)
        self._accounts = get_outlook_accounts()
        self._combobox_account = ttk.Combobox(
            frame, values=self._accounts, state="readonly", width=50
        )
        if self._accounts:
            self._combobox_account.current(0)
        self._combobox_account.pack(anchor="w")

    def _build_auto_close_frame(self) -> None:
        """Frame con checkboxes de comportamiento automático y su configuración."""
        frame = tk.LabelFrame(
            self.root, text=self._t("auto_close"), padx=10, pady=5
        )
        frame.pack(padx=10, pady=5, fill="both")

        self._auto_close_var = tk.BooleanVar()
        tk.Checkbutton(
            frame, text=self._t("auto_close_check"), variable=self._auto_close_var
        ).pack(anchor="w")

        self._auto_send_on_open_var = tk.BooleanVar()
        tk.Checkbutton(
            frame,
            text=self._t("auto_send_on_open_check"),
            variable=self._auto_send_on_open_var,
        ).pack(anchor="w", pady=(5, 0))

        tk.Label(frame, text=self._t("auto_close_delay")).pack(anchor="w", pady=(5, 0))
        self._auto_close_delay_var = tk.StringVar()
        tk.Entry(frame, textvariable=self._auto_close_delay_var, width=10).pack(anchor="w")

    def _build_action_buttons(self) -> None:
        """Fila de botones principales: Enviar, Guardar configuración, Salir."""
        frame = tk.Frame(self.root)
        frame.pack(pady=10)

        # Botón Enviar: verde para destacarlo como acción principal
        self._btn_send = tk.Button(
            frame,
            text=self._t("send"),
            width=18,
            command=self._send_email,
            bg=COLOR_SEND_BG,
            fg=COLOR_SEND_FG,
        )
        self._btn_send.pack(side=tk.LEFT, padx=4)

        tk.Button(
            frame, text=self._t("save_config"), width=18, command=self._save_config
        ).pack(side=tk.LEFT, padx=4)

        self._btn_exit = tk.Button(
            frame, text=self._t("exit"), width=12, command=self._exit
        )
        self._btn_exit.pack(side=tk.LEFT, padx=4)

    def _build_beer_button(self) -> None:
        """Botón de donación: abre PayPal en el navegador predeterminado."""
        tk.Button(
            self.root,
            text=self._t("buy_beer"),
            fg=COLOR_BEER_FG,
            cursor="hand2",       # Cursor de mano para indicar que es clickeable
            relief=tk.FLAT,       # Sin borde: aspecto de enlace
            command=lambda: webbrowser.open(BEER_URL),
        ).pack(pady=(0, 6))

    def _build_status_bar(self) -> None:
        """Bloque de estado visible con acceso rápido al archivo de log."""
        frame = tk.LabelFrame(self.root, text=self._t("status"), padx=8, pady=6)
        frame.pack(padx=10, pady=(0, 8), fill=tk.X)

        self._status_label = tk.Label(
            frame,
            text=self._t("msg_ready"),
            anchor=tk.W,
            justify=tk.LEFT,
            wraplength=410,
            bg="#FFF7D6",
            fg=COLOR_STATUS_INFO,
            relief=tk.SUNKEN,
            bd=1,
            padx=8,
            pady=6,
        )
        self._status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        tk.Button(
            frame,
            text=self._t("open_log"),
            width=12,
            command=self._open_log_file,
        ).pack(side=tk.LEFT, padx=(8, 0))

    # ── Carga de configuración en la UI ───────────────────────────────────────

    def _load_config_into_ui(self) -> None:
        """Puebla los widgets con los valores del config.json."""
        cfg = self._config.data

        self._replace_recipients_in_ui(cfg.get("destinatarios", []))

        self._entry_subject.insert(0, cfg.get("asunto", ""))
        self._text_body.insert("1.0", cfg.get("cuerpo", ""))
        self._auto_close_var.set(cfg.get("auto_close", True))
        self._auto_close_delay_var.set(str(cfg.get("auto_close_delay", 60)))
        self._auto_send_on_open_var.set(cfg.get("auto_send_on_open", False))

        # Restaura la última cuenta usada si sigue estando disponible en Outlook
        saved_account = cfg.get("cuenta_seleccionada", "")
        if saved_account and saved_account in self._accounts:
            idx = self._accounts.index(saved_account)
            self._combobox_account.current(idx)

    def _schedule_auto_send_on_open_if_enabled(self) -> None:
        """Programa el autoenvío opcional una vez que la ventana ya fue creada."""
        if not self._auto_send_on_open_var.get():
            return

        self._update_status(self._t("msg_auto_send_pending"), COLOR_STATUS_INFO)
        self.root.after(500, self._send_email)

    def _open_log_file(self) -> None:
        """Abre el archivo de log actual para revisar errores de envío."""
        try:
            os.startfile(self._get_log_path())
            self._update_status(self._t("msg_open_log_ok"), COLOR_STATUS_INFO)
        except OSError as exc:
            logger.exception("No se pudo abrir el archivo de log")
            self._update_status(self._t("msg_open_log_error", e=exc), COLOR_STATUS_ERR)

    # ── Acciones de usuario ───────────────────────────────────────────────────

    def _add_recipient(self) -> None:
        """Muestra un diálogo para ingresar un nuevo correo destinatario."""
        email = simpledialog.askstring(
            self._t("dialog_add_recipient"), self._t("msg_add_recipient")
        )
        if email and email.strip():
            recipients = self._get_recipients_from_ui()
            recipients.append(email)
            self._replace_recipients_in_ui(recipients)
            self._update_status(self._t("msg_added"), COLOR_STATUS_OK)

    def _remove_recipient(self) -> None:
        """Elimina los destinatarios seleccionados en el listbox."""
        selection = self._listbox_recipients.curselection()
        if not selection:
            self._update_status(self._t("msg_select_remove"), COLOR_STATUS_ERR)
            return
        # Elimina en orden inverso para que los índices no se desplacen
        for idx in reversed(selection):
            self._listbox_recipients.delete(idx)
        self._update_status(self._t("msg_removed"), COLOR_STATUS_OK)

    def _send_email(self) -> None:
        """
        Valida los campos y dispara el envío en un hilo de fondo.

        El hilo es daemon=True para que no bloquee el cierre del proceso
        si el usuario cierra la app mientras envía.
        """
        recipients = self._get_recipients_from_ui()
        if not recipients:
            self._update_status(self._t("msg_add_one"), COLOR_STATUS_ERR)
            return

        # Resolver placeholders ANTES de enviar (con fechas actuales)
        subject = resolve_placeholders(self._entry_subject.get().strip())
        body = resolve_placeholders(self._text_body.get("1.0", tk.END).strip())
        account = self._combobox_account.get()

        # Deshabilitar botón para evitar dobles envíos accidentales
        self._btn_send.config(state=tk.DISABLED)
        logger.info(
            "Preparando envío. Destinatarios GUI=%s | Cuenta=%s | Log=%s",
            recipients,
            account,
            self._get_log_path(),
        )
        self._update_status(self._t("msg_sending"), COLOR_STATUS_INFO)

        thread = threading.Thread(
            target=self._send_in_background,
            args=(recipients, subject, body, account),
            daemon=True,
        )
        thread.start()

    def _send_in_background(
        self, recipients: list, subject: str, body: str, account: str
    ) -> None:
        """
        Ejecuta el envío COM en el hilo de fondo.
        Actualiza el estado y arranca el countdown via root.after (thread-safe).
        """
        try:
            send_email(recipients, subject, body, account)
            self._update_status(self._t("msg_sent"), COLOR_STATUS_OK)
            logger.info("Correo enviado. Destinatarios: %s | Cuenta: %s", recipients, account)

            if self._auto_close_var.get():
                delay = int(self._auto_close_delay_var.get())
                # root.after es la forma segura de llamar a tkinter desde un hilo
                self.root.after(0, lambda: self._start_countdown(delay))

        except Exception as exc:
            logger.exception("Error al enviar correo en hilo de fondo")
            self._update_status(
                f"{self._t('msg_error_send', e=exc)} {self._t('msg_check_log')}",
                COLOR_STATUS_ERR,
            )

        finally:
            # Siempre re-habilita el botón al terminar (éxito o error)
            self.root.after(0, lambda: self._btn_send.config(state=tk.NORMAL))

    def _start_countdown(self, seconds: int) -> None:
        """
        Inicia la cuenta regresiva para el cierre automático.
        Usa root.after recursivo para no bloquear el event loop de tkinter.
        """
        def _tick():
            nonlocal seconds
            if seconds > 0:
                self._update_status(self._t("msg_closing", n=seconds), COLOR_STATUS_OK)
                seconds -= 1
                self.root.after(1000, _tick)
            else:
                self._exit()

        _tick()

    def _save_config(self) -> None:
        """Recopila el estado actual de los widgets y lo persiste en config.json."""
        data = {
            "destinatarios": self._get_recipients_from_ui(),
            "asunto": self._entry_subject.get(),
            "cuerpo": self._text_body.get("1.0", tk.END).strip(),
            "auto_close": self._auto_close_var.get(),
            "auto_close_delay": int(self._auto_close_delay_var.get() or 60),
            "auto_send_on_open": self._auto_send_on_open_var.get(),
            "cuenta_seleccionada": self._combobox_account.get(),
            "language": self._i18n.language,
        }
        try:
            self._config.save(data)
            self._update_status(self._t("msg_saved"), COLOR_STATUS_OK)
        except OSError as exc:
            self._update_status(self._t("msg_error_save", e=exc), COLOR_STATUS_ERR)

    def _change_language(self, lang: str) -> None:
        """
        Cambia el idioma de la UI.

        Guarda la preferencia en config.json, destruye la ventana actual
        y recrea toda la aplicación con el nuevo idioma — es el enfoque más
        limpio para tkinter sin mantener StringVars por cada widget.
        """
        # Capturar el estado actual antes de destruir la ventana
        current_data = {
            "destinatarios": self._get_recipients_from_ui(),
            "asunto": self._entry_subject.get(),
            "cuerpo": self._text_body.get("1.0", tk.END).strip(),
            "auto_close": self._auto_close_var.get(),
            "auto_close_delay": int(self._auto_close_delay_var.get() or 60),
            "auto_send_on_open": self._auto_send_on_open_var.get(),
            "cuenta_seleccionada": self._combobox_account.get(),
            "language": lang,  # Guardar el nuevo idioma
        }
        self._config.save(current_data)

        # Destruir ventana actual y recrear la app con el nuevo idioma
        self.root.destroy()
        new_app = ReminderApp(self._base_path)
        new_app.run()

    def _exit(self) -> None:
        """Guarda la configuración y cierra la aplicación."""
        self._save_config()
        self.root.destroy()

    # ── Ciclo principal ───────────────────────────────────────────────────────

    def run(self) -> None:
        """Inicia el event loop de tkinter. Bloquea hasta que se cierra la ventana."""
        self.root.mainloop()
