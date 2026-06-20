import unittest
from unittest.mock import Mock, patch

from src.frontend import app as app_module


class FakeListbox:
    def __init__(self, items=None):
        self.items = list(items or [])
        self.selection = ()

    def get(self, start, end=None):
        if start == 0 and end == app_module.tk.END:
            return tuple(self.items)
        return self.items[start]

    def insert(self, index, value):
        self.items.append(value)

    def delete(self, start, end=None):
        if start == 0 and end == app_module.tk.END:
            self.items.clear()
            return
        del self.items[start]

    def curselection(self):
        return self.selection

    def set_selection(self, *indexes):
        self.selection = indexes


class FakeEntry:
    def __init__(self, value=""):
        self.value = value

    def get(self):
        return self.value

    def insert(self, index, value):
        self.value = value

    def set(self, value):
        self.value = value


class FakeText:
    def __init__(self, value=""):
        self.value = value

    def get(self, start, end):
        return self.value

    def insert(self, index, value):
        self.value = value


class FakeVar:
    def __init__(self, value):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


class FakeCombobox(FakeEntry):
    def __init__(self, value=""):
        super().__init__(value)
        self.current_index = None

    def current(self, index):
        self.current_index = index


class FakeButton:
    def __init__(self):
        self.states = []

    def config(self, **kwargs):
        self.states.append(kwargs)


class FakeRoot:
    def __init__(self):
        self.after_calls = []

    def after(self, delay, callback):
        self.after_calls.append((delay, callback))


class ReminderAppRecipientFlowTests(unittest.TestCase):
    def _build_app(self, recipients=None):
        app = app_module.ReminderApp.__new__(app_module.ReminderApp)
        app.root = FakeRoot()
        app._listbox_recipients = FakeListbox(recipients)
        app._entry_subject = FakeEntry("Asunto")
        app._text_body = FakeText("Cuerpo")
        app._combobox_account = FakeCombobox("sender@example.com")
        app._auto_close_var = FakeVar(False)
        app._auto_close_delay_var = FakeEntry("60")
        app._auto_send_on_open_var = FakeVar(False)
        app._btn_send = FakeButton()
        app._config = Mock()
        app._i18n = Mock(language="es")
        app._update_status = Mock()
        app._t = lambda key, **kwargs: key
        app._start_countdown = Mock()
        app._accounts = ["sender@example.com", "other@example.com"]
        app._base_path = "D:\\app"
        return app

    def test_get_recipients_from_ui_normalizes_and_deduplicates(self):
        app = self._build_app([
            "  first@example.com  ",
            "",
            "SECOND@example.com",
            "second@example.com  ",
        ])

        recipients = app._get_recipients_from_ui()

        self.assertEqual(recipients, ["first@example.com", "SECOND@example.com"])

    @patch("src.frontend.app.simpledialog.askstring", return_value="  new@example.com  ")
    def test_add_recipient_updates_gui_list(self, askstring):
        app = self._build_app(["existing@example.com"])

        app._add_recipient()

        self.assertEqual(
            app._listbox_recipients.items,
            ["existing@example.com", "new@example.com"],
        )
        app._update_status.assert_called_once_with("msg_added", app_module.COLOR_STATUS_OK)
        askstring.assert_called_once()

    def test_remove_recipient_deletes_selected_indexes(self):
        app = self._build_app(["a@example.com", "b@example.com", "c@example.com"])
        app._listbox_recipients.set_selection(0, 2)

        app._remove_recipient()

        self.assertEqual(app._listbox_recipients.items, ["b@example.com"])
        app._update_status.assert_called_once_with("msg_removed", app_module.COLOR_STATUS_OK)

    @patch("src.frontend.app.resolve_placeholders", side_effect=lambda value: f"resolved::{value}")
    @patch("src.frontend.app.threading.Thread")
    def test_send_email_uses_current_gui_recipients(self, thread_cls, resolve_placeholders):
        app = self._build_app([" gui1@example.com ", "gui2@example.com", "GUI2@example.com"])
        app._config.data = {"destinatarios": ["config-only@example.com"]}

        thread_instance = Mock()
        thread_cls.return_value = thread_instance

        app._send_email()

        thread_cls.assert_called_once()
        _, kwargs = thread_cls.call_args
        self.assertEqual(
            kwargs["args"],
            (["gui1@example.com", "gui2@example.com"], "resolved::Asunto", "resolved::Cuerpo", "sender@example.com"),
        )
        self.assertTrue(kwargs["daemon"])
        thread_instance.start.assert_called_once_with()
        self.assertIn({"state": app_module.tk.DISABLED}, app._btn_send.states)
        app._update_status.assert_called_once_with("msg_sending", app_module.COLOR_STATUS_INFO)
        self.assertEqual(resolve_placeholders.call_count, 2)

    def test_save_config_persists_current_gui_recipients(self):
        app = self._build_app([" kept@example.com ", "KEPT@example.com", "second@example.com"])
        app._auto_send_on_open_var.set(True)

        app._save_config()

        saved_data = app._config.save.call_args.args[0]
        self.assertEqual(
            saved_data["destinatarios"],
            ["kept@example.com", "second@example.com"],
        )
        self.assertTrue(saved_data["auto_send_on_open"])

    def test_schedule_auto_send_on_open_when_enabled(self):
        app = self._build_app(["first@example.com"])
        app._send_email = Mock()
        app._auto_send_on_open_var.set(True)

        app._schedule_auto_send_on_open_if_enabled()

        self.assertEqual(app.root.after_calls, [(500, app._send_email)])
        app._update_status.assert_called_once_with(
            "msg_auto_send_pending", app_module.COLOR_STATUS_INFO
        )

    def test_load_config_populates_auto_send_and_account(self):
        app = self._build_app([])
        app._config.data = {
            "destinatarios": [" gui@example.com ", "gui@example.com"],
            "asunto": "Asunto cargado",
            "cuerpo": "Cuerpo cargado",
            "auto_close": True,
            "auto_close_delay": 45,
            "auto_send_on_open": True,
            "cuenta_seleccionada": "other@example.com",
        }

        app._load_config_into_ui()

        self.assertEqual(app._listbox_recipients.items, ["gui@example.com"])
        self.assertEqual(app._entry_subject.get(), "Asunto cargado")
        self.assertEqual(app._text_body.get("1.0", app_module.tk.END), "Cuerpo cargado")
        self.assertTrue(app._auto_close_var.get())
        self.assertEqual(app._auto_close_delay_var.get(), "45")
        self.assertTrue(app._auto_send_on_open_var.get())
        self.assertEqual(app._combobox_account.current_index, 1)

    def test_get_log_path_points_to_base_path(self):
        app = self._build_app([])

        self.assertEqual(app._get_log_path(), "D:\\app\\reminderagua.log")

    @patch("src.frontend.app.logger")
    @patch("src.frontend.app.send_email", side_effect=RuntimeError("fallo outlook"))
    def test_send_in_background_reports_error_and_log_hint(self, send_email, logger_mock):
        app = self._build_app(["gui@example.com"])
        app._status_label = Mock()

        app._send_in_background(["gui@example.com"], "Asunto", "Cuerpo", "sender@example.com")

        for _, callback in app.root.after_calls:
            callback()

        logger_mock.exception.assert_called_once()
        self.assertIn({"state": app_module.tk.NORMAL}, app._btn_send.states)


if __name__ == "__main__":
    unittest.main()