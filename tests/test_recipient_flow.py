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


class FakeText:
    def __init__(self, value=""):
        self.value = value

    def get(self, start, end):
        return self.value


class FakeVar:
    def __init__(self, value):
        self.value = value

    def get(self):
        return self.value


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
        app._combobox_account = FakeEntry("sender@example.com")
        app._auto_close_var = FakeVar(False)
        app._auto_close_delay_var = FakeEntry("60")
        app._btn_send = FakeButton()
        app._config = Mock()
        app._i18n = Mock(language="es")
        app._update_status = Mock()
        app._t = lambda key, **kwargs: key
        app._start_countdown = Mock()
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

        app._save_config()

        saved_data = app._config.save.call_args.args[0]
        self.assertEqual(
            saved_data["destinatarios"],
            ["kept@example.com", "second@example.com"],
        )


if __name__ == "__main__":
    unittest.main()