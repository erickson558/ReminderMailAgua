import unittest
from unittest.mock import Mock, patch

from src.backend.email_service import send_email


class SendEmailTests(unittest.TestCase):
    @patch("src.backend.email_service.pythoncom.CoUninitialize")
    @patch("src.backend.email_service.pythoncom.CoInitialize")
    @patch("src.backend.email_service.win32.Dispatch")
    def test_send_email_keeps_sender_in_recipient_list(self, dispatch, coinitialize, couninitialize):
        sender_account = "erickson558@hotmail.com"

        account = Mock()
        account.SmtpAddress = sender_account
        account.AccountType = 3

        mail = Mock()
        outlook = Mock()
        outlook.CreateItem.return_value = mail
        outlook.Session.Accounts = [account]
        dispatch.return_value = outlook

        send_email(
            recipients=[sender_account, "other@example.com"],
            subject="subject",
            body="body",
            sender_account=sender_account,
        )

        self.assertEqual(mail.To, "erickson558@hotmail.com; other@example.com")
        mail.Send.assert_called_once_with()
        coinitialize.assert_called_once_with()
        couninitialize.assert_called_once_with()

    @patch("src.backend.email_service.pythoncom.CoUninitialize")
    @patch("src.backend.email_service.pythoncom.CoInitialize")
    @patch("src.backend.email_service.win32.Dispatch")
    def test_send_email_rejects_empty_recipient_list(self, dispatch, coinitialize, couninitialize):
        mail = Mock()
        outlook = Mock()
        outlook.CreateItem.return_value = mail
        outlook.Session.Accounts = []
        dispatch.return_value = outlook

        with self.assertRaisesRegex(ValueError, "La lista estaba vacía"):
            send_email([], "subject", "body")

        coinitialize.assert_called_once_with()
        couninitialize.assert_called_once_with()

    @patch("src.backend.email_service.pythoncom.CoUninitialize")
    @patch("src.backend.email_service.pythoncom.CoInitialize")
    @patch("src.backend.email_service.win32.Dispatch")
    def test_send_email_uses_exact_runtime_recipients_in_mail_to(self, dispatch, coinitialize, couninitialize):
        runtime_recipients = ["gui1@example.com", "gui2@example.com"]

        account = Mock()
        account.SmtpAddress = "sender@example.com"
        account.AccountType = 0

        mail = Mock()
        outlook = Mock()
        outlook.CreateItem.return_value = mail
        outlook.Session.Accounts = [account]
        dispatch.return_value = outlook

        send_email(
            recipients=runtime_recipients,
            subject="subject",
            body="body",
            sender_account="sender@example.com",
        )

        self.assertEqual(mail.To, "gui1@example.com; gui2@example.com")
        mail.Send.assert_called_once_with()
        coinitialize.assert_called_once_with()
        couninitialize.assert_called_once_with()


if __name__ == "__main__":
    unittest.main()