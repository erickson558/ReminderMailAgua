import unittest
from unittest.mock import Mock, patch

from src.backend.email_service import send_email


class SendEmailTests(unittest.TestCase):
    @patch("src.backend.email_service.win32.Dispatch")
    def test_send_email_keeps_sender_in_recipient_list(self, dispatch):
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

    @patch("src.backend.email_service.win32.Dispatch")
    def test_send_email_rejects_empty_recipient_list(self, dispatch):
        mail = Mock()
        outlook = Mock()
        outlook.CreateItem.return_value = mail
        outlook.Session.Accounts = []
        dispatch.return_value = outlook

        with self.assertRaisesRegex(ValueError, "La lista estaba vacía"):
            send_email([], "subject", "body")


if __name__ == "__main__":
    unittest.main()