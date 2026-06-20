import datetime
import unittest
from unittest.mock import patch

from src.backend.date_utils import resolve_placeholders


REAL_DATETIME = datetime.datetime


class ResolvePlaceholdersTests(unittest.TestCase):
    @patch("src.backend.date_utils.datetime.datetime")
    def test_resolve_placeholders_uses_previous_month_from_pc_date(self, mock_datetime):
        mock_datetime.now.return_value = REAL_DATETIME(2026, 6, 20, 8, 30, 0)

        resolved = resolve_placeholders(
            "Reminder de Pagar el Agua del mes de [Mes anterior en letras] de [año en numero]"
        )

        self.assertEqual(resolved, "Reminder de Pagar el Agua del mes de Mayo de 2026")

    @patch("src.backend.date_utils.datetime.datetime")
    def test_resolve_placeholders_rolls_back_year_in_january(self, mock_datetime):
        mock_datetime.now.return_value = REAL_DATETIME(2026, 1, 5, 9, 0, 0)

        resolved = resolve_placeholders(
            "Recordatorio [Mes anterior en letras] [año en numero]"
        )

        self.assertEqual(resolved, "Recordatorio Diciembre 2025")


if __name__ == "__main__":
    unittest.main()