import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import Workbook
from report import snapshot


class HeaderTests(unittest.TestCase):
    def test_multiline_excel_header_preserves_order_id(self):
        with TemporaryDirectory() as folder:
            path = Path(folder) / "orders.xlsx"
            book = Workbook()
            book.active.append([" Customer\n  Order ", " STATE "])
            book.active.append(["001", "Open"])
            book.save(path)
            book.close()

            self.assertEqual(snapshot(path), {"001": "Open"})
