import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import load_workbook
from report import compare


class ReopenedOrderTest(unittest.TestCase):
    def test_only_previously_completed_order_counts_as_reopened(self):
        with TemporaryDirectory() as folder:
            root = Path(folder)
            previous = root / "previous.csv"
            current = root / "current.csv"
            output = root / "report.xlsx"

            previous.write_text("Order,Status\nA,Completed\nB,On Hold\n")
            current.write_text("Order,Status\nA,In Progress\nB,In Progress\n")

            compare(previous, current, output)

            book = load_workbook(output, read_only=True)
            try:
                rows = book["Orders"].iter_rows(min_row=2, values_only=True)
                changes = {row[0]: row[3] for row in rows}
                self.assertEqual(changes, {"A": "Reopened", "B": "Status changed"})

                counts = dict(book["Summary"].iter_rows(min_row=2, values_only=True))
                self.assertEqual(counts["Reopened"], 1)
            finally:
                book.close()
