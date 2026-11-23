"""Compare weekly service-order snapshots and write a review workbook."""
import argparse
import csv
from collections import Counter
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

ALIASES = ({"order", "order id", "customer order"}, {"status", "state"})


def read_rows(path):
    if path.suffix.lower() == ".csv":
        with path.open(encoding="utf-8-sig", newline="") as source:
            yield from csv.reader(source)
    elif path.suffix.lower() == ".xlsx":
        book = load_workbook(path, read_only=True, data_only=True)
        try:
            yield from book.active.values
        finally:
            book.close()
    else:
        raise ValueError(f"{path}: expected CSV or XLSX")


def snapshot(path):
    orders = {}
    rows = read_rows(path)
    headers = headers = [" ".join(str(value).lower().split()) for value in next(rows, ())]
    columns = []
    for aliases in ALIASES:
        matches = [i for i, name in enumerate(headers) if name in aliases]
        if len(matches) != 1:
            raise ValueError(f"{path}: expected one column matching {sorted(aliases)}")
        columns.append(matches[0])
    for line, row in enumerate(rows, 2):
        if not any(str(v or "").strip() for v in row):
            continue
        values = [str(row[i]).strip() if i < len(row) and row[i] is not None else "" for i in columns]
        order, status = values
        if not order or not status:
            raise ValueError(f"{path}, row {line}: missing order ID or status")
        if order in orders and orders[order].casefold() != status.casefold():
            raise ValueError(f"{path}, row {line}: conflicting statuses for {order}: {orders[order]!r} and {status!r}")
        orders[order] = status
    return orders


def compare(previous, current, output):
    before, after = snapshot(previous), snapshot(current)
    book = Workbook()
    summary = book.active
    summary.title = "Summary"
    detail = book.create_sheet("Orders")
    detail.append(["Order ID", "Previous status", "Current status", "Change"])
    counts = Counter()
    for order in sorted(before.keys() | after.keys()):
        old, new = before.get(order, ""), after.get(order, "")
        if not old:
            change = "Added"
        elif not new:
            change = "Missing this week"
        elif old.casefold() == new.casefold():
            change = "Unchanged"
        elif old.casefold() == "completed" and new.casefold() == "in progress":
            change = "Reopened"
        elif new.casefold() == "completed":
            change = "Completed"
        else:
            change = "Status changed"
        detail.append([order, old, new, change])
        for cell in detail[detail.max_row]:
            cell.data_type = "s"
        if change == "Reopened":
            detail.cell(detail.max_row, 4).fill = PatternFill("solid", fgColor="FFF2CC")
        elif change == "Missing this week":
            detail.cell(detail.max_row, 4).fill = PatternFill("solid", fgColor="FCE4D6")
        counts[change] += 1
    summary.append(["Change", "Orders"])
    for change in ["Added", "Completed", "Reopened", "Status changed", "Missing this week", "Unchanged"]:
        summary.append([change, counts[change]])
    for sheet in book:
        sheet.freeze_panes = "A2"
        for cell in sheet[1]:
            cell.font = Font(color="FFFFFF", bold=True)
            cell.fill = PatternFill("solid", fgColor="24516D")
            sheet.column_dimensions[cell.column_letter].width = 24
    detail.auto_filter.ref = detail.dimensions
    book.save(output)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("previous", type=Path)
    parser.add_argument("current", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    if args.output.resolve() in {args.previous.resolve(), args.current.resolve()} or args.output.suffix.lower() != ".xlsx":
        parser.error("Choose a separate .xlsx output file")
    try:
        compare(args.previous, args.current, args.output)
    except (ValueError, OSError) as error:
        parser.exit(1, f"{error}\n")
    print(f"Created {args.output}")
