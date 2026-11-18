# Excel Report Toolkit

Compare two weekly service-order exports by order ID and see what changed. The example orders are fictional.

## Run the example

With Python 3 installed, run from the project directory:

```sh
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python report.py examples/week1.csv examples/week2.csv weekly-report.xlsx
```

Open weekly-report.xlsx. Summary shows counts. Orders shows the two statuses side by side. The example gives one order in each category.

## Use your own exports

Replace the first two arguments with last week's and this week's files. Both CSV and XLSX are supported. For Excel, the active sheet must contain the data. Put headers in the first row. Accepted ID headers are Order, Order ID and Customer Order. Accepted status headers are State and Status. Keep IDs with leading zeros as text.

Identical duplicates count once. Missing IDs, missing statuses and conflicting duplicate statuses produce an error. An order absent this week is marked missing. An existing order moving to Completed counts as completed. A newly seen order counts as added regardless of its status. Running again replaces the output workbook.
