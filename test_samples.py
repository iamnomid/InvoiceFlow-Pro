"""Offline tests for V4; requires pymupdf, pandas and openpyxl."""
from io import BytesIO
from pathlib import Path
import pandas as pd
from app import pdf_to_text, parse_invoice_text, detect_potential_duplicates, make_excel

folder = Path(__file__).parent / 'samples'
files = sorted(folder.glob('*.pdf'))
assert len(files) == 7

def parse(names):
    return detect_potential_duplicates([
        parse_invoice_text(name, pdf_to_text((folder / name).read_bytes())) for name in names
    ])

six = [file.name for file in files if 'DUPLICATE' not in file.name]
base = parse(six)
assert sum(r['status'] == 'Complete' for r in base) == 4
assert sum(r['status'] == 'Needs review' for r in base) == 2
assert sum(r['total_usd'] or 0 for r in base if r['status'] == 'Complete') == 598.5

all_rows = parse([file.name for file in files])
assert sum(r['status'] == 'Complete' for r in all_rows) == 3
assert sum(r['status'] == 'Needs review' for r in all_rows) == 4
assert sum('Potential duplicate' in r['review_reason'] for r in all_rows) == 2
assert sum(r['total_usd'] or 0 for r in all_rows if r['status'] == 'Complete') == 473.5

from openpyxl import load_workbook
wb = load_workbook(BytesIO(make_excel(pd.DataFrame(all_rows))))
assert wb.sheetnames == ['Invoices', 'Review Needed', 'Summary']
summary = {row[0]: row[1] for row in wb['Summary'].values if row[0] not in (None, 'Metric')}
assert summary['Files processed'] == 7
assert summary['Complete'] == 3
assert summary['Needs review'] == 4
assert summary['Verified total (USD)'] == 473.5
assert wb['Review Needed'].max_row == 5
print('PASS: 6 file baseline and 7 file duplicate case; Excel report 3 sheets')
