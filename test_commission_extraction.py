import os
import sys
import json
import glob
from pathlib import Path
sys.path.insert(0, '.')
from dotenv import load_dotenv
load_dotenv()

from parser.commission_extractor import process_commission_file

# Commission samples folder in project
samples_dir = os.path.join(
    os.path.dirname(__file__),
    'test_data', 'commission_samples'
)

# Also check Downloads as fallback
downloads = os.path.expanduser('~/Downloads')

# Find all commission files
test_files = []

# Check project samples folder first
if os.path.exists(samples_dir):
    for ext in ['*.pdf', '*.xlsx',
                '*.xls', '*.txt', '*.csv']:
        test_files.extend(
            glob.glob(os.path.join(samples_dir, ext))
        )

# Fallback to Downloads
if not test_files:
    for ext in ['*.pdf', '*.xlsx',
                '*.xls', '*.txt']:
        for f in glob.glob(
            os.path.join(downloads, ext)
        ):
            name = os.path.basename(f).lower()
            if any(kw in name for kw in [
                'comm', 'cambro', 'dormont',
                'follett', 'southbend', 'blodgett',
                'star', 'blendtec', 'metalcraft'
            ]):
                test_files.append(f)

# Use project db for templates
db_path = os.path.join(
    os.path.dirname(__file__),
    'storage', 'workbooks.db'
)

print(f"Found {len(test_files)} files to test")
for f in test_files:
    print(f"  {os.path.basename(f)}")
print()

for file_path in test_files:
    print(f"\n{'='*60}")
    result = process_commission_file(
        file_path,
        db_path=db_path
    )
    
    if result['status'] == 'error':
        print(f"ERROR: {result['error']}")
        continue
    
    if result['status'] == 'duplicate':
        print("DUPLICATE - already processed")
        continue
    
    summary = result['summary']
    print(f"\nRESULTS:")
    print(f"  Manufacturer: {summary['manufacturer']}")
    print(f"  Period: {summary['period']}")
    print(f"  Total POs: {summary['total_pos']}")
    print(f"  Total commission: ${summary['total_commission']:,.2f}")
    print(f"  Items needing review: {summary['items_needing_review']}")
    print(f"  No-PO items: {summary['total_no_po']}")
    
    # Show first 3 grouped POs
    print(f"\n  First 3 PO groups:")
    grouped = result['grouped']
    for i, (po_norm, group) in enumerate(
        list(grouped.items())[:3]
    ):
        print(f"    PO: {group['po_number']}")
        print(f"      Dealer: {group['dealer_name']}")
        print(f"      Commission: ${group['total_commission']:,.2f}")
        print(f"      Line items: {group['line_item_count']}")
        print(f"      Needs review: {group['needs_review']}")
    
    # Show validation issues summary
    all_issues = []
    for group in grouped.values():
        all_issues.extend(group['all_issues'])
    
    if all_issues:
        from collections import Counter
        issue_counts = Counter(all_issues)
        print(f"\n  Validation issues found:")
        for issue, count in issue_counts.most_common(5):
            print(f"    {issue}: {count}")

print(f"\n{'='*60}")
print("Test complete.")
