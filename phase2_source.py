#!/usr/bin/env python3
"""
Phase 2: Listing Sourcer
Accepts structured listing data (JSON from Claude's DMO scrape) and produces
or appends to the regional listings XLSX in the RTN standard format.

Usage:
  python tools/phase2_source.py --region finger_lakes --input listings.json [--append]
  python tools/phase2_source.py --template  # Print the JSON input template

Input JSON format:
  {
    "region": "Finger Lakes",
    "county": "Schuyler",
    "dmo_url": "https://...",
    "listings": [
      {
        "business_name": "...",
        "address": "...",
        "city": "...",
        "county": "...",
        "state": "NY",
        "category": "craft_beverage",
        "trail_name": "Seneca Lake Wine Trail",
        "trail_member": true,
        "website": "...",
        "phone": "...",
        "notes": "...",
        "lat": null,
        "lng": null
      }
    ]
  }

Categories (6 standard):
  food_drink | culture_heritage | attractions | wellness | outdoor | agriculture
"""

import json
import argparse
import sys
from pathlib import Path
from datetime import date

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("ERROR: openpyxl not installed. Run: pip install openpyxl --break-system-packages")
    sys.exit(1)

OUTPUTS_PATH = Path(__file__).parent.parent / "outputs"

# RTN Standard column spec
COLUMNS = [
    ("Business Name",    38),
    ("Street Address",   30),
    ("Town/City",        20),
    ("County",           14),
    ("State",             6),
    ("Category",         24),
    ("Trail Member",     14),
    ("Trail Name",       34),
    ("Website",          34),
    ("Phone",            16),
    ("Notes",            55),
]

CATEGORY_LABELS = {
    "food_drink":        "Food & Drink",
    "culture_heritage":  "Culture & Heritage",
    "attractions":       "Attractions",
    "wellness":          "Wellness",
    "outdoor":           "Outdoor",
    "agriculture":       "Agriculture",
}

# Styling
HDR_FILL  = PatternFill("solid", fgColor="001F4E79")
HDR_FONT  = Font(name="Arial", bold=True, color="FFFFFFFF", size=11)
EVEN_FILL = PatternFill("solid", fgColor="00E2EFDA")
ODD_FILL  = PatternFill("solid", fgColor="00FFFFFF")
TRAIL_FILL= PatternFill("solid", fgColor="00C6EFCE")  # Highlighted trail members

THIN_BORDER = Border(
    left=Side(style="thin", color="FFCCCCCC"),
    right=Side(style="thin", color="FFCCCCCC"),
    bottom=Side(style="thin", color="FFCCCCCC")
)

TEMPLATE = {
    "region": "Finger Lakes",
    "county": "Schuyler",
    "dmo_url": "https://www.schuylerny.com",
    "scraped_date": str(date.today()),
    "listings": [
        {
            "business_name": "Example Winery",
            "address": "123 Seneca Lake Dr",
            "city": "Watkins Glen",
            "county": "Schuyler",
            "state": "NY",
            "category": "food_drink",
            "trail_name": "Seneca Lake Wine Trail",
            "trail_member": True,
            "website": "examplewinery.com",
            "phone": "607-555-0100",
            "notes": "Estate winery on Seneca Lake's eastern shore. Known for Riesling and dry rosé."
        }
    ]
}


def get_or_create_workbook(path):
    if path.exists():
        wb = openpyxl.load_workbook(path)
        ws = wb["Listings"] if "Listings" in wb.sheetnames else wb.active
        summary_ws = wb["Summary"] if "Summary" in wb.sheetnames else None
        return wb, ws, summary_ws, True
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Listings"
        summary_ws = wb.create_sheet("Summary")
        return wb, ws, summary_ws, False


def write_header(ws):
    for col_idx, (header, width) in enumerate(COLUMNS, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.row_dimensions[1].height = 20
    ws.freeze_panes = "A2"


def setup_summary_sheet(ws):
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 28
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 44

    # Title row
    ws.merge_cells("A1:E1")
    title_cell = ws["A1"]
    title_cell.value = "Upstate NY Tourism Listings — Regional Summary"
    title_cell.font = Font(name="Arial", bold=True, color="FFFFFFFF", size=13)
    title_cell.fill = HDR_FILL
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24

    # Sub-header
    headers = ["County", "# Listings", "Trail Members", "Date Added", "DMO Source"]
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col_idx, value=h)
        cell.font = Font(name="Arial", bold=True, size=10)
        cell.fill = PatternFill("solid", fgColor="00DDEBF7")
        cell.alignment = Alignment(horizontal="center", vertical="center")


def write_listing_row(ws, row_num, listing):
    trail_member = listing.get("trail_member", False)
    fill = TRAIL_FILL if trail_member else (EVEN_FILL if row_num % 2 == 0 else ODD_FILL)

    category_key = listing.get("category", "").lower()
    category_label = CATEGORY_LABELS.get(category_key, listing.get("category", ""))

    values = [
        listing.get("business_name", ""),
        listing.get("address", ""),
        listing.get("city", ""),
        listing.get("county", ""),
        listing.get("state", "NY"),
        category_label,
        "✓ Trail Member" if trail_member else "",
        listing.get("trail_name", ""),
        listing.get("website", ""),
        listing.get("phone", ""),
        listing.get("notes", ""),
    ]

    for col_idx, value in enumerate(values, 1):
        cell = ws.cell(row=row_num, column=col_idx, value=value)
        cell.font = Font(name="Arial", size=10)
        cell.fill = fill
        cell.border = THIN_BORDER
        cell.alignment = Alignment(
            vertical="center",
            wrap_text=(col_idx == len(COLUMNS))
        )
        if col_idx == len(COLUMNS):  # Notes column
            ws.row_dimensions[row_num].height = 30


def add_summary_row(summary_ws, county, listing_count, trail_count, dmo_url):
    next_row = summary_ws.max_row + 1

    palette = ["00E2EFDA", "00C6EFCE", "00FFF2CC", "00FCE4D6", "00DDEBF7", "00F2F2F2"]
    fill_color = palette[(next_row - 3) % len(palette)]
    fill = PatternFill("solid", fgColor=fill_color)

    values = [county, listing_count, trail_count, str(date.today()), dmo_url]
    for col_idx, value in enumerate(values, 1):
        cell = summary_ws.cell(row=next_row, column=col_idx, value=value)
        cell.font = Font(name="Arial", size=10)
        cell.fill = fill
        cell.alignment = Alignment(vertical="center")


def process_listings(input_data, output_path, append_mode):
    wb, ws, summary_ws, is_existing = get_or_create_workbook(output_path)

    if not is_existing:
        write_header(ws)
        if summary_ws:
            setup_summary_sheet(summary_ws)

    # Determine starting row
    start_row = ws.max_row + 1

    listings = input_data.get("listings", [])
    county = input_data.get("county", "Unknown")
    dmo_url = input_data.get("dmo_url", "")
    region = input_data.get("region", "")

    trail_members = 0
    for i, listing in enumerate(listings):
        row_num = start_row + i
        write_listing_row(ws, row_num, listing)
        if listing.get("trail_member"):
            trail_members += 1

    # Add summary row
    if summary_ws:
        add_summary_row(summary_ws, county, len(listings), trail_members, dmo_url)

    wb.save(output_path)
    return len(listings), trail_members


def main():
    parser = argparse.ArgumentParser(description="Phase 2: Build listings XLSX from Claude-scraped data")
    parser.add_argument("--region", "-r", help="Region key (used for output filename if --output not set)")
    parser.add_argument("--input", "-i", help="Path to JSON file with listing data")
    parser.add_argument("--output", "-o", help="Output XLSX path (default: outputs/<region>-listings.xlsx)")
    parser.add_argument("--append", "-a", action="store_true", help="Append to existing file")
    parser.add_argument("--template", action="store_true", help="Print the expected JSON input format")
    parser.add_argument("--validate", "-v", action="store_true", help="Validate JSON input only, no output")
    args = parser.parse_args()

    if args.template:
        print(json.dumps(TEMPLATE, indent=2))
        return

    if not args.input:
        print("ERROR: --input <file.json> required. Use --template to see the expected format.")
        sys.exit(1)

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"ERROR: Input file not found: {input_path}")
        sys.exit(1)

    with open(input_path) as f:
        input_data = json.load(f)

    # Validate
    if "listings" not in input_data:
        print("ERROR: JSON must contain a 'listings' array. Use --template to see format.")
        sys.exit(1)

    valid_cats = set(CATEGORY_LABELS.keys())
    issues = []
    for i, listing in enumerate(input_data["listings"]):
        if not listing.get("business_name"):
            issues.append(f"  Row {i+1}: missing business_name")
        if listing.get("category", "").lower() not in valid_cats:
            issues.append(f"  Row {i+1}: invalid category '{listing.get('category')}' (valid: {', '.join(valid_cats)})")

    if issues:
        print(f"VALIDATION ISSUES ({len(issues)} found):")
        for issue in issues:
            print(issue)
        if args.validate:
            return
        print("\nProceeding anyway — fix issues before Tourismo import.\n")

    if args.validate:
        print(f"Validation complete: {len(input_data['listings'])} listings, {len(issues)} issues.")
        return

    # Determine output path
    OUTPUTS_PATH.mkdir(parents=True, exist_ok=True)
    region_slug = args.region or input_data.get("region", "region").lower().replace(" ", "_")
    output_path = Path(args.output) if args.output else OUTPUTS_PATH / f"{region_slug}-listings.xlsx"

    # Process
    count, trail_count = process_listings(input_data, output_path, args.append)

    print(f"\n{'='*55}")
    print(f"PHASE 2 — LISTINGS ADDED")
    print(f"{'='*55}")
    print(f"Region:        {input_data.get('region', 'Unknown')}")
    print(f"County:        {input_data.get('county', 'Unknown')}")
    print(f"Listings:      {count} added")
    print(f"Trail members: {trail_count} (highlighted green)")
    print(f"Output:        {output_path}")
    print(f"\nNEXT: Run phase2_source.py for the next county, or")
    print(f"      python tools/phase3_analyze.py --region {region_slug}")
    print(f"{'='*55}\n")


if __name__ == "__main__":
    main()
