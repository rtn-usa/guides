#!/usr/bin/env python3
"""
Phase 1 — Regional Mapping & Setup

Generates the initial setup brief for a region, including:
- Region metadata (counties, base towns, drive times)
- DMO list with URLs to verify
- Trail data with county mapping placeholders
- Trail gap identification placeholders

Usage:
    python tools/phase1_setup.py --region <key>
    python tools/phase1_setup.py --list

After running, Claude should:
1. Verify each DMO URL is live
2. Research trail-to-county mapping
3. Flag counties with no active trail (gap counties)
4. Add trail URLs and passport availability
5. Save completed brief to outputs/setup-<region>.json
"""

import argparse
import json
import os
import sys

REGIONS_FILE = os.path.join(os.path.dirname(__file__), "..", "data", "ny_regions.json")
OUTPUTS_DIR = os.path.join(os.path.dirname(__file__), "..", "outputs")


def load_regions():
    with open(REGIONS_FILE, "r") as f:
        return json.load(f)


def list_regions(data):
    print("Available regions:")
    print("-" * 50)
    for key, region in sorted(data["regions"].items()):
        county_count = len(region["counties"])
        trail_count = len(region.get("trails", []))
        print(f"  {key:<25} {region['name']:<25} ({county_count} counties, {trail_count} trails)")
    print()
    print(f"Total regions: {len(data['regions'])}")


def generate_setup_brief(data, region_key):
    if region_key not in data["regions"]:
        print(f"Error: Region '{region_key}' not found.", file=sys.stderr)
        print(f"Available: {', '.join(sorted(data['regions'].keys()))}", file=sys.stderr)
        sys.exit(1)

    region = data["regions"][region_key]

    # Build trail-to-county mapping with verification placeholders
    trails_brief = []
    counties_with_trails = set()

    for trail in region.get("trails", []):
        trail_counties = trail.get("counties", [])
        counties_with_trails.update(trail_counties)
        trails_brief.append({
            "name": trail["name"],
            "type": trail["type"],
            "counties": trail_counties,
            "website": trail.get("website"),
            "website_verified": None,  # Claude fills: true/false
            "stops_approx": trail.get("stops_approx"),
            "passport": trail.get("passport"),
            "passport_verified": None,  # Claude fills: true/false
            "member_directory_url": None,  # Claude fills
            "notes": None  # Claude fills
        })

    # Identify gap counties
    all_counties = set(region["counties"])
    gap_counties = all_counties - counties_with_trails
    covered_counties = all_counties & counties_with_trails

    gap_county_briefs = []
    for county in sorted(gap_counties):
        gap_county_briefs.append({
            "county": county,
            "has_trail": False,
            "potential_trail_type": None,  # Claude fills
            "existing_producers": None,  # Claude fills
            "dmo_partner": None,  # Claude fills
            "notes": None  # Claude fills
        })

    # DMO verification list
    dmos_brief = []
    for dmo in region.get("dmos", []):
        dmos_brief.append({
            "name": dmo["name"],
            "url": dmo["url"],
            "url_verified": None,  # Claude fills: true/false
            "counties_served": dmo["counties_served"],
            "has_listings_page": None,  # Claude fills: true/false
            "listings_url": None,  # Claude fills
            "notes": None  # Claude fills
        })

    # Assemble the brief
    brief = {
        "region_key": region_key,
        "region_name": region["name"],
        "description": region["description"],
        "counties": {
            "all": sorted(region["counties"]),
            "with_trails": sorted(counties_with_trails),
            "gap_counties": sorted(gap_counties),
            "total": len(region["counties"]),
            "coverage_pct": round(len(covered_counties) / len(all_counties) * 100, 1) if all_counties else 0
        },
        "base_towns": region["base_towns"],
        "drive_times": region["drive_times"],
        "amtrak": region.get("amtrak", {}),
        "trails": trails_brief,
        "trail_summary": {
            "total": len(trails_brief),
            "with_passport": sum(1 for t in trails_brief if t.get("passport")),
            "with_website": sum(1 for t in trails_brief if t.get("website")),
            "types": sorted(set(t["type"] for t in trails_brief))
        },
        "gap_counties": gap_county_briefs,
        "dmos": dmos_brief,
        "phase1_status": {
            "dmo_urls_verified": False,
            "trails_mapped_to_counties": False,
            "gap_counties_analyzed": False,
            "trail_urls_verified": False,
            "completed": False
        }
    }

    return brief


def main():
    parser = argparse.ArgumentParser(description="Phase 1: Regional Mapping & Setup")
    parser.add_argument("--region", type=str, help="Region key (e.g., finger_lakes)")
    parser.add_argument("--list", action="store_true", help="List all available regions")
    args = parser.parse_args()

    data = load_regions()

    if args.list:
        list_regions(data)
        return

    if not args.region:
        parser.print_help()
        print("\nUse --list to see available regions.")
        sys.exit(1)

    brief = generate_setup_brief(data, args.region)

    # Ensure outputs directory exists
    os.makedirs(OUTPUTS_DIR, exist_ok=True)

    # Print the brief summary
    print(f"Phase 1 Setup Brief: {brief['region_name']}")
    print("=" * 60)
    print(f"Counties: {brief['counties']['total']} total, {len(brief['counties']['with_trails'])} with trails, {len(brief['counties']['gap_counties'])} gaps")
    print(f"Trails: {brief['trail_summary']['total']} ({brief['trail_summary']['with_passport']} with passport)")
    print(f"DMOs: {len(brief['dmos'])}")
    print()

    if brief["counties"]["gap_counties"]:
        print("GAP COUNTIES (no active trail):")
        for county in brief["counties"]["gap_counties"]:
            print(f"  - {county}")
        print()

    print("DMO URLs to verify:")
    for dmo in brief["dmos"]:
        print(f"  - {dmo['name']}: {dmo['url']}")
    print()

    print("Trail websites to verify:")
    for trail in brief["trails"]:
        url = trail["website"] or "(no URL known)"
        print(f"  - {trail['name']}: {url}")
    print()

    # Save the brief
    output_path = os.path.join(OUTPUTS_DIR, f"setup-{args.region}.json")
    with open(output_path, "w") as f:
        json.dump(brief, f, indent=2)
    print(f"Brief saved to: {output_path}")
    print()
    print("NEXT STEPS for Claude:")
    print("1. Fetch each DMO URL to verify it's live")
    print("2. Research trail-to-county mapping at county level")
    print("3. Flag gap counties and assess partnership potential")
    print("4. Verify trail URLs and passport availability")
    print(f"5. Save completed brief to {output_path}")


if __name__ == "__main__":
    main()
