#!/usr/bin/env python3
"""
Phase 1: Regional Mapping & Setup
Outputs a structured setup brief for a given region.
Usage: python tools/phase1_setup.py --region finger_lakes [--output outputs/]
"""

import json
import argparse
import sys
from pathlib import Path
from datetime import date

DATA_PATH = Path(__file__).parent.parent / "data" / "ny_regions.json"
OUTPUTS_PATH = Path(__file__).parent.parent / "outputs"


def load_regions():
    with open(DATA_PATH) as f:
        return json.load(f)["regions"]


def list_regions(regions):
    print("\nAvailable regions:")
    for key, r in regions.items():
        county_count = len(r["counties"])
        trail_count = len(r["trails"])
        print(f"  {key:30s} → {county_count} counties, {trail_count} known trails")
    print()


def build_setup_brief(region_key, regions):
    if region_key not in regions:
        print(f"ERROR: '{region_key}' not found. Run with --list to see available regions.")
        sys.exit(1)

    r = regions[region_key]
    brief = {
        "region": r["display_name"],
        "region_key": region_key,
        "generated": str(date.today()),
        "counties": [],
        "trail_map": [],
        "trail_gaps": [],
        "metro_proximity": r.get("metro_proximity", {}),
        "key_anchors": r.get("key_anchors", [])
    }

    # Build county list with DMO info
    counties_with_trails = set()
    for trail in r.get("trails", []):
        # Mark all counties as having trails (simplified — Claude will refine this)
        for c in r["counties"]:
            counties_with_trails.add(c["name"])
            break  # Just flag that the region has trails; per-county mapping is Claude's job

    for county in r["counties"]:
        county_entry = {
            "county": county["name"],
            "dmo_name": county["dmo"],
            "dmo_url": county["dmo_url"],
            "has_active_trail": None,  # Claude will research and fill this
            "trail_names": []          # Claude will research and fill this
        }
        brief["counties"].append(county_entry)

    # Trail list
    for trail in r.get("trails", []):
        brief["trail_map"].append({
            "trail_name": trail,
            "counties_served": [],   # Claude fills during research
            "has_passport": None,    # Claude fills during research
            "trail_url": None,       # Claude fills during research
            "member_count": None     # Claude fills during research
        })

    # Trail gaps = counties where no trail is currently active (Claude identifies these)
    brief["trail_gaps"] = {
        "note": "Claude: research each county above and flag any with no active tourism trail — these are RTN's primary partnership opportunities.",
        "gaps_identified": []
    }

    return brief


def output_brief(brief, output_dir):
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    slug = brief["region_key"]
    out_path = output_dir / f"setup-{slug}.json"
    with open(out_path, "w") as f:
        json.dump(brief, f, indent=2)

    print(f"\n{'='*60}")
    print(f"PHASE 1 SETUP BRIEF — {brief['region']}")
    print(f"{'='*60}")
    print(f"Generated: {brief['generated']}")
    print(f"Counties:  {len(brief['counties'])}")
    print(f"Known trails: {len(brief['trail_map'])}")
    print()
    print("COUNTIES & DMOs:")
    for c in brief["counties"]:
        print(f"  {c['county']:18s} | {c['dmo_name']}")
        print(f"  {'':18s}   {c['dmo_url']}")
    print()
    print("KNOWN TRAILS:")
    for t in brief["trail_map"]:
        print(f"  • {t['trail_name']}")
    print()
    print("METRO PROXIMITY:")
    for city, dist in brief["metro_proximity"].items():
        print(f"  {city}: {dist}")
    print()
    print("KEY ANCHORS:")
    for a in brief["key_anchors"]:
        print(f"  • {a}")
    print()
    print(f"NEXT STEPS FOR CLAUDE CODE:")
    print(f"  1. Verify and complete trail_map entries (URL, passport availability, member count)")
    print(f"  2. Identify trail_gaps — counties with no active trail")
    print(f"  3. Confirm DMO URLs are live and sourcing-ready")
    print(f"  4. Run: python tools/phase2_source.py --region {slug} --setup outputs/setup-{slug}.json")
    print()
    print(f"Brief saved → {out_path}")
    print(f"{'='*60}\n")

    return out_path


def main():
    parser = argparse.ArgumentParser(description="Phase 1: Generate regional setup brief")
    parser.add_argument("--region", "-r", help="Region key (e.g. finger_lakes, catskills)")
    parser.add_argument("--list", "-l", action="store_true", help="List all available regions")
    parser.add_argument("--output", "-o", default="outputs", help="Output directory")
    args = parser.parse_args()

    regions = load_regions()

    if args.list or not args.region:
        list_regions(regions)
        if not args.region:
            print("Use --region <key> to generate a setup brief.")
        return

    brief = build_setup_brief(args.region, regions)
    output_brief(brief, args.output)


if __name__ == "__main__":
    main()
