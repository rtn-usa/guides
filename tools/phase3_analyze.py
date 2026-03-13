#!/usr/bin/env python3
"""
Phase 3 — Destination Analysis

Generates the analysis template and reviews completeness.

Usage:
    python tools/phase3_analyze.py --region <key>
    python tools/phase3_analyze.py --review <analysis_json>

Four workstreams:
1. Assets Inventory — natural features, heritage, culinary identity, craft beverages
2. Infrastructure Assessment — lodging, transport, visitor centers, weekend viability
3. UVP Synthesis — identity statement, visitor archetypes, RTN pitch
4. Seasonality Profile — peak, shoulder, off-season, anchor events
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


def generate_analysis_template(region_key):
    """Generate the analysis JSON with null fields for Claude to fill."""
    data = load_regions()
    if region_key not in data["regions"]:
        print(f"Error: Region '{region_key}' not found.", file=sys.stderr)
        sys.exit(1)

    region = data["regions"][region_key]

    analysis = {
        "region_key": region_key,
        "region_name": region["name"],
        "workstream_1_assets": {
            "natural_features": {
                "major_peaks": None,
                "lakes_rivers": None,
                "gorges_falls": None,
                "state_parks": None,
                "unique_geology": None,
                "notes": None
            },
            "heritage_culture": {
                "historical_narrative": None,
                "key_sites": None,
                "indigenous_history": None,
                "industrial_legacy": None,
                "arts_scene": None,
                "notes": None
            },
            "culinary_identity": {
                "signature_products": None,
                "national_reputation": None,
                "farm_to_table_scene": None,
                "food_festivals": None,
                "notes": None
            },
            "craft_beverage": {
                "winery_count": None,
                "brewery_count": None,
                "distillery_count": None,
                "cidery_count": None,
                "flagship_producers": None,
                "ava_designations": None,
                "notes": None
            }
        },
        "workstream_2_infrastructure": {
            "lodging": {
                "types_available": None,
                "approximate_capacity": None,
                "flagship_properties": None,
                "airbnb_presence": None,
                "camping_options": None,
                "notes": None
            },
            "transport": {
                "drive_times": region.get("drive_times", {}),
                "drive_times_verified": None,
                "major_routes": None,
                "amtrak": region.get("amtrak", {}),
                "amtrak_verified": None,
                "bus_service": None,
                "airport_proximity": None,
                "notes": None
            },
            "visitor_services": {
                "visitor_centers": None,
                "distribution_quality": None,
                "cell_coverage": None,
                "ev_charging": None,
                "notes": None
            },
            "weekend_viability": {
                "can_plan_fri_sun": None,
                "minimum_nights_recommended": None,
                "best_base_towns": None,
                "itinerary_density": None,
                "notes": None
            }
        },
        "workstream_3_uvp": {
            "identity_statement": None,
            "destination_or_corridor": None,
            "destination_type_explanation": None,
            "differentiation_from_neighbors": None,
            "visitor_archetypes": [
                {
                    "archetype": None,
                    "description": None,
                    "why_here": None
                }
            ],
            "rtn_pitch": None,
            "competitive_position": None,
            "notes": None
        },
        "workstream_4_seasonality": {
            "peak_season": {
                "months": None,
                "drivers": None,
                "crowd_level": None,
                "price_premium": None,
                "notes": None
            },
            "shoulder_season": {
                "months": None,
                "specific_opportunity": None,
                "undervisited_assets": None,
                "editorial_angle": None,
                "notes": None
            },
            "off_season": {
                "months": None,
                "anchors": None,
                "winter_activities": None,
                "notes": None
            },
            "anchor_events": [
                {
                    "name": None,
                    "month": None,
                    "description": None,
                    "visitor_draw": None
                }
            ]
        }
    }

    os.makedirs(OUTPUTS_DIR, exist_ok=True)
    output_path = os.path.join(OUTPUTS_DIR, f"analysis-{region_key}.json")

    with open(output_path, "w") as f:
        json.dump(analysis, f, indent=2)

    print(f"Phase 3 Analysis Template: {region['name']}")
    print("=" * 60)
    print(f"Output: {output_path}")
    print()
    print("WORKSTREAMS TO COMPLETE:")
    print()
    print("1. ASSETS INVENTORY")
    print("   - Natural features (peaks, lakes, gorges, parks, geology)")
    print("   - Heritage & culture (history, key sites, arts)")
    print("   - Culinary identity (signature products, farm-to-table)")
    print("   - Craft beverage (counts, flagships, AVAs)")
    print()
    print("2. INFRASTRUCTURE ASSESSMENT")
    print("   - Lodging (types, capacity, flagships)")
    print("   - Transport (verify drive times, routes, Amtrak, bus)")
    print("   - Visitor services (centers, cell coverage, EV)")
    print("   - Weekend viability (Fri-Sun trip feasibility)")
    print()
    print("3. UVP SYNTHESIS")
    print("   - Identity statement (one singular sentence)")
    print("   - Destination vs. corridor classification")
    print("   - Differentiation from neighbors")
    print("   - Visitor archetypes (specific, not generic)")
    print("   - RTN pitch for DMOs")
    print()
    print("4. SEASONALITY PROFILE")
    print("   - Peak season (when, drivers, crowd level)")
    print("   - Shoulder season (THE editorial priority)")
    print("   - Off-season anchors")
    print("   - 3-5 named anchor events")
    print()
    print("After completing, run:")
    print(f"  python tools/phase3_analyze.py --review {output_path}")


def count_fields(obj, prefix=""):
    """Recursively count total and filled fields."""
    total = 0
    filled = 0

    if isinstance(obj, dict):
        for key, value in obj.items():
            if key in ("region_key", "region_name"):
                continue
            if isinstance(value, (dict, list)):
                t, f = count_fields(value, f"{prefix}.{key}")
                total += t
                filled += f
            else:
                total += 1
                if value is not None:
                    filled += 1
    elif isinstance(obj, list):
        for i, item in enumerate(obj):
            t, f = count_fields(item, f"{prefix}[{i}]")
            total += t
            filled += f

    return total, filled


def review_analysis(filepath):
    """Review an analysis JSON for completeness."""
    if not os.path.exists(filepath):
        print(f"Error: File not found: {filepath}", file=sys.stderr)
        sys.exit(1)

    with open(filepath, "r") as f:
        analysis = json.load(f)

    region_name = analysis.get("region_name", "Unknown")
    print(f"Phase 3 Review: {region_name}")
    print("=" * 60)

    # Check each workstream
    workstreams = {
        "workstream_1_assets": "Assets Inventory",
        "workstream_2_infrastructure": "Infrastructure Assessment",
        "workstream_3_uvp": "UVP Synthesis",
        "workstream_4_seasonality": "Seasonality Profile"
    }

    overall_total = 0
    overall_filled = 0

    for ws_key, ws_name in workstreams.items():
        ws_data = analysis.get(ws_key, {})
        total, filled = count_fields(ws_data)
        overall_total += total
        overall_filled += filled
        pct = round(filled / total * 100, 1) if total else 0
        status = "COMPLETE" if pct >= 80 else "NEEDS WORK" if pct >= 50 else "INCOMPLETE"
        print(f"\n{ws_name}: {filled}/{total} fields ({pct}%) — {status}")

        # Show null fields
        null_fields = find_null_fields(ws_data)
        if null_fields:
            for field in null_fields[:5]:
                print(f"  Missing: {field}")
            if len(null_fields) > 5:
                print(f"  ... and {len(null_fields) - 5} more")

    overall_pct = round(overall_filled / overall_total * 100, 1) if overall_total else 0
    print()
    print(f"OVERALL: {overall_filled}/{overall_total} fields ({overall_pct}%)")

    if overall_pct >= 80:
        print("STATUS: Ready for Phase 4")
    elif overall_pct >= 50:
        print("STATUS: Needs more research — focus on missing fields above")
    else:
        print("STATUS: Incomplete — significant research needed")


def find_null_fields(obj, prefix=""):
    """Find all null fields in a nested structure."""
    nulls = []
    if isinstance(obj, dict):
        for key, value in obj.items():
            path = f"{prefix}.{key}" if prefix else key
            if isinstance(value, (dict, list)):
                nulls.extend(find_null_fields(value, path))
            elif value is None:
                nulls.append(path)
    elif isinstance(obj, list):
        for i, item in enumerate(obj):
            nulls.extend(find_null_fields(item, f"{prefix}[{i}]"))
    return nulls


def main():
    parser = argparse.ArgumentParser(description="Phase 3: Destination Analysis")
    parser.add_argument("--region", type=str, help="Region key to generate template")
    parser.add_argument("--review", type=str, help="Path to analysis JSON to review")
    args = parser.parse_args()

    if args.review:
        review_analysis(args.review)
        return

    if not args.region:
        parser.print_help()
        sys.exit(1)

    generate_analysis_template(args.region)


if __name__ == "__main__":
    main()
