#!/usr/bin/env python3
"""
Phase 3: Destination Analysis
Generates a structured analysis template for a region. Claude Code fills in
all four workstreams through research, then saves the completed JSON.

Usage:
  python tools/phase3_analyze.py --region finger_lakes [--setup outputs/setup-finger_lakes.json]
  python tools/phase3_analyze.py --review outputs/analysis-finger_lakes.json  # Score completeness
"""

import json
import argparse
import sys
from pathlib import Path
from datetime import date

DATA_PATH = Path(__file__).parent.parent / "data" / "ny_regions.json"
OUTPUTS_PATH = Path(__file__).parent.parent / "outputs"


def load_region(region_key):
    with open(DATA_PATH) as f:
        regions = json.load(f)["regions"]
    if region_key not in regions:
        print(f"ERROR: region '{region_key}' not found.")
        sys.exit(1)
    return regions[region_key]


def load_setup_brief(setup_path):
    if setup_path and Path(setup_path).exists():
        with open(setup_path) as f:
            return json.load(f)
    return None


def generate_template(region_key, region_data, setup_brief):
    """Produce the blank analysis template Claude will fill in."""

    county_names = [c["name"] for c in region_data.get("counties", [])]
    known_trails = region_data.get("trails", [])
    metro_proximity = region_data.get("metro_proximity", {})
    key_anchors = region_data.get("key_anchors", [])

    template = {
        "meta": {
            "region": region_data["display_name"],
            "region_key": region_key,
            "generated": str(date.today()),
            "analysis_status": "DRAFT — Claude to complete all sections below",
            "completeness_score": 0
        },

        # ── WORKSTREAM 1: ASSETS INVENTORY ─────────────────────────────────
        "assets": {
            "_instructions": "Research and document the region's natural, cultural, culinary, and craft beverage assets. Be specific — name the mountains, the foods, the producers, the history.",
            "natural_anchors": {
                "_instructions": "Major natural features that draw visitors. Name specific peaks, lakes, gorges, waterfalls, parks.",
                "known_from_data": key_anchors,
                "researched": []
            },
            "heritage_and_culture": {
                "_instructions": "Historic sites, museums, notable events, Indigenous history, immigration heritage, industrial legacy.",
                "researched": []
            },
            "culinary_identity": {
                "_instructions": "What food or agricultural product is this region KNOWN for? Not just 'farm-to-table' — what specific thing? (e.g., Finger Lakes = Riesling; Catskills = trout; Hudson Valley = foie gras era but also apple orchards)",
                "signature_products": [],
                "notable_producers": [],
                "culinary_reputation_nationally": None
            },
            "craft_beverage_profile": {
                "_instructions": "Wine, beer, spirits, cider — how developed is the scene? Names of notable producers. Regional styles.",
                "wine_trails": [],
                "brewery_cidery_scene": None,
                "spirits": None,
                "flagship_producers": []
            }
        },

        # ── WORKSTREAM 2: INFRASTRUCTURE ASSESSMENT ─────────────────────────
        "infrastructure": {
            "_instructions": "Can a visitor actually spend 2+ nights here? How do they get there? What visitor services exist?",
            "lodging": {
                "_instructions": "Range and volume. Not just 'there are hotels' — how many rooms, what types (B&B, inn, resort, camping, glamping), and what price tier?",
                "lodging_types_present": [],
                "approx_room_count": None,
                "price_range": None,
                "flagship_properties": [],
                "gaps_or_weaknesses": None
            },
            "access": {
                "_instructions": "How do visitors actually get here? What's the drive from NYC, Buffalo, Albany, Toronto? Is there Amtrak or bus access?",
                "drive_times": metro_proximity,
                "amtrak_served": None,
                "bus_access": None,
                "nearest_major_airport": None,
                "primary_arrival_route": None
            },
            "visitor_infrastructure": {
                "visitor_centers": [],
                "wayfinding_quality": None,
                "digital_presence": None,
                "passport_program_active": None,
                "trail_coordination_body": None
            },
            "weekend_viability": {
                "_instructions": "Can a visitor plan a Fri-Sun trip and have enough to fill two days? Or is this a corridor stop?",
                "verdict": None,
                "reasoning": None,
                "anchor_experience_for_overnight": None
            }
        },

        # ── WORKSTREAM 3: UVP SYNTHESIS ─────────────────────────────────────
        "uvp": {
            "_instructions": "This is the heart of the guide. What does this region have that no other NY region does? What's the identity? What's the RTN pitch to DMOs?",
            "singular_identity": {
                "_instructions": "One clear, specific sentence that captures what makes this region irreplaceable. No generic language. No 'something for everyone.'",
                "statement": None,
                "supporting_evidence": []
            },
            "destination_vs_corridor": {
                "_instructions": "Is this region a destination in its own right, or is it primarily a corridor stop (e.g., 'on the way from Albany to Montreal')?",
                "verdict": None,
                "corridor_routes": [],
                "reasoning": None
            },
            "competitive_differentiation": {
                "_instructions": "How is this different from neighboring regions? What would you miss if you skipped it?",
                "vs_neighboring_regions": {},
                "unique_assets": []
            },
            "visitor_archetypes": {
                "_instructions": "Who actually comes here and why? 2-3 specific visitor types with their motivations.",
                "primary": None,
                "secondary": None,
                "emerging": None
            },
            "rtn_partnership_pitch": {
                "_instructions": "What would you say to a DMO in this region to get them to join the Upstate network? What's their gain?",
                "key_message": None,
                "gaps_we_solve": [],
                "comparable_success": None
            }
        },

        # ── WORKSTREAM 4: SEASONALITY PROFILE ───────────────────────────────
        "seasonality": {
            "_instructions": "When should visitors come, and for what? The shoulder season is the most underexploited — what's the opportunity?",
            "peak_season": {
                "months": [],
                "primary_draws": [],
                "crowd_level": None,
                "pricing_pressure": None
            },
            "shoulder_season": {
                "_instructions": "This is Upstate's biggest editorial opportunity — when is it, and what's there to do?",
                "months": [],
                "opportunity": None,
                "what_makes_it_worth_it": [],
                "under_visited_reason": None
            },
            "off_season": {
                "months": [],
                "anchors": [],
                "viable": None
            },
            "signature_events": {
                "_instructions": "Named annual events that anchor visits — festivals, races, openings, harvests.",
                "events": []
            },
            "seasonal_guide_headline": {
                "_instructions": "One line per season that would work as an Upstate editorial callout box.",
                "spring": None,
                "summer": None,
                "fall": None,
                "winter": None
            }
        },

        # ── COUNTY-LEVEL TRAIL MAP ───────────────────────────────────────────
        "county_trail_map": {
            "_instructions": "For each county, list the active trails. Flag counties with NO trail as priority RTN outreach targets.",
            "counties": {
                county: {
                    "active_trails": [],
                    "trail_gap": None,
                    "gap_opportunity": None
                }
                for county in county_names
            }
        },

        # ── GUIDE FRAMING NOTES ──────────────────────────────────────────────
        "guide_framing": {
            "_instructions": "Quick notes for the Phase 4 guide writer. What's the emotional hook? What's the lede?",
            "emotional_hook": None,
            "lede_angle": None,
            "what_to_avoid": [],
            "comparable_editorial_tone": None
        }
    }

    return template


def score_completeness(analysis):
    """Count how many null/empty fields remain. Returns 0-100."""
    total_fields = 0
    filled_fields = 0

    def walk(obj):
        nonlocal total_fields, filled_fields
        if isinstance(obj, dict):
            for k, v in obj.items():
                if k.startswith("_"):  # Skip instruction fields
                    continue
                if isinstance(v, (dict, list)):
                    walk(v)
                else:
                    total_fields += 1
                    if v is not None and v != "" and v != []:
                        filled_fields += 1
        elif isinstance(obj, list):
            for item in obj:
                walk(item)

    walk(analysis)
    return int((filled_fields / total_fields * 100) if total_fields > 0 else 0)


def review_analysis(analysis_path):
    with open(analysis_path) as f:
        analysis = json.load(f)

    score = score_completeness(analysis)
    region = analysis.get("meta", {}).get("region", "Unknown")

    print(f"\n{'='*55}")
    print(f"PHASE 3 REVIEW — {region}")
    print(f"{'='*55}")
    print(f"Completeness: {score}%")

    # Check each workstream
    workstreams = {
        "Assets Inventory":           analysis.get("assets", {}),
        "Infrastructure Assessment":  analysis.get("infrastructure", {}),
        "UVP Synthesis":              analysis.get("uvp", {}),
        "Seasonality Profile":        analysis.get("seasonality", {}),
    }

    print("\nWorkstream status:")
    for name, data in workstreams.items():
        ws_score = score_completeness(data)
        status = "✓ READY" if ws_score >= 80 else ("⚠ PARTIAL" if ws_score >= 40 else "✗ INCOMPLETE")
        print(f"  {name:30s} {ws_score:3d}%  {status}")

    # Trail gap check
    county_map = analysis.get("county_trail_map", {}).get("counties", {})
    gaps = [county for county, data in county_map.items() if data.get("trail_gap") is True]
    if gaps:
        print(f"\nTrail gaps identified ({len(gaps)} counties):")
        for g in gaps:
            opp = county_map[g].get("gap_opportunity", "")
            print(f"  • {g}: {opp}")

    if score >= 80:
        region_key = analysis.get("meta", {}).get("region_key", "region")
        print(f"\n✓ Ready for Phase 4. Run:")
        print(f"  python tools/phase4_guide.py --region {region_key} --analysis {analysis_path}")
    else:
        print(f"\nComplete the analysis before proceeding to Phase 4.")

    print(f"{'='*55}\n")


def main():
    parser = argparse.ArgumentParser(description="Phase 3: Generate destination analysis template")
    parser.add_argument("--region", "-r", help="Region key (e.g. finger_lakes)")
    parser.add_argument("--setup", "-s", help="Path to Phase 1 setup brief JSON")
    parser.add_argument("--output", "-o", help="Output path (default: outputs/analysis-<region>.json)")
    parser.add_argument("--review", help="Review completeness of an existing analysis JSON")
    args = parser.parse_args()

    if args.review:
        review_analysis(args.review)
        return

    if not args.region:
        print("ERROR: --region required (or --review to score existing analysis)")
        sys.exit(1)

    region_data = load_region(args.region)
    setup_brief = load_setup_brief(args.setup)

    template = generate_template(args.region, region_data, setup_brief)

    OUTPUTS_PATH.mkdir(parents=True, exist_ok=True)
    out_path = Path(args.output) if args.output else OUTPUTS_PATH / f"analysis-{args.region}.json"

    with open(out_path, "w") as f:
        json.dump(template, f, indent=2)

    print(f"\n{'='*60}")
    print(f"PHASE 3 TEMPLATE — {region_data['display_name']}")
    print(f"{'='*60}")
    print(f"Saved to: {out_path}")
    print()
    print("CLAUDE CODE INSTRUCTIONS:")
    print("Fill in all null fields using web_search and web_fetch.")
    print("Work through each section in order:")
    print()
    print("  1. assets          — natural, cultural, culinary, craft bev")
    print("  2. infrastructure  — lodging, access, visitor services")
    print("  3. uvp             — singular identity, differentiation")
    print("  4. seasonality     — peak, shoulder, off-season, events")
    print("  5. county_trail_map— verify which counties have active trails")
    print("  6. guide_framing   — emotional hook and lede angle")
    print()
    print("When complete, run:")
    print(f"  python tools/phase3_analyze.py --review {out_path}")
    print()
    print("Then proceed to Phase 4:")
    print(f"  python tools/phase4_guide.py --region {args.region} --analysis {out_path}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
