#!/usr/bin/env python3
"""
Phase 4 — Guide Creation

Generates the content template and assembles the final DOCX guide.

Usage:
    python tools/phase4_guide.py --content-template > outputs/content-<region>.json
    python tools/phase4_guide.py --region <key> \\
        --analysis outputs/analysis-<region>.json \\
        --content outputs/content-<region>.json

Six sections:
1. Hero Narrative (~150 words)
2. Region at a Glance
3. Why Come Here
4. Trail Network
5. Featured Operators
6. Seasonal + Practical
"""

import argparse
import json
import os
import sys

try:
    from docx import Document
    from docx.shared import Inches, Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.section import WD_ORIENT
except ImportError:
    print("Error: python-docx is required. Install with: pip install python-docx", file=sys.stderr)
    sys.exit(1)

REGIONS_FILE = os.path.join(os.path.dirname(__file__), "..", "data", "ny_regions.json")
OUTPUTS_DIR = os.path.join(os.path.dirname(__file__), "..", "outputs")

FORBIDDEN_PHRASES = [
    "hidden gem", "charming", "quaint", "nestled", "breathtaking",
    "rustic charm", "off the beaten path", "world-class", "scenic byway",
    "there's something for everyone", "come experience", "immerse yourself",
    "a must-see", "you won't be disappointed", "explore all that upstate has to offer",
    "a true gem of new york"
]


def print_content_template():
    """Print the content JSON template Claude should fill."""
    template = {
        "section_1_hero": {
            "title": "REGION_NAME",
            "subtitle": "BRAND_TAGLINE",
            "narrative": "~150 words. Mood and feeling FIRST. No place names in the first sentence. Example register: 'The light changes somewhere around...' or 'There's a moment, usually on the second day...'"
        },
        "section_2_glance": {
            "counties": ["County1", "County2"],
            "county_count": 0,
            "trail_count": 0,
            "drive_times": {
                "from_nyc": "",
                "from_albany": "",
                "from_other": ""
            },
            "base_towns": ["Town1", "Town2"],
            "amtrak": "Service description or 'No direct service'",
            "quick_facts": ["Fact 1", "Fact 2", "Fact 3"]
        },
        "section_3_why": {
            "paragraphs": [
                "Paragraph 1: The UVP in human language",
                "Paragraph 2: What the trail network adds",
                "Paragraph 3: End with something that makes you want to look up a route"
            ]
        },
        "section_4_trails": [
            {
                "name": "Trail Name",
                "type": "wine | craft_beverage | food | outdoor | heritage | culture | scenic | agriculture",
                "description": "2-3 sentences in Upstate voice. Specific, factual, no embellishment.",
                "passport": True,
                "cross_trail_callout": "Optional: note where this trail connects with another"
            }
        ],
        "section_5_operators": {
            "food_drink": [
                {
                    "name": "Operator Name",
                    "location": "Town, County",
                    "trail_member": True,
                    "trail_name": "Trail Name or null",
                    "description": "What makes this one worth the drive. Specific, factual."
                }
            ],
            "culture_heritage": [],
            "attractions": [],
            "wellness": [],
            "outdoor": [],
            "agriculture": []
        },
        "section_6_seasonal": {
            "seasons": {
                "spring": "One sharp sentence",
                "summer": "One sharp sentence",
                "fall": "One sharp sentence",
                "winter": "One sharp sentence"
            },
            "shoulder_callout": {
                "title": "SHOULDER SEASON SPOTLIGHT",
                "months": "May-June or September-October",
                "content": "Why this is the time to come. Upstate's editorial signature."
            },
            "where_to_base": [
                {
                    "town": "Town Name",
                    "why": "Why this town works as a base"
                }
            ],
            "getting_here": {
                "by_car": "Route names and drive descriptions",
                "by_train": "Amtrak service details or 'No direct service'",
                "by_bus": "Bus options if any"
            }
        }
    }
    print(json.dumps(template, indent=2))


def load_regions():
    with open(REGIONS_FILE, "r") as f:
        return json.load(f)


def check_voice(text):
    """Check text against forbidden phrases."""
    violations = []
    text_lower = text.lower()
    for phrase in FORBIDDEN_PHRASES:
        if phrase in text_lower:
            violations.append(phrase)
    if "!" in text:
        violations.append("exclamation mark")
    return violations


def check_all_content(content):
    """Check all content sections for voice violations."""
    violations = []

    def walk(obj, path=""):
        if isinstance(obj, str):
            v = check_voice(obj)
            if v:
                violations.append((path, v))
        elif isinstance(obj, dict):
            for k, val in obj.items():
                walk(val, f"{path}.{k}" if path else k)
        elif isinstance(obj, list):
            for i, item in enumerate(obj):
                walk(item, f"{path}[{i}]")

    walk(content)
    return violations


def add_heading(doc, text, level=1):
    """Add a styled heading."""
    heading = doc.add_heading(text, level=level)
    for run in heading.runs:
        run.font.color.rgb = RGBColor(0x2E, 0x40, 0x57)
    return heading


def add_body(doc, text):
    """Add body text paragraph."""
    para = doc.add_paragraph(text)
    para.style.font.size = Pt(11)
    para.style.font.name = "Georgia"
    para.paragraph_format.space_after = Pt(8)
    para.paragraph_format.line_spacing = 1.4
    return para


def build_guide(region_key, analysis_path, content_path):
    """Assemble the final DOCX guide."""
    data = load_regions()
    if region_key not in data["regions"]:
        print(f"Error: Region '{region_key}' not found.", file=sys.stderr)
        sys.exit(1)

    region = data["regions"][region_key]

    # Load analysis and content
    with open(analysis_path, "r") as f:
        analysis = json.load(f)
    with open(content_path, "r") as f:
        content = json.load(f)

    # Voice check
    violations = check_all_content(content)
    if violations:
        print("VOICE VIOLATIONS DETECTED:", file=sys.stderr)
        for path, phrases in violations:
            print(f"  {path}: {', '.join(phrases)}", file=sys.stderr)
        print(f"\nTotal violations: {len(violations)}", file=sys.stderr)
        print("Fix these before generating the guide.", file=sys.stderr)
        sys.exit(1)

    # Build DOCX
    doc = Document()

    # Page setup
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

    # --- Section 1: Hero ---
    hero = content.get("section_1_hero", {})
    title = doc.add_heading(hero.get("title", region["name"]), level=0)
    for run in title.runs:
        run.font.color.rgb = RGBColor(0x2E, 0x40, 0x57)
        run.font.size = Pt(28)

    if hero.get("subtitle"):
        subtitle = doc.add_paragraph(hero["subtitle"])
        subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in subtitle.runs:
            run.font.size = Pt(14)
            run.font.color.rgb = RGBColor(0x6B, 0x70, 0x5C)
            run.italic = True

    doc.add_paragraph()  # spacer

    if hero.get("narrative"):
        add_body(doc, hero["narrative"])

    doc.add_page_break()

    # --- Section 2: Region at a Glance ---
    add_heading(doc, "Region at a Glance", level=1)
    glance = content.get("section_2_glance", {})

    glance_items = []
    if glance.get("counties"):
        glance_items.append(f"Counties: {', '.join(glance['counties'])} ({glance.get('county_count', len(glance['counties']))})")
    if glance.get("trail_count"):
        glance_items.append(f"Active Trails: {glance['trail_count']}")
    if glance.get("base_towns"):
        glance_items.append(f"Base Towns: {', '.join(glance['base_towns'])}")

    drive = glance.get("drive_times", {})
    for key, val in drive.items():
        if val:
            label = key.replace("from_", "From ").replace("_", " ").title()
            glance_items.append(f"{label}: {val}")

    if glance.get("amtrak"):
        glance_items.append(f"Amtrak: {glance['amtrak']}")

    for item in glance_items:
        doc.add_paragraph(item, style="List Bullet")

    if glance.get("quick_facts"):
        doc.add_paragraph()
        add_heading(doc, "Quick Facts", level=2)
        for fact in glance["quick_facts"]:
            doc.add_paragraph(fact, style="List Bullet")

    doc.add_page_break()

    # --- Section 3: Why Come Here ---
    add_heading(doc, "Why Come Here", level=1)
    why = content.get("section_3_why", {})
    for para_text in why.get("paragraphs", []):
        add_body(doc, para_text)

    doc.add_page_break()

    # --- Section 4: Trail Network ---
    add_heading(doc, "Trail Network", level=1)
    trails = content.get("section_4_trails", [])
    for trail in trails:
        trail_heading = trail.get("name", "Unknown Trail")
        if trail.get("passport"):
            trail_heading += " [Passport Available]"
        add_heading(doc, trail_heading, level=2)

        trail_type = trail.get("type", "")
        if trail_type:
            type_para = doc.add_paragraph(f"Type: {trail_type}")
            for run in type_para.runs:
                run.font.size = Pt(9)
                run.font.color.rgb = RGBColor(0x6B, 0x70, 0x5C)

        if trail.get("description"):
            add_body(doc, trail["description"])

        if trail.get("cross_trail_callout"):
            callout = doc.add_paragraph(f"Cross-trail: {trail['cross_trail_callout']}")
            for run in callout.runs:
                run.italic = True
                run.font.color.rgb = RGBColor(0x1B, 0x5E, 0x20)

    doc.add_page_break()

    # --- Section 5: Featured Operators ---
    add_heading(doc, "Featured Operators", level=1)
    operators = content.get("section_5_operators", {})
    category_labels = {
        "food_drink": "Food & Drink",
        "culture_heritage": "Culture & Heritage",
        "attractions": "Attractions",
        "wellness": "Wellness",
        "outdoor": "Outdoor",
        "agriculture": "Agriculture"
    }

    for cat_key, cat_label in category_labels.items():
        ops = operators.get(cat_key, [])
        if not ops:
            continue

        add_heading(doc, cat_label, level=2)
        for op in ops:
            name = op.get("name", "Unknown")
            location = op.get("location", "")
            trail_badge = ""
            if op.get("trail_member") and op.get("trail_name"):
                trail_badge = f" — {op['trail_name']} member"

            op_para = doc.add_paragraph()
            run_name = op_para.add_run(name)
            run_name.bold = True
            run_name.font.size = Pt(11)

            if location:
                op_para.add_run(f" | {location}")
            if trail_badge:
                run_trail = op_para.add_run(trail_badge)
                run_trail.font.color.rgb = RGBColor(0x1B, 0x5E, 0x20)
                run_trail.italic = True

            if op.get("description"):
                add_body(doc, op["description"])

    doc.add_page_break()

    # --- Section 6: Seasonal + Practical ---
    add_heading(doc, "Seasonal Guide", level=1)
    seasonal = content.get("section_6_seasonal", {})

    # Seasons
    seasons = seasonal.get("seasons", {})
    season_labels = {"spring": "Spring", "summer": "Summer", "fall": "Fall", "winter": "Winter"}
    for season_key, season_label in season_labels.items():
        text = seasons.get(season_key, "")
        if text:
            para = doc.add_paragraph()
            run = para.add_run(f"{season_label}: ")
            run.bold = True
            para.add_run(text)

    # Shoulder callout
    shoulder = seasonal.get("shoulder_callout", {})
    if shoulder.get("content"):
        doc.add_paragraph()
        callout_heading = add_heading(doc, shoulder.get("title", "SHOULDER SEASON SPOTLIGHT"), level=2)
        if shoulder.get("months"):
            months_para = doc.add_paragraph(shoulder["months"])
            for run in months_para.runs:
                run.bold = True
                run.font.color.rgb = RGBColor(0x6B, 0x70, 0x5C)
        add_body(doc, shoulder["content"])

    # Where to base
    bases = seasonal.get("where_to_base", [])
    if bases:
        doc.add_paragraph()
        add_heading(doc, "Where to Base", level=2)
        for base in bases:
            para = doc.add_paragraph()
            run = para.add_run(f"{base.get('town', '')}: ")
            run.bold = True
            para.add_run(base.get("why", ""))

    # Getting here
    getting = seasonal.get("getting_here", {})
    if any(getting.values()):
        doc.add_paragraph()
        add_heading(doc, "Getting Here", level=2)
        transport_labels = {"by_car": "By Car", "by_train": "By Train", "by_bus": "By Bus"}
        for key, label in transport_labels.items():
            text = getting.get(key, "")
            if text:
                para = doc.add_paragraph()
                run = para.add_run(f"{label}: ")
                run.bold = True
                para.add_run(text)

    # --- Footer ---
    doc.add_paragraph()
    footer = doc.add_paragraph()
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer.add_run("Upstate — Where New York goes to remember what matters")
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0x6B, 0x70, 0x5C)
    run.italic = True

    footer2 = doc.add_paragraph()
    footer2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = footer2.add_run("Powered by Tourismo | upstate.travel")
    run2.font.size = Pt(8)
    run2.font.color.rgb = RGBColor(0x99, 0x99, 0x99)

    # Save
    os.makedirs(OUTPUTS_DIR, exist_ok=True)
    output_path = os.path.join(OUTPUTS_DIR, f"guide-{region_key}.docx")
    doc.save(output_path)

    print(f"Guide assembled: {output_path}")
    print(f"Region: {region['name']}")
    print(f"Sections: 6")
    print(f"Trails documented: {len(trails)}")
    total_ops = sum(len(operators.get(cat, [])) for cat in category_labels)
    print(f"Featured operators: {total_ops}")
    print(f"Voice violations: 0")


def main():
    parser = argparse.ArgumentParser(description="Phase 4: Guide Creation")
    parser.add_argument("--content-template", action="store_true", help="Print content JSON template")
    parser.add_argument("--region", type=str, help="Region key")
    parser.add_argument("--analysis", type=str, help="Path to completed analysis JSON")
    parser.add_argument("--content", type=str, help="Path to completed content JSON")
    args = parser.parse_args()

    if args.content_template:
        print_content_template()
        return

    if not all([args.region, args.analysis, args.content]):
        parser.print_help()
        print("\nUsage examples:")
        print("  python tools/phase4_guide.py --content-template > outputs/content-finger_lakes.json")
        print("  python tools/phase4_guide.py --region finger_lakes \\")
        print("      --analysis outputs/analysis-finger_lakes.json \\")
        print("      --content outputs/content-finger_lakes.json")
        sys.exit(1)

    for path in [args.analysis, args.content]:
        if not os.path.exists(path):
            print(f"Error: File not found: {path}", file=sys.stderr)
            sys.exit(1)

    build_guide(args.region, args.analysis, args.content)


if __name__ == "__main__":
    main()
