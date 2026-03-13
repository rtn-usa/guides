#!/usr/bin/env python3
"""
Phase 4: Guide Generator
Takes a completed Phase 3 analysis JSON and a content JSON written by Claude,
then assembles the final Regional Guide as a Word document.

Usage:
  python tools/phase4_guide.py --region finger_lakes --analysis outputs/analysis-finger_lakes.json
  python tools/phase4_guide.py --content-template  # Print the content JSON template Claude fills

The content JSON (Claude writes this) contains all six guide sections as text.
This script assembles it into a properly formatted Word doc.
"""

import json
import argparse
import sys
from pathlib import Path
from datetime import date

OUTPUTS_PATH = Path(__file__).parent.parent / "outputs"
DATA_PATH = Path(__file__).parent.parent / "data" / "ny_regions.json"

try:
    from docx import Document
    from docx.shared import Pt, Inches, RGBColor, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.style import WD_STYLE_TYPE
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError:
    print("ERROR: python-docx not installed. Run: pip install python-docx --break-system-packages")
    sys.exit(1)

# ── Brand colors (Upstate palette) ────────────────────────────────────────────
COLOR_DEEP_GREEN  = RGBColor(0x1A, 0x3C, 0x2B)  # Primary brand dark green
COLOR_WARM_BROWN  = RGBColor(0x6B, 0x45, 0x26)  # Accent warm brown
COLOR_LIGHT_SAGE  = RGBColor(0xE8, 0xF0, 0xE3)  # Background tints
COLOR_DARK_TEXT   = RGBColor(0x1A, 0x1A, 0x18)  # Body text
COLOR_MEDIUM_GREY = RGBColor(0x6B, 0x6B, 0x6B)  # Subheads, captions
COLOR_CREAM       = RGBColor(0xFB, 0xF8, 0xF2)  # Pull quote bg

CONTENT_TEMPLATE = {
    "_instructions": "Fill every section below with Upstate brand voice content. Read /mnt/skills/user/upstate-brand-voice/SKILL.md before writing. Run the Upstate Test on every paragraph.",
    "region": "Finger Lakes",
    "guide_version": "1.0",
    "authored_date": str(date.today()),

    "section_1_hero": {
        "_instructions": "150 words. Mood first, facts second. No forbidden phrases. Lead with feeling — what does it FEEL like to be here?",
        "headline": None,
        "body": None
    },

    "section_2_at_a_glance": {
        "_instructions": "Factual snapshot. Drawn directly from the analysis JSON.",
        "counties": [],
        "trail_count": None,
        "active_trails": [],
        "seasonal_windows": {
            "peak": None,
            "shoulder": None,
            "off_season": None
        },
        "drive_times": {},
        "base_towns": []
    },

    "section_3_why_come": {
        "_instructions": "2-3 paragraphs. The UVP distilled. What's irreplaceable about this region? Include RTN's collaborative angle — what the trail network adds that a single destination can't.",
        "paragraphs": []
    },

    "section_4_trail_network": {
        "_instructions": "One entry per active trail. Description = 2-3 sentences in Upstate voice. Cross-trail callout = where this trail connects to others.",
        "intro": None,
        "trails": [
            {
                "trail_name": None,
                "description": None,
                "passport_available": None,
                "cross_trail_callout": None,
                "trail_url": None
            }
        ]
    },

    "section_5_featured_operators": {
        "_instructions": "Curated by category — NOT a directory dump. Trail members get priority. 3-5 operators per category, written in full Upstate listing voice. Include what makes each one worth the drive.",
        "intro": None,
        "categories": {
            "food_drink": {
                "intro": None,
                "operators": []
            },
            "culture_heritage": {
                "intro": None,
                "operators": []
            },
            "attractions": {
                "intro": None,
                "operators": []
            },
            "wellness": {
                "intro": None,
                "operators": []
            },
            "outdoor": {
                "intro": None,
                "operators": []
            },
            "agriculture": {
                "intro": None,
                "operators": []
            }
        }
    },

    "section_6_seasonal_practical": {
        "_instructions": "When to go, where to base, how to get there. Shoulder season gets its own callout box — it's the editorial priority.",
        "seasonal_guide": {
            "spring": None,
            "summer": None,
            "fall": None,
            "winter": None,
            "shoulder_season_callout": {
                "headline": None,
                "body": None
            }
        },
        "where_to_base": {
            "intro": None,
            "towns": []
        },
        "how_to_get_there": {
            "by_car": None,
            "by_train": None,
            "by_bus": None
        },
        "practical_tips": []
    }
}


def set_paragraph_color(paragraph, color):
    for run in paragraph.runs:
        run.font.color.rgb = color


def add_section_divider(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run("─" * 60)
    run.font.color.rgb = COLOR_MEDIUM_GREY
    run.font.size = Pt(8)


def add_callout_box(doc, headline, body, color=None):
    """Simulated callout box using bordered paragraph."""
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.3)
    p.paragraph_format.right_indent = Inches(0.3)
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(12)

    if headline:
        run = p.add_run(f"▸ {headline}\n")
        run.bold = True
        run.font.color.rgb = color or COLOR_DEEP_GREEN
        run.font.size = Pt(11)

    if body:
        run = p.add_run(body)
        run.font.size = Pt(10)
        run.font.color.rgb = COLOR_DARK_TEXT


def build_guide(content, analysis, output_path):
    doc = Document()

    # ── Page margins ──────────────────────────────────────────────────────────
    section = doc.sections[0]
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1.15)
    section.right_margin = Inches(1.15)

    region = content.get("region", "Unknown Region")

    # ════════════════════════════════════════════════════════════════════════
    # HEADER
    # ════════════════════════════════════════════════════════════════════════
    brand_line = doc.add_paragraph()
    brand_line.alignment = WD_ALIGN_PARAGRAPH.LEFT
    brand_run = brand_line.add_run("UPSTATE")
    brand_run.font.color.rgb = COLOR_DEEP_GREEN
    brand_run.font.bold = True
    brand_run.font.size = Pt(9)
    brand_run.font.name = "Arial"

    tagline_run = brand_line.add_run("  |  Regional Guide Series")
    tagline_run.font.color.rgb = COLOR_MEDIUM_GREY
    tagline_run.font.size = Pt(9)
    tagline_run.font.name = "Arial"

    add_section_divider(doc)

    # ── SECTION 1: Hero ──────────────────────────────────────────────────────
    hero = content.get("section_1_hero", {})
    headline = hero.get("headline") or f"Go Further {region}"
    hero_body = hero.get("body") or "[Hero narrative — Claude to write. ~150 words. Mood first.]"

    h1 = doc.add_paragraph()
    h1.alignment = WD_ALIGN_PARAGRAPH.LEFT
    h1_run = h1.add_run(headline)
    h1_run.font.size = Pt(28)
    h1_run.font.bold = True
    h1_run.font.color.rgb = COLOR_DEEP_GREEN
    h1_run.font.name = "Georgia"
    h1.paragraph_format.space_after = Pt(14)

    body_p = doc.add_paragraph()
    body_run = body_p.add_run(hero_body)
    body_run.font.size = Pt(11)
    body_run.font.color.rgb = COLOR_DARK_TEXT
    body_run.font.name = "Georgia"
    body_p.paragraph_format.space_after = Pt(20)
    body_p.paragraph_format.line_spacing = Pt(16)

    add_section_divider(doc)

    # ── SECTION 2: At a Glance ───────────────────────────────────────────────
    glance = content.get("section_2_at_a_glance", {})

    s2_head = doc.add_paragraph()
    s2_run = s2_head.add_run("Region at a Glance")
    s2_run.font.size = Pt(13)
    s2_run.font.bold = True
    s2_run.font.color.rgb = COLOR_WARM_BROWN
    s2_run.font.name = "Arial"
    s2_head.paragraph_format.space_before = Pt(16)
    s2_head.paragraph_format.space_after = Pt(8)

    # Quick-facts table
    facts = []
    if glance.get("counties"):
        facts.append(("Counties", ", ".join(glance["counties"])))
    if glance.get("trail_count"):
        facts.append(("Active trails", str(glance["trail_count"])))
    if glance.get("seasonal_windows", {}).get("peak"):
        facts.append(("Peak season", glance["seasonal_windows"]["peak"]))
    if glance.get("seasonal_windows", {}).get("shoulder"):
        facts.append(("Shoulder season", glance["seasonal_windows"]["shoulder"]))
    if glance.get("base_towns"):
        facts.append(("Best base towns", ", ".join(glance["base_towns"])))

    if glance.get("drive_times"):
        for city, time in glance["drive_times"].items():
            facts.append((f"From {city}", time))

    if facts:
        tbl = doc.add_table(rows=len(facts), cols=2)
        tbl.style = "Table Grid"
        for i, (label, value) in enumerate(facts):
            tbl.cell(i, 0).text = label
            tbl.cell(i, 1).text = value
            for cell in tbl.row_cells(i):
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.font.size = Pt(10)
                        run.font.name = "Arial"
    else:
        p = doc.add_paragraph("[At-a-glance data — Claude to complete from analysis JSON]")
        p.runs[0].font.color.rgb = COLOR_MEDIUM_GREY

    doc.add_paragraph()
    add_section_divider(doc)

    # ── SECTION 3: Why Come Here ─────────────────────────────────────────────
    s3_head = doc.add_paragraph()
    s3_run = s3_head.add_run("Why Come Here")
    s3_run.font.size = Pt(13)
    s3_run.font.bold = True
    s3_run.font.color.rgb = COLOR_WARM_BROWN
    s3_run.font.name = "Arial"
    s3_head.paragraph_format.space_before = Pt(16)
    s3_head.paragraph_format.space_after = Pt(8)

    why_paragraphs = content.get("section_3_why_come", {}).get("paragraphs", [])
    if why_paragraphs:
        for para_text in why_paragraphs:
            p = doc.add_paragraph(para_text)
            p.paragraph_format.space_after = Pt(10)
            for run in p.runs:
                run.font.size = Pt(11)
                run.font.name = "Georgia"
                run.font.color.rgb = COLOR_DARK_TEXT
    else:
        p = doc.add_paragraph("[Why come here — 2-3 paragraphs. UVP + RTN collaborative angle.]")
        p.runs[0].font.color.rgb = COLOR_MEDIUM_GREY

    add_section_divider(doc)

    # ── SECTION 4: Trail Network ─────────────────────────────────────────────
    s4_head = doc.add_paragraph()
    s4_run = s4_head.add_run("Trail Network")
    s4_run.font.size = Pt(13)
    s4_run.font.bold = True
    s4_run.font.color.rgb = COLOR_WARM_BROWN
    s4_run.font.name = "Arial"
    s4_head.paragraph_format.space_before = Pt(16)
    s4_head.paragraph_format.space_after = Pt(8)

    trail_intro = content.get("section_4_trail_network", {}).get("intro")
    if trail_intro:
        p = doc.add_paragraph(trail_intro)
        p.paragraph_format.space_after = Pt(10)
        for run in p.runs:
            run.font.size = Pt(11)
            run.font.name = "Georgia"

    trails = content.get("section_4_trail_network", {}).get("trails", [])
    for trail in trails:
        name = trail.get("trail_name", "")
        desc = trail.get("description", "")
        passport = trail.get("passport_available")
        cross = trail.get("cross_trail_callout")

        trail_head = doc.add_paragraph()
        trail_head.paragraph_format.space_before = Pt(12)
        tn_run = trail_head.add_run(name)
        tn_run.font.bold = True
        tn_run.font.size = Pt(11)
        tn_run.font.color.rgb = COLOR_DEEP_GREEN
        tn_run.font.name = "Arial"
        if passport:
            pp_run = trail_head.add_run("  •  Passport Program")
            pp_run.font.size = Pt(9)
            pp_run.font.color.rgb = COLOR_WARM_BROWN
            pp_run.font.name = "Arial"

        if desc:
            p = doc.add_paragraph(desc)
            p.paragraph_format.space_after = Pt(4)
            for run in p.runs:
                run.font.size = Pt(10)
                run.font.name = "Georgia"

        if cross:
            p = doc.add_paragraph(f"→ Cross-trail: {cross}")
            for run in p.runs:
                run.font.size = Pt(9)
                run.font.italic = True
                run.font.color.rgb = COLOR_MEDIUM_GREY

    if not trails:
        p = doc.add_paragraph("[Trail network — one entry per active trail in the region]")
        p.runs[0].font.color.rgb = COLOR_MEDIUM_GREY

    add_section_divider(doc)

    # ── SECTION 5: Featured Operators ────────────────────────────────────────
    s5_head = doc.add_paragraph()
    s5_run = s5_head.add_run("Featured Operators")
    s5_run.font.size = Pt(13)
    s5_run.font.bold = True
    s5_run.font.color.rgb = COLOR_WARM_BROWN
    s5_run.font.name = "Arial"
    s5_head.paragraph_format.space_before = Pt(16)
    s5_head.paragraph_format.space_after = Pt(8)

    op_intro = content.get("section_5_featured_operators", {}).get("intro")
    if op_intro:
        p = doc.add_paragraph(op_intro)
        for run in p.runs:
            run.font.size = Pt(11)
            run.font.name = "Georgia"

    CATEGORY_DISPLAY = {
        "food_drink":       "Food & Drink",
        "culture_heritage": "Culture & Heritage",
        "attractions":      "Attractions",
        "wellness":         "Wellness",
        "outdoor":          "Outdoor",
        "agriculture":      "Agriculture",
    }

    categories = content.get("section_5_featured_operators", {}).get("categories", {})
    for cat_key, cat_display in CATEGORY_DISPLAY.items():
        cat_data = categories.get(cat_key, {})
        operators = cat_data.get("operators", [])
        if not operators:
            continue

        cat_head = doc.add_paragraph()
        cat_head.paragraph_format.space_before = Pt(14)
        cat_run = cat_head.add_run(cat_display)
        cat_run.font.size = Pt(12)
        cat_run.font.bold = True
        cat_run.font.color.rgb = COLOR_DEEP_GREEN
        cat_run.font.name = "Arial"

        cat_intro_text = cat_data.get("intro")
        if cat_intro_text:
            p = doc.add_paragraph(cat_intro_text)
            p.paragraph_format.space_after = Pt(8)
            for run in p.runs:
                run.font.size = Pt(10)
                run.font.italic = True
                run.font.name = "Georgia"

        for op in operators:
            op_head = doc.add_paragraph()
            op_head.paragraph_format.space_before = Pt(8)
            name_run = op_head.add_run(op.get("name", "Operator Name"))
            name_run.font.bold = True
            name_run.font.size = Pt(10)
            name_run.font.name = "Arial"
            name_run.font.color.rgb = COLOR_DARK_TEXT

            if op.get("trail_member"):
                tm_run = op_head.add_run("  ★ Trail Member")
                tm_run.font.size = Pt(8)
                tm_run.font.color.rgb = COLOR_WARM_BROWN

            if op.get("location"):
                loc_run = op_head.add_run(f"  |  {op['location']}")
                loc_run.font.size = Pt(9)
                loc_run.font.color.rgb = COLOR_MEDIUM_GREY

            if op.get("description"):
                p = doc.add_paragraph(op["description"])
                p.paragraph_format.space_after = Pt(6)
                for run in p.runs:
                    run.font.size = Pt(10)
                    run.font.name = "Georgia"

    add_section_divider(doc)

    # ── SECTION 6: Seasonal & Practical ─────────────────────────────────────
    s6_head = doc.add_paragraph()
    s6_run = s6_head.add_run("When to Go & How to Get There")
    s6_run.font.size = Pt(13)
    s6_run.font.bold = True
    s6_run.font.color.rgb = COLOR_WARM_BROWN
    s6_run.font.name = "Arial"
    s6_head.paragraph_format.space_before = Pt(16)
    s6_head.paragraph_format.space_after = Pt(8)

    seasonal = content.get("section_6_seasonal_practical", {}).get("seasonal_guide", {})
    for season_name, season_key in [("Spring", "spring"), ("Summer", "summer"), ("Fall", "fall"), ("Winter", "winter")]:
        text = seasonal.get(season_key)
        if text:
            season_head = doc.add_paragraph()
            sr = season_head.add_run(f"{season_name}  ")
            sr.font.bold = True
            sr.font.size = Pt(10)
            sr.font.color.rgb = COLOR_DEEP_GREEN
            sr.font.name = "Arial"
            body_r = season_head.add_run(text)
            body_r.font.size = Pt(10)
            body_r.font.name = "Georgia"
            season_head.paragraph_format.space_after = Pt(6)

    # Shoulder season callout box
    shoulder = seasonal.get("shoulder_season_callout", {})
    if shoulder.get("headline") or shoulder.get("body"):
        add_callout_box(
            doc,
            shoulder.get("headline", "Shoulder Season"),
            shoulder.get("body", ""),
            COLOR_DEEP_GREEN
        )

    # Where to base
    base = content.get("section_6_seasonal_practical", {}).get("where_to_base", {})
    if base.get("intro") or base.get("towns"):
        base_head = doc.add_paragraph()
        base_head.paragraph_format.space_before = Pt(12)
        bh_run = base_head.add_run("Where to Base Yourself")
        bh_run.font.bold = True
        bh_run.font.size = Pt(11)
        bh_run.font.color.rgb = COLOR_DEEP_GREEN
        bh_run.font.name = "Arial"

        if base.get("intro"):
            p = doc.add_paragraph(base["intro"])
            for run in p.runs:
                run.font.size = Pt(10)
                run.font.name = "Georgia"

    # Getting there
    getting_there = content.get("section_6_seasonal_practical", {}).get("how_to_get_there", {})
    if any(getting_there.values()):
        gt_head = doc.add_paragraph()
        gt_head.paragraph_format.space_before = Pt(12)
        gh_run = gt_head.add_run("Getting There")
        gh_run.font.bold = True
        gh_run.font.size = Pt(11)
        gh_run.font.color.rgb = COLOR_DEEP_GREEN
        gh_run.font.name = "Arial"

        for mode, label in [("by_car", "By car"), ("by_train", "By train"), ("by_bus", "By bus")]:
            text = getting_there.get(mode)
            if text:
                p = doc.add_paragraph()
                lr = p.add_run(f"{label}  ")
                lr.font.bold = True
                lr.font.size = Pt(10)
                lr.font.name = "Arial"
                tr = p.add_run(text)
                tr.font.size = Pt(10)
                tr.font.name = "Georgia"

    # Footer
    doc.add_paragraph()
    add_section_divider(doc)
    footer = doc.add_paragraph()
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fr = footer.add_run(f"Upstate  |  upstate.travel  |  Part of the Rural Tourism Network  |  {date.today().strftime('%B %Y')}")
    fr.font.size = Pt(8)
    fr.font.color.rgb = COLOR_MEDIUM_GREY
    fr.font.name = "Arial"

    doc.save(output_path)
    return output_path


def main():
    parser = argparse.ArgumentParser(description="Phase 4: Assemble the Regional Guide Word doc")
    parser.add_argument("--region", "-r", help="Region key")
    parser.add_argument("--analysis", "-a", help="Path to completed Phase 3 analysis JSON")
    parser.add_argument("--content", "-c", help="Path to Claude-written content JSON")
    parser.add_argument("--output", "-o", help="Output path for the .docx")
    parser.add_argument("--content-template", action="store_true", help="Print the content JSON template")
    args = parser.parse_args()

    if args.content_template:
        print(json.dumps(CONTENT_TEMPLATE, indent=2))
        return

    if not args.region:
        print("ERROR: --region required. Use --content-template to see what Claude should write.")
        sys.exit(1)

    # Load analysis if provided
    analysis = {}
    if args.analysis and Path(args.analysis).exists():
        with open(args.analysis) as f:
            analysis = json.load(f)

    # Load content if provided, else use empty template
    content = {}
    if args.content and Path(args.content).exists():
        with open(args.content) as f:
            content = json.load(f)
    else:
        # Pre-fill what we can from analysis
        region_data = analysis.get("meta", {})
        content = dict(CONTENT_TEMPLATE)
        content["region"] = region_data.get("region", args.region.replace("_", " ").title())
        print("NOTE: No --content file provided. Building guide with placeholder text.")
        print("      Claude should fill outputs/content-<region>.json first.\n")

    OUTPUTS_PATH.mkdir(parents=True, exist_ok=True)
    region_slug = args.region
    out_path = Path(args.output) if args.output else OUTPUTS_PATH / f"guide-{region_slug}.docx"

    output_path = build_guide(content, analysis, out_path)

    print(f"\n{'='*55}")
    print(f"PHASE 4 COMPLETE — {content.get('region', region_slug)}")
    print(f"{'='*55}")
    print(f"Guide saved → {output_path}")
    print()
    print("NEXT STEPS:")
    print("  1. Review the guide in Word")
    print("  2. Add photography placeholders")
    print("  3. Share with RTN for review")
    print("  4. Adapt hero section as Upstate website copy")
    print(f"{'='*55}\n")


if __name__ == "__main__":
    main()
