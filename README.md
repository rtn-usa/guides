# Upstate Regional Guide System

A Claude Code toolkit for building the 9 Upstate NY regional guides.

## Setup

```bash
cd upstate-guide-system
pip install -r requirements.txt --break-system-packages
```

## Usage

See `CLAUDE.md` for the full 4-phase process.

Quick start (Finger Lakes):
```bash
python tools/phase1_setup.py --region finger_lakes
python tools/phase2_source.py --template
python tools/phase3_analyze.py --region finger_lakes
python tools/phase4_guide.py --content-template
```

## Tools

| Tool | Phase | What it does |
|---|---|---|
| `tools/phase1_setup.py` | 1 | Generates county/DMO/trail setup brief |
| `tools/phase2_source.py` | 2 | Converts Claude-scraped data into formatted XLSX |
| `tools/phase3_analyze.py` | 3 | Generates analysis template; scores completeness |
| `tools/phase4_guide.py` | 4 | Assembles final Word doc guide |

## Data

`data/ny_regions.json` — all 9 regions with counties, DMOs, known trails, and metro distances.

## Outputs

All generated files go to `outputs/`:
- `setup-<region>.json` — Phase 1 brief
- `<region>-listings.xlsx` — Phase 2 business listings
- `analysis-<region>.json` — Phase 3 destination analysis
- `content-<region>.json` — Phase 4 guide content (Claude writes this)
- `guide-<region>.docx` — Final regional guide
