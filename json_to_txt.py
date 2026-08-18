"""
json_to_txt.py

One-off converter: turns your existing results/pairXX_result.json files
into results/pairXX_response.txt files containing just the raw model
response text. Run this once for results you already generated before
the script started writing .txt files automatically.

Usage:
    python json_to_txt.py
"""

import json
from pathlib import Path

RESULTS_DIR = Path("results")

count = 0
for json_path in sorted(RESULTS_DIR.glob("pair*_result.json")):
    data = json.loads(json_path.read_text())
    pair_id = data["pair_id"]
    raw_text = data["raw_response"]

    txt_path = RESULTS_DIR / f"{pair_id}_response.txt"
    txt_path.write_text(raw_text)
    count += 1
    print(f"Wrote {txt_path}")

print(f"\nDone. Converted {count} file(s).")