"""
rate_pairs_openrouter.py

Sends each dating-profile pair to a model via OpenRouter. Each pair's
questionnaire .docx (Pair_XX_questionnaire.docx — intro, Male Profile,
Female Profile, and all 13 questions, already combined) is read directly
and its full text becomes the ONE single message sent to the model. No
system prompt is used, and no part of the wording is hardcoded in this
script — everything comes straight from the .docx files in
QUESTIONNAIRES_DIR.

Each pair still gets sent as a completely fresh API call (no message
history carried over between pairs), and each response is parsed
independently — a pair's result never depends on, or leaks into, any
other pair's result. This applies across temperatures too: every
(pair, temperature) combination is its own isolated call.

Requirements:
    pip install openai python-dotenv python-docx --break-system-packages

Usage:
    Create a .env file in this same directory containing:
        OPENROUTER_API_KEY=sk-or-...
    Then just run:
        python rate_pairs_openrouter.py

    Every pair is rated once per temperature in TEMPERATURES (near the
    top of this file) — by default: 0.2, the model's own default
    (temperature omitted), and 0.7.

Input:
    profiles/Pair_01_questionnaire.docx
    profiles/Pair_02_questionnaire.docx
    ...
    (one fully-assembled questionnaire per pair — nothing else needed)

Output (per temperature, under results/temp_<value>/ or results/temp_default/):
    PairXX_result.json   (raw response + parsed fields, per pair)
    PairXX_response.txt  (just the raw model response text)
    all_results.csv      (that temperature's results, one row per pair)

Also written directly under results/:
    all_results_all_temperatures.csv   (every pair x every temperature, combined)
"""

import os
import re
import csv
import json
import time
from pathlib import Path

import docx
from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()  # reads a .env file in the current directory into os.environ, if present

# --- Model selection -------------------------------------------------------
# Any OpenRouter model slug works here, e.g.:
#   "openai/gpt-5.5"
#   "anthropic/claude-sonnet-4.5"
#   "google/gemini-2.5-pro"
# See https://openrouter.ai/models for the current list/slugs.
MODEL = "openai/gpt-5.5"

QUESTIONNAIRES_DIR = Path("profiles")
RESULTS_DIR = Path("results")
RESULTS_DIR.mkdir(exist_ok=True)

# Each pair is rated once per temperature listed here. Use None for
# "don't set temperature at all" (the API's own default), and a float
# for an explicit value.
TEMPERATURES = [0.2, None, 0.7]

# The title heading each generated questionnaire starts with, e.g.
# "Profile Pair Compatibility Questionnaire — Pair 1". This is skipped
# when building the prompt — it's a filing label for humans, not part of
# the actual task text sent to the model.
TITLE_HEADING_STYLE = "Heading 1"


def extract_prompt_text(path):
    """
    Reads a Pair_XX_questionnaire.docx top to bottom and returns its full
    body text (intro, Male Profile, Female Profile, all 13 questions) as
    a single string, exactly as written in the document — skipping only
    the title heading line. Nothing about the wording is duplicated or
    hardcoded in this script.
    """
    d = docx.Document(path)
    lines = []
    for p in d.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        style_name = p.style.name if p.style else ""
        if style_name == TITLE_HEADING_STYLE:
            continue
        lines.append(text)
    return "\n\n".join(lines)


def discover_questionnaires():
    """Find every Pair_XX_questionnaire.docx in QUESTIONNAIRES_DIR (profiles/), in order."""
    files = sorted(QUESTIONNAIRES_DIR.glob("Pair_*_questionnaire.docx"))
    if not files:
        raise SystemExit(
            f"No questionnaire .docx files found in {QUESTIONNAIRES_DIR}/ "
            "— check the folder name and that the files are there."
        )
    return files


def rate_one_pair(client, prompt_text, temperature=None):
    """
    Makes a single, self-contained chat-completion call for one pair at
    one temperature. A brand-new `messages` list — containing exactly
    one user message — is built here every time. Nothing is carried
    over from previous calls (across pairs OR across temperatures), and
    no system message is used at all.

    temperature=None means the parameter is omitted entirely, so the
    API/model's own default temperature is used.
    """
    kwargs = dict(
        model=MODEL,
        messages=[
            {"role": "user", "content": prompt_text}
        ],
        max_tokens=1500,
    )
    if temperature is not None:
        kwargs["temperature"] = temperature

    response = client.chat.completions.create(**kwargs)
    return response.choices[0].message.content


def parse_numeric_fields(raw_text):
    """
    Best-effort parse of the numbered answers into a flat dict.
    Falls back to leaving fields blank if the model's formatting varies —
    the raw text is always saved too, so nothing is lost.
    """
    fields = {}
    patterns = {
        "q1_realism": r"1\.\s*.*?([1-7])\b",
        "q2_construction": r"2\..*?\n?\s*([1-3])\b",
        "q3_social_demo": r"3\..*?(\d{1,3})\b",
        "q4_social_demo_conf": r"4\..*?(\d{1,3})\b",
        "q5_psych": r"5\..*?(\d{1,3})\b",
        "q6_psych_conf": r"6\..*?(\d{1,3})\b",
        "q7_hobbies": r"7\..*?(\d{1,3})\b",
        "q8_hobbies_conf": r"8\..*?(\d{1,3})\b",
        "q9_expectations": r"9\..*?(\d{1,3})\b",
        "q10_expectations_conf": r"10\..*?(\d{1,3})\b",
        "q11_overall": r"11\..*?(\d{1,3})\b",
        "q12_overall_conf": r"12\..*?(\d{1,3})\b",
    }
    for key, pat in patterns.items():
        m = re.search(pat, raw_text, re.DOTALL)
        fields[key] = m.group(1) if m else None

    # Question 13 is free text (five ranked factors), not a single number.
    q13_match = re.search(r"13\.(.*)", raw_text, re.DOTALL)
    q13_block = q13_match.group(1).strip() if q13_match else ""
    fields["q13_factors_raw"] = q13_block.replace("\n", " | ")

    # Numbered items are matched directly so any preamble text before the
    # first numbered item (like "Top five factors:") is discarded rather
    # than counted as factor #1.
    factor_matches = re.findall(
        r"\d+[\.\)]\s*(.+?)(?=\n\s*\d+[\.\)]|\Z)", q13_block, re.DOTALL
    )
    if not factor_matches:
        factor_matches = re.findall(r"[-•]\s*(.+?)(?=\n\s*[-•]|\Z)", q13_block, re.DOTALL)
    factor_lines = [f.strip().replace("\n", " ") for f in factor_matches if f.strip()]
    for i in range(5):
        fields[f"q13_factor_{i+1}"] = factor_lines[i] if i < len(factor_lines) else None

    return fields


def temp_label(temperature):
    """Filesystem-friendly label for a temperature value, e.g. 0.2 -> 'temp_0.2', None -> 'temp_default'."""
    return f"temp_{temperature}" if temperature is not None else "temp_default"


def main():
    api_key = os.environ.get("OPENROUTER_API_KEY")
    if not api_key:
        raise SystemExit("Set OPENROUTER_API_KEY (e.g. in a .env file) before running.")

    client = OpenAI(
        base_url="https://openrouter.ai/api/v1",
        api_key=api_key,
        default_headers={
            "HTTP-Referer": "https://your-site-or-project.example",  # replace or remove
            "X-Title": "Profile Pair Compatibility Rating",
        },
    )

    questionnaire_files = discover_questionnaires()
    print(f"Found {len(questionnaire_files)} questionnaires: "
          f"{[f.stem for f in questionnaire_files]}")
    print(f"Running at temperatures: {TEMPERATURES}")

    combined_rows = []  # every pair x every temperature, for one all-up CSV

    for temperature in TEMPERATURES:
        label = temp_label(temperature)
        temp_dir = RESULTS_DIR / label
        temp_dir.mkdir(parents=True, exist_ok=True)

        print(f"\n=== Temperature: {temperature if temperature is not None else 'default'} ===")

        csv_rows = []

        for qfile in questionnaire_files:
            pair_id = qfile.stem.replace("_questionnaire", "")  # e.g. "Pair_01"

            prompt_text = extract_prompt_text(qfile)

            print(f"Rating {pair_id} ({label}) ...")
            # Fresh, isolated call every time: new pair, new temperature,
            # new messages list — nothing carried over from any other
            # call, at this temperature or any other.
            raw_text = rate_one_pair(client, prompt_text, temperature=temperature)

            parsed = parse_numeric_fields(raw_text)

            result = {
                "pair_id": pair_id,
                "model": MODEL,
                "temperature": temperature,
                "raw_response": raw_text,
                "parsed": parsed,
            }
            (temp_dir / f"{pair_id}_result.json").write_text(json.dumps(result, indent=2))
            (temp_dir / f"{pair_id}_response.txt").write_text(raw_text)

            row = {"pair_id": pair_id, "temperature": temperature, **parsed}
            csv_rows.append(row)
            combined_rows.append(row)

            time.sleep(1)  # gentle pacing; adjust/remove based on your rate limits

        if csv_rows:
            fieldnames = ["pair_id", "temperature"] + [
                k for k in csv_rows[0].keys() if k not in ("pair_id", "temperature")
            ]
            with open(temp_dir / "all_results.csv", "w", newline="") as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(csv_rows)

    # One combined CSV across all temperatures, for easy side-by-side comparison
    if combined_rows:
        fieldnames = ["pair_id", "temperature"] + [
            k for k in combined_rows[0].keys() if k not in ("pair_id", "temperature")
        ]
        with open(RESULTS_DIR / "all_results_all_temperatures.csv", "w", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(combined_rows)

    print(f"\nDone. Results saved under {RESULTS_DIR}/ "
          f"(one subfolder per temperature, plus all_results_all_temperatures.csv)")


if __name__ == "__main__":
    main()