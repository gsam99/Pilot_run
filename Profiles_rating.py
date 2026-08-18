"""
rate_pairs_openrouter.py

Sends each dating-profile pair to a model via OpenRouter as a completely
fresh, isolated request — no shared conversation history, so pair N never
has access to anything about pair N-1, N+1, or any other pair.

Requirements:
    pip install openai python-dotenv --break-system-packages
    (OpenRouter is OpenAI-API-compatible, so we just point the OpenAI SDK
     at OpenRouter's base_url — no separate SDK needed.)

Usage:
    Create a .env file in this same directory containing:
        OPENROUTER_API_KEY=sk-or-...
    Then just run:
        python rate_pairs_openrouter.py

    (You can still use `export OPENROUTER_API_KEY=...` instead of a .env
    file if you prefer — the script checks the environment either way.)

Input:
    profiles/pairXX_male.txt
    profiles/pairXX_female.txt
    (one pair of files per profile pair — only the 4 paragraphs, nothing else)

Output:
    results/pairXX_result.json   (raw response + parsed fields, per pair)
    results/all_results.csv      (tidy summary table, one row per pair)
"""

import os
import re
import csv
import json
import time
from pathlib import Path

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

PROFILES_DIR = Path("profiles")
RESULTS_DIR = Path("results")
RESULTS_DIR.mkdir(exist_ok=True)

SYSTEM_PROMPT = """You will be presented with a pair of dating profiles selected from an online dating site. Evaluate the pair independently and base your judgments only on the information provided in the profiles below. Rate the pair on four compatibility dimensions followed by an overall compatibility assessment.

Questions:

1. How realistic do these two dating profiles appear? (1 = Not at all realistic, 7 = Highly realistic)

2. Which of the following best describes how these profiles were constructed:
   1. Written by LLMs without human input
   2. Written by Humans without LLM assistance
   3. Written through human-AI collaboration

3. On a scale of 0-100 (0 = No compatibility, 100 = Perfect compatibility), rate the compatibility of the pair based on their social and demographic characteristics such as age, ethnicity, religion, education, height, body type, and other relevant background characteristics.

4. How confident are you in your social and demographic rating? (0 = Not at all confident, 100 = Extremely confident)

5. On a scale of 0-100 (0 = No compatibility, 100 = Perfect compatibility), rate the compatibility of the pair based on their psychological characteristics, personality, and interpersonal style.

6. How confident are you in your psychological compatibility rating? (0 = Not at all confident, 100 = Extremely confident)?

7. On a scale of 0-100 (0 = No compatibility, 100 = Perfect compatibility), rate the compatibility of the pair based on their hobbies, lifestyles, activities, and interests.

8. How confident are you in your hobbies, lifestyle, and interests compatibility rating? (0 = Not at all confident, 100 = Extremely confident)?

9. On a scale of 0-100 (0 = No compatibility, 100 = Perfect compatibility), rate the compatibility of the pair based on their expectations from their partner or relationship.

10. How confident are you in your partner-expectation compatibility rating? (0 = Not at all confident, 100 = Extremely confident)

11. Considering the profiles as a whole, on a scale of 0-100 (0 = No compatibility, 100 = Perfect compatibility), rate the overall compatibility of the pair?

12. How confident are you in your overall compatibility rating? (0 = Not at all confident, 100 = Extremely confident)?

13. List the five most important factors that influenced your overall compatibility rating, ordered from most important to least important. For each factor, briefly indicate whether it increased or decreased compatibility.

Please return your responses in numbered order and provide one numeric value for each scale item."""


def discover_pairs():
    """Find all pairXX_male.txt / pairXX_female.txt pairs in PROFILES_DIR."""
    male_files = sorted(PROFILES_DIR.glob("pair*_male.txt"))
    pair_ids = []
    for mf in male_files:
        pair_id = mf.stem.replace("_male", "")  # e.g. "pair01"
        ff = PROFILES_DIR / f"{pair_id}_female.txt"
        if ff.exists():
            pair_ids.append(pair_id)
        else:
            print(f"WARNING: no matching female file for {mf}, skipping")
    return pair_ids


def rate_one_pair(client, male_text, female_text):
    """
    Makes a single, self-contained chat-completion call for one pair.
    A brand-new `messages` list is built here every time — nothing is
    carried over from previous calls, which is what guarantees isolation.
    Each call also opens a fresh HTTP request; OpenRouter (like the
    underlying providers) does not retain state between calls.
    """
    user_content = f"Male Profile:\n{male_text}\n\nFemale Profile:\n{female_text}"

    response = client.chat.completions.create(
        model=MODEL,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},  # static, no profile content
            {"role": "user", "content": user_content},
        ],
        max_tokens=1500,
    )

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
    # Capture everything from "13." to the end of the response, then also
    # split it into up to five separate factor lines for easier reading
    # in the CSV.
    q13_match = re.search(r"13\.(.*)", raw_text, re.DOTALL)
    q13_block = q13_match.group(1).strip() if q13_match else ""
    fields["q13_factors_raw"] = q13_block.replace("\n", " | ")

    # Best-effort split into individual factor lines. Numbered items
    # (e.g. "1. ...", "2) ...") are matched directly so any preamble text
    # before the first numbered item (like "Top five factors:") is
    # discarded rather than counted as factor #1.
    factor_matches = re.findall(
        r"\d+[\.\)]\s*(.+?)(?=\n\s*\d+[\.\)]|\Z)", q13_block, re.DOTALL
    )
    if not factor_matches:
        # Fall back to bullet points if the model didn't use numbers
        factor_matches = re.findall(r"[-•]\s*(.+?)(?=\n\s*[-•]|\Z)", q13_block, re.DOTALL)
    factor_lines = [f.strip().replace("\n", " ") for f in factor_matches if f.strip()]
    for i in range(5):
        fields[f"q13_factor_{i+1}"] = factor_lines[i] if i < len(factor_lines) else None

    return fields


def main():
    api_key = os.environ.get("OPENROUTER_API_KEY")
    if not api_key:
        raise SystemExit("Set OPENROUTER_API_KEY before running.")

    client = OpenAI(
        base_url="https://openrouter.ai/api/v1",
        api_key=api_key,
        # Optional but recommended by OpenRouter for routing/analytics attribution:
        default_headers={
            "HTTP-Referer": "https://your-site-or-project.example",  # replace or remove
            "X-Title": "Profile Pair Compatibility Rating",
        },
    )

    pair_ids = discover_pairs()
    print(f"Found {len(pair_ids)} pairs: {pair_ids}")

    csv_rows = []

    for pair_id in pair_ids:
        male_text = (PROFILES_DIR / f"{pair_id}_male.txt").read_text().strip()
        female_text = (PROFILES_DIR / f"{pair_id}_female.txt").read_text().strip()

        print(f"Rating {pair_id} ...")
        raw_text = rate_one_pair(client, male_text, female_text)

        parsed = parse_numeric_fields(raw_text)

        result = {
            "pair_id": pair_id,
            "model": MODEL,
            "raw_response": raw_text,
            "parsed": parsed,
        }
        (RESULTS_DIR / f"{pair_id}_result.json").write_text(json.dumps(result, indent=2))
        (RESULTS_DIR / f"{pair_id}_response.txt").write_text(raw_text)

        row = {"pair_id": pair_id, **parsed}
        csv_rows.append(row)

        time.sleep(1)  # gentle pacing; adjust/remove based on your rate limits

    # Tidy CSV summary
    if csv_rows:
        fieldnames = ["pair_id"] + [k for k in csv_rows[0].keys() if k != "pair_id"]
        with open(RESULTS_DIR / "all_results.csv", "w", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(csv_rows)

    print(f"\nDone. Results saved to {RESULTS_DIR}/")


if __name__ == "__main__":
    main()