"""
rate_pairs_openrouter.py

Sends each dating-profile pair to one or more models via OpenRouter, at
whatever temperature the model is actually configured with (no explicit
temperature override — see TEMPERATURE below). Each pair's questionnaire
.docx (Pair_XX_questionnaire.docx — intro, Male Profile, Female Profile,
and all 13 questions, already combined) is read directly and its full
text becomes the ONE single message sent to the model. No system prompt
is used, and no part of the wording is hardcoded in this script —
everything comes straight from the .docx files in QUESTIONNAIRES_DIR.

Every pair is rated once per model in MODELS — so total calls =
len(MODELS) x number of pairs. Each call is completely fresh and
isolated: a brand-new `messages` list with no history carried over from
any other pair or model, and each response is parsed independently.

Results (the 12 numeric scores plus the 5 ranked factors from question
13) are written directly into spreadsheets — one per model, saved in a
results/ folder — no per-pair .json or .txt files are created.

Requirements:
    pip install openai python-dotenv python-docx openpyxl --break-system-packages

Usage:
    Create a .env file in this same directory containing:
        OPENROUTER_API_KEY=sk-or-...
    Then just run:
        python rate_pairs_openrouter.py

    Edit MODELS near the top of this file to control which model(s) run.

Input:
    profiles/Pair_01_questionnaire.docx
    profiles/Pair_02_questionnaire.docx
    ...
    (one fully-assembled questionnaire per pair — nothing else needed)

Output:
    results/<model>_results.xlsx   (one file per model, one row per pair)
"""

import os
import re
import time
from pathlib import Path

import docx
from dotenv import load_dotenv
from openai import OpenAI
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

load_dotenv()  # reads a .env file in the current directory into os.environ, if present

# --- Model selection -------------------------------------------------------
# Any OpenRouter model slug works here, e.g.:
#   "openai/gpt-5.5"
#   "anthropic/claude-opus-4.8"
#   "google/gemini-2.5-pro"
# See https://openrouter.ai/models for the current list/slugs.
# Every pair is rated once per model listed here.
MODELS = [
    "openai/gpt-5.5",
    "anthropic/claude-opus-4.8",
    "google/gemini-2.5-pro",
]

# Temperature is intentionally NOT set on the API call — this leaves each
# model at whatever temperature it's actually configured with by default
# on OpenRouter. Set this to a float instead if you want one fixed
# temperature applied to every call.
TEMPERATURE = None

QUESTIONNAIRES_DIR = Path("profiles")
RESULTS_DIR = Path("results")
RESULTS_DIR.mkdir(exist_ok=True)

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


def rate_one_pair(client, prompt_text, model):
    """
    Makes a single, self-contained chat-completion call for one pair, at
    one model. A brand-new `messages` list — containing exactly one user
    message — is built here every time. Nothing is carried over from
    previous calls (across pairs OR models), and no system message is
    used at all.
    """
    kwargs = dict(
        model=model,
        messages=[
            {"role": "user", "content": prompt_text}
        ],
        max_tokens=1500,
    )
    if TEMPERATURE is not None:
        kwargs["temperature"] = TEMPERATURE

    response = client.chat.completions.create(**kwargs)
    return response.choices[0].message.content


def parse_numeric_fields(raw_text):
    """
    Best-effort parse of the numbered answers into a flat dict.
    Falls back to leaving fields blank if the model's formatting varies —
    the raw response text is included in the return value too, so
    nothing is lost even if a field fails to parse.
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

    fields["raw_response"] = raw_text
    return fields


# Column layout for the output spreadsheet: (header label, field key)
# No "Model" column — each spreadsheet is already scoped to one model.
COLUMNS = [
    ("Pair", "pair_id"),
    ("Q1 Realism (1-7)", "q1_realism"),
    ("Q2 Construction (1-3)", "q2_construction"),
    ("Q3 Social/Demo Compat", "q3_social_demo"),
    ("Q4 Social/Demo Conf.", "q4_social_demo_conf"),
    ("Q5 Psych Compat", "q5_psych"),
    ("Q6 Psych Conf.", "q6_psych_conf"),
    ("Q7 Hobbies Compat", "q7_hobbies"),
    ("Q8 Hobbies Conf.", "q8_hobbies_conf"),
    ("Q9 Expectations Compat", "q9_expectations"),
    ("Q10 Expectations Conf.", "q10_expectations_conf"),
    ("Q11 Overall Compat", "q11_overall"),
    ("Q12 Overall Conf.", "q12_overall_conf"),
    ("Q13 Factor 1", "q13_factor_1"),
    ("Q13 Factor 2", "q13_factor_2"),
    ("Q13 Factor 3", "q13_factor_3"),
    ("Q13 Factor 4", "q13_factor_4"),
    ("Q13 Factor 5", "q13_factor_5"),
]


def model_filename(model):
    """
    Turns an OpenRouter model slug into a filesystem-safe filename, e.g.
    'openai/gpt-5.5' -> 'openai_gpt-5.5_results.xlsx'.
    """
    safe = model.replace("/", "_")
    return f"{safe}_results.xlsx"


def build_workbook(rows):
    """
    Builds a formatted spreadsheet from the collected rows — one row per
    (pair, model) combination, all scores and factors as columns.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Results"

    FONT = "Arial"
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name=FONT, bold=True, color="FFFFFF", size=10)
    cell_font = Font(name=FONT, size=10)
    pair_font = Font(name=FONT, bold=True, size=10)
    thin = Side(style="thin", color="B7B7B7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")
    wrap = Alignment(wrap_text=True, vertical="top")

    for c, (label, _) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=c, value=label)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    ws.row_dimensions[1].height = 30

    for r, row in enumerate(rows, start=2):
        for c, (_, key) in enumerate(COLUMNS, start=1):
            val = row.get(key)
            cell = ws.cell(row=r, column=c, value=val)
            cell.border = border
            if key == "pair_id":
                cell.font = pair_font
                cell.alignment = center
            elif key.startswith("q13_factor"):
                cell.font = cell_font
                cell.alignment = wrap
            else:
                cell.font = cell_font
                cell.alignment = center

    ws.column_dimensions["A"].width = 10
    for c in range(2, 14):
        ws.column_dimensions[get_column_letter(c)].width = 13
    for c in range(14, 19):
        ws.column_dimensions[get_column_letter(c)].width = 34

    ws.freeze_panes = "B2"

    # Print setup: fit to page width, landscape, repeat header row
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:1"

    return wb


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
    print(f"Running models: {MODELS}")
    total_calls = len(MODELS) * len(questionnaire_files)
    print(f"Total API calls this run: {total_calls}")

    for model in MODELS:
        print(f"\n=== Model: {model} ===")

        rows = []

        for qfile in questionnaire_files:
            pair_id = qfile.stem.replace("_questionnaire", "")  # e.g. "Pair_01"

            prompt_text = extract_prompt_text(qfile)

            print(f"Rating {pair_id} ({model}) ...")
            # Fresh, isolated call every time: new pair, new model, new
            # messages list — nothing carried over from any other call.
            raw_text = rate_one_pair(client, prompt_text, model=model)

            parsed = parse_numeric_fields(raw_text)

            row = {"pair_id": pair_id, **parsed}
            rows.append(row)

            time.sleep(1)  # gentle pacing; adjust/remove based on your rate limits

        wb = build_workbook(rows)
        out_path = RESULTS_DIR / model_filename(model)
        wb.save(out_path)
        print(f"Saved {len(rows)} rows to {out_path}")

    print(f"\nDone. One spreadsheet per model saved under {RESULTS_DIR}/")


if __name__ == "__main__":
    main()