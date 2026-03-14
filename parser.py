import pdfplumber
import pandas as pd
import re
from tqdm import tqdm
import os

category_order = [
    "TFWS","GOPEN","LOPEN",
    "GOBC","LOBC",
    "GSEBC","LSEBC",
    "GVJ","LVJ",
    "GNT","LNT",
    "GSC","LSC",
    "GST","LST"
]


def normalize_category(cat):

    if cat.startswith("DEF"):
        return None

    cat = re.sub(r'[SH]$', '', cat)

    return cat


def update_category(row, category, rank, percentile):

    r_col = f"{category} Rank"
    p_col = f"{category} Percentile"

    if r_col not in row:
        row[r_col] = rank
        row[p_col] = percentile
        return

    existing_rank = row.get(r_col)
    existing_pct = row.get(p_col)

    if rank:
        if existing_rank:
            row[r_col] = min(int(existing_rank), int(rank))
        else:
            row[r_col] = rank

    if percentile:
        if existing_pct:
            row[p_col] = max(float(existing_pct), float(percentile))
        else:
            row[p_col] = percentile


def extract_pdf_to_excel(pdf_path, output_folder):

    records = []

    college_code = None
    college_name = None
    branch_code = None
    branch_name = None

    with pdfplumber.open(pdf_path) as pdf:

        for page in tqdm(pdf.pages):

            text = page.extract_text()

            if not text:
                continue

            lines = text.split("\n")

            i = 0
            while i < len(lines):

                line = lines[i].strip()

                college_match = re.match(r'(\d{5})\s-\s(.+)', line)
                if college_match:
                    college_code = college_match.group(1)
                    college_name = college_match.group(2)

                branch_match = re.match(r'(\d{10})\s-\s(.+)', line)
                if branch_match:
                    branch_code = branch_match.group(1)
                    branch_name = branch_match.group(2)

                if line.startswith("Stage"):

                    headers = line.split()

                    rank_values = []
                    percentile_values = []

                    if i + 1 < len(lines):
                        rank_values = lines[i+1].split()

                    if i + 2 < len(lines):
                        percentile_values = re.findall(r'\((.*?)\)', lines[i+2])

                    row = {
                        "College Code": college_code,
                        "College Name": college_name,
                        "Branch Code": branch_code,
                        "Branch Name": branch_name
                    }

                    for j in range(1, len(headers)):

                        raw_category = headers[j]
                        category = normalize_category(raw_category)

                        if category is None:
                            continue

                        rank = None
                        pct = None

                        if j < len(rank_values):
                            rank = rank_values[j]

                        if j-1 < len(percentile_values):
                            pct = percentile_values[j-1]

                        update_category(row, category, rank, pct)

                    records.append(row)

                    i += 2

                i += 1

    df = pd.DataFrame(records)

    fixed_cols = [
        "College Code",
        "College Name",
        "Branch Code",
        "Branch Name"
    ]

    ordered_cols = []

    for cat in category_order:

        r = f"{cat} Rank"
        p = f"{cat} Percentile"

        if r in df.columns:
            ordered_cols.append(r)

        if p in df.columns:
            ordered_cols.append(p)

    df = df[fixed_cols + ordered_cols]

    output_file = os.path.join(output_folder, "cutoff_output.xlsx")

    df.to_excel(output_file, index=False)

    return output_file