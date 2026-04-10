from __future__ import annotations
import argparse
import csv
import sys
import unicodedata
from pathlib import Path
from openpyxl import load_workbook

def header_treatment(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = " ".join(text.split())
    return text


def clean_column(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()


def choose_set(workbook, requested_sheet: str | None):
    if requested_sheet:
        if requested_sheet not in workbook.sheetnames:
            raise ValueError(
                f"Feuille introuvable : {requested_sheet}. "
                f"Feuilles disponibles : {', '.join(workbook.sheetnames)}"
            )
        return workbook[requested_sheet]

    for name in workbook.sheetnames:
        ws = workbook[name]
        first_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
        if not first_row:
            continue
        headers = {header_treatment(v) for v in first_row if v is not None}
        if {"code", "label fr", "parent immediat"}.issubset(headers):
            return ws
            
    return workbook[workbook.sheetnames[0]]


def extract_header_map(worksheet) -> dict[str, int]:
    first_row = next(worksheet.iter_rows(min_row=1, max_row=1, values_only=True), None)
    if not first_row:
        raise ValueError("The excel file is empty")

    header_map: dict[str, int] = {}
    for idx, value in enumerate(first_row):
        norm = header_treatment(value)
        if norm:
            header_map[norm] = idx

    required = ["code", "label fr", "parent immediat"]
    missing = [col for col in required if col not in header_map]
    if missing:
        raise ValueError(
            "Column absent : "
            + ", ".join(missing)
            + f". Column detected : {', '.join(header_map.keys())}"
        )

    return header_map


def iter_useful_rows(xlsx_path: Path, sheet_name: str | None = None):
    wb = load_workbook(filename=xlsx_path, read_only=True, data_only=True)
    try:
        ws = choose_set(wb, sheet_name)
        header_map = extract_header_map(ws)

        code_idx = header_map["code"]
        label_idx = header_map["label fr"]
        parent_idx = header_map["parent immediat"]

        for row in ws.iter_rows(min_row=2, values_only=True):
            code = clean_column(row[code_idx] if code_idx < len(row) else None)
            label = clean_column(row[label_idx] if label_idx < len(row) else None)
            parent = clean_column(row[parent_idx] if parent_idx < len(row) else None)
            
            if not code:
                continue

            yield code, label, parent
    finally:
        wb.close()


def write_csv(rows, output_csv: Path) -> int:
    output_csv.parent.mkdir(parents=True, exist_ok=True)

    count = 0
    with output_csv.open("w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["code", "label", "parent"])

        for code, label, parent in rows:
            writer.writerow([code, label, parent])
            count += 1

    return count


def parse_args():
    parser = argparse.ArgumentParser(
        description="..."
    )
    parser.add_argument(
        "--input",
        required=True,
        help="...",
    )
    parser.add_argument(
        "--output-dir",
        required=True,
        help="...",
    )
    parser.add_argument(
        "--output-name",
        default="cim10_hierarchie.csv",
        help="...",
    )
    parser.add_argument(
        "--sheet",
        default=None,
        help="...",
    )
    return parser.parse_args()


def main():
    args = parse_args()

    input_path = Path(args.input).expanduser().resolve()
    output_dir = Path(args.output_dir).expanduser().resolve()
    output_csv = output_dir / args.output_name

    if not input_path.exists():
        raise FileNotFoundError(f"Not found : {input_path}")

    if input_path.suffix.lower() != ".xlsx":
        raise ValueError(f"Need to be an .xlsx : {input_path}")

    rows = iter_useful_rows(input_path, args.sheet)
    count = write_csv(rows, output_csv)

    print(f"Source : {input_path}")
    print(f"Output : {output_csv}")
    print(f"Lines : {count}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Erreur : {e}", file=sys.stderr)
        sys.exit(1)
