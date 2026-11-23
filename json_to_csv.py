"""
Build a single CSV from the per-product JSON files written by
lumens_scraper.py.

Usage:
    python json_to_csv.py [--input temp_data.json] [--output File.csv]
"""

import os, json, glob, argparse
import pandas as pd


COLUMNS = [
    "Lumnes Product Name",
    "Lumnes Product Url",
    "Lumnes Product Image Url",
    "Lumens Product Price",
    "Lumens BSR rank",
    "MFR ID",
    "Domain",
]


def load_records(input_dir: str):
    records = []
    if os.path.isfile(input_dir):
        paths = [input_dir]
    else:
        pattern = os.path.join(input_dir, "**", "*.json")
        paths = glob.glob(pattern, recursive=True)

    for path in paths:
        try:
            with open(path, encoding="utf-8") as f:
                data = json.load(f)
            items = data if isinstance(data, list) else [data]
            records.extend(
                rec for rec in items
                if isinstance(rec, dict) and "Lumnes Product Url" in rec
            )
        except Exception as e:
            print(f"[skip] {path}: {e}")
    return records


def build_csv(input_dir: str, output_file: str):
    records = load_records(input_dir)
    if not records:
        print("No product records found.")
        return

    df = pd.DataFrame(records)

    # ensure all expected columns exist, in order
    for col in COLUMNS:
        if col not in df.columns:
            df[col] = None
    df = df[COLUMNS]

    # dedupe by MFR ID, prefer rows with a price
    # df["_has_price"] = df["Lumens Product Price"].notna() & \
    #                    (df["Lumens Product Price"].astype("string") != "")
    # df = df.sort_values("_has_price", ascending=False)
    # df = df.drop_duplicates("MFR ID", keep="first").drop(columns="_has_price")
    df["Lumens Product Price"] = df["Lumens Product Price"].str.replace(",","")
    df.to_csv(output_file, index=False)
    print(f"Wrote {len(df)} rows to {output_file}")


if __name__ == "__main__":
    p = argparse.ArgumentParser()
    p.add_argument("--input",  default="temp_data.json",
                   help="JSON file or directory containing product records")
    p.add_argument("--output", default="File.csv",
                   help="Output CSV path")
    args = p.parse_args()
    build_csv(args.input, args.output)