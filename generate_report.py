"""
Used Cars Market Report generator.

Usage:
    python generate_report.py
    python generate_report.py --region Riyadh
    python generate_report.py --make Toyota --region Jeddah
    python generate_report.py --interactive
"""
import argparse
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------- settings
# All thresholds are assumptions. Change them here and rerun.
REFERENCE_YEAR = 2026
MAX_MILEAGE_KM = 1_000_000
MIN_MILEAGE_KM = 1_000
OLD_CAR_AGE = 5
MIN_YEAR_FOR_TREND = 2000
MIN_PRICED_PER_REGION = 30
MIN_PRICED_PER_YEAR_FILTERED = 10  # used only when a filter is active
HIGHLIGHT = "#D32F2F"
MUTED = "#A8A8A8"
LINE = "#5B7C99"


# ---------------------------------------------------------------- cleaning
def load_and_clean(path):
    df = pd.read_csv(path)
    log = {"rows_raw": len(df)}

    price_text = df["Price"].astype(str).str.replace(",", "", regex=False)
    df["Price"] = pd.to_numeric(price_text, errors="coerce")
    log["price_missing_text"] = int(df["Price"].isna().sum())

    zero = df["Price"] == 0
    log["price_zero"] = int(zero.sum())
    df.loc[zero, "Price"] = pd.NA
    df["Price"] = df["Price"].astype(float)

    too_high = df["Mileage"] > MAX_MILEAGE_KM
    log["mileage_too_high"] = int(too_high.sum())
    df.loc[too_high, "Mileage"] = pd.NA

    df["Age"] = REFERENCE_YEAR - df["Year"]
    suspicious = (df["Mileage"] < MIN_MILEAGE_KM) & (df["Age"] > OLD_CAR_AGE)
    log["mileage_suspicious"] = int(suspicious.sum())
    df["Mileage_adjusted"] = suspicious
    df.loc[suspicious, "Mileage"] = df.loc[suspicious, "Mileage"] * 1000


    before = len(df)
    df = df.drop_duplicates()
    df = df.drop(columns=["Link"], errors="ignore")
    log["duplicates_removed"] = before - len(df)
    log["rows_clean"] = len(df)
    log["rows_priced"] = int(df["Price"].notna().sum())
    return df, log


# ---------------------------------------------------------------- filtering
def _match(series, value):
    return series.astype(str).str.strip().str.lower() == value.strip().lower()


def apply_filters(df, region=None, make=None):
    """Return (filtered_df, folder_label, readable_text)."""
    label_parts, text_parts = [], []
    if region:
        mask = _match(df["Region"], region)
        if not mask.any():
            options = ", ".join(df["Region"].value_counts().index[:15])
            raise SystemExit(f"Region '{region}' not found. Examples: {options}")
        df = df[mask]
        region = df["Region"].iloc[0]
        label_parts.append(f"region_{region}")
        text_parts.append(f"region = {region}")
    if make:
        mask = _match(df["Make"], make)
        if not mask.any():
            options = ", ".join(df["Make"].value_counts().index[:15])
            raise SystemExit(f"Make '{make}' not found"
                             f"{' in this region' if region else ''}. Examples: {options}")
        df = df[mask]
        make = df["Make"].iloc[0]
        label_parts.append(f"make_{make}")
        text_parts.append(f"make = {make}")
    label = "_".join(label_parts).replace(" ", "-")
    return df, label, ", ".join(text_parts)


def ask_user(df):
    print("\nPress Enter to skip a question (= all).")
    print("Regions with most listings:", ", ".join(df["Region"].value_counts().index[:10]))
    region = input("Region: ").strip() or None
    print("Makes with most listings:", ", ".join(df["Make"].value_counts().index[:10]))
    make = input("Make: ").strip() or None
    return region, make


# ---------------------------------------------------------------- analysis
def analyse(df, min_year_priced=1):
    by_make = (df.groupby("Make")
                 .agg(listings=("Make", "size"),
                      priced=("Price", "count"),
                      median_price=("Price", "median"))
                 .sort_values("listings", ascending=False)
                 .head(10))

    modern = df[df["Year"] >= MIN_YEAR_FOR_TREND]
    by_year = (modern.groupby("Year")
                     .agg(priced=("Price", "count"),
                          median_price=("Price", "median")))
    by_year = by_year[by_year["priced"] >= min_year_priced]

    by_region = (df.groupby("Region")
                   .agg(listings=("Region", "size"),
                        priced=("Price", "count"),
                        median_price=("Price", "median")))
    reliable = (by_region[by_region["priced"] >= MIN_PRICED_PER_REGION]
                .sort_values("median_price", ascending=False))
    return by_make, by_year, reliable


# ---------------------------------------------------------------- charts
def _style(ax):
    ax.spines[["top", "right"]].set_visible(False)


def make_charts(by_make, by_year, reliable, out, skip_make=False, skip_region=False):
    paths = {}

    if not skip_make and len(by_make):
        top10 = by_make.sort_values("listings")
        leader = by_make["listings"].idxmax()
        fig, ax = plt.subplots(figsize=(9, 5))
        colors = [HIGHLIGHT if m == leader else MUTED for m in top10.index]
        bars = ax.barh(top10.index, top10["listings"], color=colors)
        ax.bar_label(bars, fmt="{:,.0f}", padding=4)
        ax.set_title(f"{leader} has the most listings")
        ax.set_xlabel("Number of listings (all, including unpriced)")
        _style(ax)
        plt.tight_layout()
        paths["make"] = out / "make_listings_chart.png"
        plt.savefig(paths["make"], dpi=150)
        plt.close(fig)

    data = by_year[by_year["priced"] > 0]
    if len(data) >= 2:
        fig, ax = plt.subplots(figsize=(9, 5))
        ax.plot(data.index, data["median_price"], marker="o", color=LINE)
        ax.set_title("Median price by model year")
        ax.set_xlabel("Model year")
        ax.set_ylabel("Median price")
        ax.yaxis.set_major_formatter(lambda x, _: f"{x/1000:.0f}K")
        _style(ax)
        plt.tight_layout()
        paths["year"] = out / "year_price_chart.png"
        plt.savefig(paths["year"], dpi=150)
        plt.close(fig)

    reg = reliable.sort_values("median_price")
    if len(reg) and not skip_region:
        top_value = reg["median_price"].max()
        fig, ax = plt.subplots(figsize=(9, 6))
        colors = [HIGHLIGHT if v == top_value else MUTED
                  for v in reg["median_price"]]
        bars = ax.barh(reg.index, reg["median_price"], color=colors)
        ax.bar_label(bars,
                     labels=[f"{v/1000:.0f}K" for v in reg["median_price"]],
                     padding=4)
        ax.set_title(f"Median price by region "
                     f"(regions with {MIN_PRICED_PER_REGION}+ priced listings)")
        ax.set_xlabel("Median price")
        ax.xaxis.set_major_formatter(lambda x, _: f"{x/1000:.0f}K")
        _style(ax)
        plt.tight_layout()
        paths["region"] = out / "region_price_chart.png"
        plt.savefig(paths["region"], dpi=150)
        plt.close(fig)
    return paths


# ---------------------------------------------------------------- Excel
def write_table(wb, sheet_name, table, index_name):
    sh = wb.create_sheet(sheet_name)
    t = table.reset_index()
    t.columns = [index_name] + list(t.columns[1:])
    fill = PatternFill("solid", fgColor="5B7C99")

    for c, name in enumerate(t.columns, start=1):
        cell = sh.cell(row=1, column=c, value=name)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center")

    for r, row in enumerate(t.itertuples(index=False), start=2):
        for c, v in enumerate(row, start=1):
            if hasattr(v, "item"):
                v = v.item()
            if v != v:
                v = None
            sh.cell(row=r, column=c, value=v)

    for c in range(1, len(t.columns) + 1):
        sh.column_dimensions[get_column_letter(c)].width = 18
    sh.freeze_panes = "A2"


def write_report(path, log, by_make, by_year, reliable, charts, n_rows, n_priced, filter_text=None):
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws["A1"] = "Used Cars Market Report"
    ws["A1"].font = Font(size=18, bold=True)
    pct = f"{n_priced / n_rows:.0%}" if n_rows else "n/a"
    ws["A2"] = f"{n_rows:,} listings analysed; {n_priced:,} ({pct}) have a stated price."
    ws["A3"] = ("Price analysis uses priced listings only. Medians are used "
                "because expensive cars skew averages.")
    if filter_text:
        ws["A4"] = f"Filter applied: {filter_text}"
        ws["A4"].font = Font(bold=True)

    row = 6
    for key in ("make", "year", "region"):
        if key in charts:
            img = Image(str(charts[key]))
            img.width, img.height = 640, 360
            ws.add_image(img, f"A{row}")
            row += 20

    if len(by_make) > 1:
        write_table(wb, "By Make", by_make, "Make")
    write_table(wb, "By Year", by_year, "Year")
    if len(reliable) > 1:
        write_table(wb, "By Region", reliable, "Region")

    notes = wb.create_sheet("Notes")
    lines = [
        "Data cleaning decisions",
        f"1. {log['price_missing_text']:,} listings had no numeric price "
        "(marked 'Negotiable') and are excluded from price analysis.",
        f"2. {log['price_zero']} listing(s) with price 0 were excluded.",
        f"3. Mileage above {MAX_MILEAGE_KM:,} km was treated as an entry error "
        f"({log['mileage_too_high']} listings); the rest of each row was kept.",
        f"4. Mileage under {MIN_MILEAGE_KM:,} km for cars older than {OLD_CAR_AGE} "
        f"years (relative to {REFERENCE_YEAR}) was assumed to be entered in thousands "
        f"of km and multiplied by 1,000 ({log['mileage_suspicious']} listings). "
        f"This is an assumption based on the value pattern.",
        f"5. {log['duplicates_removed']} rows identical across all columns, "
        "including the listing link, were removed as true duplicates. "
        "Listings with identical specs but different links were kept.",
        f"6. Year-based analysis uses cars from {MIN_YEAR_FOR_TREND} onward.",
        f"7. Regions with fewer than {MIN_PRICED_PER_REGION} priced listings "
        "are excluded from the regional chart.",
        "8. The collection date of the data is not stated; analysis uses model year.",
        "9. Findings are descriptive; they do not show causes.",
    ]
    if filter_text:
        lines.append(f"10. Report filtered by {filter_text}. Cleaning statistics above "
                     "describe the full dataset. Years with fewer than "
                     f"{MIN_PRICED_PER_YEAR_FILTERED} priced listings are left out of the year table and chart.")
    for i, line in enumerate(lines, start=1):
        notes.cell(row=i, column=1, value=line)
    notes["A1"].font = Font(bold=True, size=14)
    notes.column_dimensions["A"].width = 120

    wb.save(path)


# ---------------------------------------------------------------- main
def main():
    parser = argparse.ArgumentParser(description="Generate the used cars report.")
    parser.add_argument("--input", default="UsedCarsSA_Unclean_EN.csv")
    parser.add_argument("--output", default="output")
    parser.add_argument("--region", help="report only this region, e.g. Riyadh")
    parser.add_argument("--make", help="report only this make, e.g. Toyota")
    parser.add_argument("--interactive", action="store_true",
                        help="ask which region and make to report on")
    args = parser.parse_args()

    df, log = load_and_clean(args.input)
    print("Cleaning summary (full dataset):")
    for k, v in log.items():
        print(f"  {k}: {v:,}")

    region, make = args.region, args.make
    if args.interactive:
        region, make = ask_user(df)

    df, label, filter_text = apply_filters(df, region, make)
    filtered = bool(label)
    if filtered:
        print(f"\nFilter: {filter_text} -> {len(df):,} listings, "
              f"{int(df['Price'].notna().sum()):,} priced")

    out = Path(args.output) / label if filtered else Path(args.output)
    out.mkdir(parents=True, exist_ok=True)

    by_make, by_year, reliable = analyse(
        df, MIN_PRICED_PER_YEAR_FILTERED if filtered else 1)
    charts = make_charts(by_make, by_year, reliable, out,
                         skip_make=bool(make), skip_region=bool(region))

    df.to_csv(out / "used_cars_clean.csv", index=False)
    report = out / "used_cars_report.xlsx"
    write_report(report, log, by_make, by_year, reliable, charts,
                 n_rows=len(df), n_priced=int(df["Price"].notna().sum()),
                 filter_text=filter_text)
    print(f"Saved {report}")


if __name__ == "__main__":
    main()
